// Infrastructure for async Excel function execution.
//
// Provides:
// - Topic key building (function name + serialized args)
// - Thread pool management
// - Async worker dispatch
// - Cache integration

const std = @import("std");
const xl_imports = @import("xl_imports.zig");
const xl = xl_imports.xl;
const async_cache = @import("async_cache.zig");
const async_handler = @import("async_handler.zig");
const rtd_call = @import("rtd_call.zig");
const xl_helpers = @import("xl_helpers.zig");
const XLValue = @import("xlvalue.zig").XLValue;

const allocator = std.heap.c_allocator;

// ============================================================================
// Thread pool — lazily initialized singleton
// ============================================================================

var global_pool: ?*std.Thread.Pool = null;
var pool_mutex: std.Io.Mutex = std.Io.Mutex.init;

const default_pool_size = 4;
const io = std.Options.debug_io;

pub fn getPool() *std.Thread.Pool {
    pool_mutex.lock(io) catch {};
    defer pool_mutex.unlock(io);
    if (global_pool) |p| return p;
    const p = allocator.create(std.Thread.Pool) catch unreachable;
    p.init(.{
        .allocator = allocator,
        .n_jobs = default_pool_size,
    }) catch unreachable;
    global_pool = p;
    return p;
}

// ============================================================================
// Topic key building
// ============================================================================

/// Append a serialized argument value to the key buffer.
fn appendArg(buf: *std.ArrayList(u8), comptime T: type, arg: T) !void {
    // Handle optional types
    const type_info = @typeInfo(T);
    if (type_info == .optional) {
        if (arg) |val| {
            try appendArg(buf, type_info.optional.child, val);
        } else {
            try buf.appendSlice(allocator, "<nil>");
        }
        return;
    }

    if (T == f64) {
        try buf.print(allocator, "{d:.15}", .{arg});
    } else if (T == bool) {
        try buf.appendSlice(allocator, if (arg) "T" else "F");
    } else if (T == []const u8) {
        try buf.appendSlice(allocator, arg);
    } else if (T == [][]const f64) {
        try buf.print(allocator, "[{d}x{d}]", .{ arg.len, if (arg.len > 0) arg[0].len else 0 });
        for (arg) |row| {
            for (row) |v| {
                try buf.print(allocator, ",{d:.10}", .{v});
            }
        }
    }
}

/// Build a topic key from function name and arguments.
/// Returns a heap-allocated string owned by the caller.
pub fn buildTopicKey(comptime name: []const u8, comptime ParamTypes: []const type, args: anytype) ![]const u8 {
    var buf: std.ArrayList(u8) = .empty;
    errdefer buf.deinit(allocator);

    try buf.appendSlice(allocator, name);
    inline for (0..ParamTypes.len) |i| {
        try buf.append(allocator, '|');
        try appendArg(&buf, ParamTypes[i], args[i]);
    }

    return buf.toOwnedSlice(allocator);
}

// ============================================================================
// Arg duplication for worker threads
// ============================================================================

/// Duplicate an argument value so it can be passed to a worker thread.
/// The caller must call `freeOwnedArg` when done.
pub fn dupeArg(comptime T: type, arg: T) T {
    const type_info = @typeInfo(T);
    if (type_info == .optional) {
        if (arg) |val| {
            return dupeArg(type_info.optional.child, val);
        }
        return null;
    }

    if (T == f64 or T == bool) {
        return arg;
    } else if (T == []const u8) {
        return allocator.dupe(u8, arg) catch &.{};
    } else if (T == [][]const f64) {
        const rows = allocator.alloc([]const f64, arg.len) catch return &.{};
        for (arg, 0..) |row, i| {
            rows[i] = allocator.dupe(f64, row) catch &.{};
        }
        return rows;
    } else {
        return arg;
    }
}

/// Free a duplicated argument.
pub fn freeOwnedArg(comptime T: type, arg: T) void {
    const type_info = @typeInfo(T);
    if (type_info == .optional) {
        if (arg) |val| {
            freeOwnedArg(type_info.optional.child, val);
        }
        return;
    }

    if (T == []const u8) {
        allocator.free(arg);
    } else if (T == [][]const f64) {
        for (arg) |row| {
            allocator.free(row);
        }
        allocator.free(arg);
    }
}

// ============================================================================
// Cache result storage
// ============================================================================

fn makeErrorXloper() *xl.XLOPER12 {
    const err_ptr = allocator.create(xl.XLOPER12) catch unreachable;
    err_ptr.* = .{
        .xltype = xl.xltypeErr | xl.xlbitDLLFree,
        .val = .{ .err = xl.xlerrValue },
    };
    return err_ptr;
}

fn makeLoadingXloper() *xl.XLOPER12 {
    const err_ptr = allocator.create(xl.XLOPER12) catch unreachable;
    err_ptr.* = .{
        .xltype = xl.xltypeErr | xl.xlbitDLLFree,
        .val = .{ .err = xl.xlerrNA },
    };
    return err_ptr;
}

/// Store a "in progress" marker in the cache.
pub fn markInProgress(key: []const u8) void {
    const cache = async_cache.getGlobalCache();
    cache.put(key, .{
        .xloper = makeLoadingXloper(),
        .completed = false,
    });
}

/// Store a completed result (success or error) in the cache,
/// then notify Excel to trigger RefreshData → next recalc.
pub fn storeResult(key: []const u8, xloper: *xl.XLOPER12) void {
    const cache = async_cache.getGlobalCache();
    cache.put(key, .{
        .xloper = xloper,
        .completed = true,
    });
    notifyRtdUpdate();
}

fn notifyRtdUpdate() void {
    if (global_rtd_context) |ctx| {
        ctx.markAllDirty();
    }
    if (global_update_event) |evt| {
        initComOnThread();
        evt.updateNotify();
    }
}

/// Ensure COM is initialized on the current thread (MTA).
/// Safe to call multiple times — returns immediately if already initialized.
fn initComOnThread() void {
    if (@import("builtin").os.tag != .windows) return;
    _ = CoInitializeEx(null, 0x0); // COINIT_MULTITHREADED
}

extern "ole32" fn CoInitializeEx(reserved: ?*anyopaque, co_init: u32) callconv(.winapi) i32;

/// Windows-safe sleep. std.Thread.sleep may not work on cross-compiled targets.
pub fn sleepMs(ms: u32) void {
    const kernel32 = struct {
        extern "kernel32" fn Sleep(dwMilliseconds: u32) callconv(.winapi) void;
    };
    kernel32.Sleep(ms);
}

/// Spawn a worker thread using Windows CreateThread directly.
/// Returns the thread handle (non-null on success).
pub fn createWorkerThread(comptime ArgsType: type, comptime worker_fn: fn (*ArgsType) void, pack: *ArgsType) ?std.os.windows.HANDLE {
    if (@import("builtin").os.tag != .windows) {
        const t = std.Thread.spawn(.{}, worker_fn, .{pack}) catch return null;
        t.detach();
        return @ptrFromInt(1); // non-null sentinel
    }

    const Wrapper = struct {
        fn threadProc(param: ?*anyopaque) callconv(.winapi) u32 {
            const p: *ArgsType = @ptrCast(@alignCast(param));
            worker_fn(p);
            return 0;
        }
    };

    const handle = CreateThread(
        null, // default security
        0, // default stack size (1MB)
        Wrapper.threadProc,
        @ptrCast(pack),
        0, // run immediately
        null, // don't need thread id
    );
    if (handle) |h| {
        // Close handle — we don't join, thread runs to completion
        _ = CloseHandle(h);
        return h;
    }
    return null;
}

extern "kernel32" fn CreateThread(
    security: ?*anyopaque,
    stack_size: usize,
    start: *const fn (?*anyopaque) callconv(.winapi) u32,
    param: ?*anyopaque,
    flags: u32,
    thread_id: ?*u32,
) callconv(.winapi) ?std.os.windows.HANDLE;

extern "kernel32" fn CloseHandle(handle: std.os.windows.HANDLE) callconv(.winapi) i32;

// Set when the async RTD server starts.
pub var global_update_event: ?*@import("rtd.zig").IRTDUpdateEvent = null;
pub var global_rtd_context: ?*@import("rtd.zig").RtdContext = null;

// ============================================================================
// AsyncContext — passed to async functions that want intermediate values
// ============================================================================

/// Context passed to async functions that yield intermediate values.
///
/// Usage:
///   fn myFunc(x: f64, ctx: *AsyncContext) !f64 {
///       ctx.yield(.{ .double = x * 0.5 });   // intermediate update
///       doExpensiveWork();
///       ctx.yield(.{ .string = "almost done" });
///       doMoreWork();
///       return x * 2.0;  // final value
///   }
pub const AsyncContext = struct {
    key: []const u8,

    /// Send an intermediate (non-final) value to the cell.
    /// The cell updates immediately but stays subscribed to RTD.
    pub fn yield(self: *AsyncContext, value: AsyncValue) void {
        const xloper = asyncValueToXloper(value);
        const cache = async_cache.getGlobalCache();
        cache.put(self.key, .{
            .xloper = xloper,
            .completed = false,
        });
        notifyRtdUpdate();
    }
};

/// Values that can be yielded from an async function.
pub const AsyncValue = union(enum) {
    int: i32,
    double: f64,
    string: []const u8, // UTF-8 — framework converts
    boolean: bool,
};

fn asyncValueToXloper(val: AsyncValue) *xl.XLOPER12 {
    const ptr = allocator.create(xl.XLOPER12) catch unreachable;
    switch (val) {
        .int => |v| {
            ptr.* = .{
                .xltype = xl.xltypeInt | xl.xlbitDLLFree,
                .val = .{ .w = v },
            };
        },
        .double => |v| {
            ptr.* = .{
                .xltype = xl.xltypeNum | xl.xlbitDLLFree,
                .val = .{ .num = v },
            };
        },
        .boolean => |v| {
            ptr.* = .{
                .xltype = xl.xltypeBool | xl.xlbitDLLFree,
                .val = .{ .xbool = if (v) 1 else 0 },
            };
        },
        .string => |v| {
            const xv = XLValue.fromUtf8String(allocator, v) catch {
                ptr.* = .{
                    .xltype = xl.xltypeErr | xl.xlbitDLLFree,
                    .val = .{ .err = xl.xlerrValue },
                };
                return ptr;
            };
            ptr.* = xv.m_val;
            ptr.xltype |= xl.xlbitDLLFree;
        },
    }
    return ptr;
}

// ============================================================================
// Async subscribe helper
// ============================================================================

/// Subscribe to the async RTD server with a given topic key.
/// Sets the pending connect key so the handler can map topic_id → key.
pub fn rtdSubscribe(key: []const u8) !*xl.XLOPER12 {
    async_handler.pending_connect_key = key;
    return rtd_call.subscribe("zigxll.async", &.{key});
}

/// Clone an XLOPER12 from the cache for returning to Excel.
/// The clone gets xlbitDLLFree so xlAutoFree12 handles cleanup.
pub fn cloneXloper(src: *xl.XLOPER12) *xl.XLOPER12 {
    const base_type = src.xltype & ~@as(u32, xl.xlbitDLLFree | xl.xlbitXLFree);
    const dst = allocator.create(xl.XLOPER12) catch return makeErrorXloper();

    switch (base_type) {
        xl.xltypeNum, xl.xltypeBool, xl.xltypeInt, xl.xltypeErr => {
            dst.* = src.*;
            dst.xltype = base_type | xl.xlbitDLLFree;
        },
        xl.xltypeStr => {
            if (src.val.str) |str_ptr| {
                const len: usize = @intCast(str_ptr[0]);
                const total = len + 2; // length prefix + chars + null
                const new_buf = allocator.alloc(u16, total) catch {
                    allocator.destroy(dst);
                    return makeErrorXloper();
                };
                @memcpy(new_buf, str_ptr[0..total]);
                dst.* = .{
                    .xltype = xl.xltypeStr | xl.xlbitDLLFree,
                    .val = .{ .str = new_buf.ptr },
                };
            } else {
                dst.* = .{
                    .xltype = xl.xltypeErr | xl.xlbitDLLFree,
                    .val = .{ .err = xl.xlerrValue },
                };
            }
        },
        else => {
            // For complex types (multi, etc.) — just return error for now
            dst.* = .{
                .xltype = xl.xltypeErr | xl.xlbitDLLFree,
                .val = .{ .err = xl.xlerrValue },
            };
        },
    }
    return dst;
}
