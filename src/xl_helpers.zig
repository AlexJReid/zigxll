const std = @import("std");
const builtin = @import("builtin");
const xl_imports = @import("xl_imports.zig");
const xl = xl_imports.xl;

// Platform-specific debug output
const outputDebug = if (builtin.os.tag == .windows)
    struct {
        extern "kernel32" fn OutputDebugStringA(lpOutputString: [*:0]const u8) callconv(.winapi) void;
    }.OutputDebugStringA
else
    struct {
        fn log(msg: [*:0]const u8) void {
            std.debug.print("{s}", .{msg});
        }
    }.log;

/// Simple debug log (no formatting)
pub inline fn debugLog(comptime msg: []const u8) void {
    outputDebug(msg ++ "\n");
}

/// Debug log with formatting
/// Usage: debugLogFmt("Error: {s}", .{@errorName(err)});
pub fn debugLogFmt(comptime fmt: []const u8, args: anytype) void {
    var buf: [512]u8 = undefined;
    const msg = std.fmt.bufPrintZ(&buf, fmt ++ "\n", args) catch {
        outputDebug("debugLogFmt: buffer too small\n");
        return;
    };
    outputDebug(msg.ptr);
}

/// Debug log with a runtime string (for Lua errors, etc.)
pub fn debugLogRuntime(msg: []const u8) void {
    var buf: [512]u8 = undefined;
    const len = @min(msg.len, buf.len - 1);
    @memcpy(buf[0..len], msg[0..len]);
    buf[len] = 0;
    outputDebug(@ptrCast(&buf));
}

pub inline fn xlFree(oper: *xl.XLOPER12) void {
    _ = xl.Excel12f(xl.xlFree, null, 1, oper);
}

fn xloperBaseType(xltype: @TypeOf(@as(xl.XLOPER12, undefined).xltype)) @TypeOf(@as(xl.XLOPER12, undefined).xltype) {
    return xltype & 0xFFF;
}

pub fn freeDllOwnedPayload(allocator: std.mem.Allocator, oper: *xl.XLOPER12) void {
    switch (xloperBaseType(oper.xltype)) {
        xl.xltypeStr => {
            if (oper.val.str) |str_ptr| {
                const len = @as(usize, @intCast(str_ptr[0]));
                allocator.free(str_ptr[0 .. len + 2]);
            }
        },
        xl.xltypeMulti => {
            const rows = @as(usize, @intCast(oper.val.array.rows));
            const cols = @as(usize, @intCast(oper.val.array.columns));
            const cells = oper.val.array.lparray[0 .. rows * cols];
            for (cells) |*cell| {
                freeDllOwnedPayload(allocator, cell);
            }
            allocator.free(cells);
        },
        else => {},
    }
}

pub fn destroyDllOwnedXloper(allocator: std.mem.Allocator, oper: *xl.XLOPER12) void {
    if ((oper.xltype & xl.xlbitDLLFree) == 0) return;
    freeDllOwnedPayload(allocator, oper);
    allocator.destroy(oper);
}
