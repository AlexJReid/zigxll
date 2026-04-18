// Lua function registration helper — like ExcelFunction but the implementation lives in Lua
const std = @import("std");
const xl_imports = @import("xl_imports.zig");
const xl = xl_imports.xl;
const XLValue = @import("xlvalue.zig").XLValue;
const lua = @import("lua.zig");

const xl_helpers = @import("xl_helpers.zig");

const allocator = std.heap.c_allocator;

const ParamMeta = @import("excel_function.zig").ParamMeta;

// Async support
const async_cache = @import("async_cache.zig");
const async_infra = @import("async_infra.zig");
const async_handler = @import("async_handler.zig");

/// Parameter type for Lua functions (declared statically since there's no Zig function to infer from)
pub const LuaParamType = enum {
    number,
    string,
    boolean,
};

/// Parameter declaration for a Lua-backed Excel function
pub const LuaParam = struct {
    name: []const u8,
    type: LuaParamType = .number,
    description: ?[]const u8 = null,
};

fn sanitizeExportName(comptime len: usize, comptime input: *const [len]u8) *const [len]u8 {
    comptime {
        var result: [len]u8 = input.*;
        for (&result) |*c| {
            if (c.* == '.') c.* = '_';
        }
        const final = result;
        return &final;
    }
}

pub fn LuaFunction(comptime meta: anytype) type {
    const name = meta.name;
    const id_slice: []const u8 = if (@hasField(@TypeOf(meta), "id")) meta.id else name;
    const lua_name: [:0]const u8 = id_slice[0..id_slice.len :0];
    const description: []const u8 = if (@hasField(@TypeOf(meta), "description")) meta.description else "";
    const category: []const u8 = if (@hasField(@TypeOf(meta), "category")) meta.category else "Lua";
    const help_url: ?[]const u8 = if (@hasField(@TypeOf(meta), "help_url")) meta.help_url else null;
    const lua_params: []const LuaParam = if (@hasField(@TypeOf(meta), "params")) meta.params else &.{};
    const is_async = if (@hasField(@TypeOf(meta), "is_async")) meta.is_async else if (@hasField(@TypeOf(meta), "async")) meta.@"async" else false;
    const thread_safe = if (is_async) false else if (@hasField(@TypeOf(meta), "thread_safe")) meta.thread_safe else true;

    const param_count = lua_params.len;

    // Build ParamMeta array for Excel registration
    const params_meta = comptime blk: {
        var result: [param_count]ParamMeta = undefined;
        for (0..param_count) |i| {
            result[i] = .{
                .name = lua_params[i].name,
                .description = lua_params[i].description,
            };
        }
        break :blk result;
    };

    // Generate Excel type string at comptime
    const type_string = comptime blk: {
        var result: []const u8 = "Q";
        for (0..param_count) |_| {
            result = result ++ "Q";
        }
        if (thread_safe) {
            result = result ++ "$";
        }
        break :blk result;
    };

    const export_name = comptime blk: {
        break :blk sanitizeExportName(name.len, name) ++ "_impl";
    };

    return struct {
        pub const excel_name = name;
        pub const excel_description = description;
        pub const excel_category = category;
        pub const excel_help_url = help_url;
        pub const excel_params = &params_meta;
        pub const excel_param_count = param_count;
        pub const excel_type_string = type_string;
        pub const excel_thread_safe = thread_safe;
        pub const excel_is_async = is_async;
        pub const is_excel_function = true;
        pub const excel_export_name = export_name;

        fn makeErrorValue() *xl.XLOPER12 {
            const err_ptr = allocator.create(xl.XLOPER12) catch unreachable;
            err_ptr.* = .{
                .xltype = xl.xltypeErr | xl.xlbitDLLFree,
                .val = .{ .err = xl.xlerrValue },
            };
            return err_ptr;
        }

        fn heapXloper(xloper: xl.XLOPER12) *xl.XLOPER12 {
            const ptr = allocator.create(xl.XLOPER12) catch return makeErrorValue();
            ptr.* = xloper;
            ptr.xltype |= xl.xlbitDLLFree;
            return ptr;
        }

        /// Push an XLOPER12 arg onto the Lua stack, converting based on declared param type
        fn pushArg(L: *lua.lua_State, xloper: *xl.XLOPER12, comptime param_idx: usize) bool {
            const val = XLValue.fromXLOPER12(allocator, xloper.*, false);
            const param_type = lua_params[param_idx].type;

            switch (param_type) {
                .number => {
                    const num = val.as_double() catch return false;
                    lua.lua_pushnumber(L, num);
                },
                .string => {
                    const str = val.as_utf8str() catch return false;
                    defer allocator.free(str);
                    lua.lua_pushlstring(L, str.ptr, str.len);
                },
                .boolean => {
                    const b = val.as_bool() catch return false;
                    lua.lua_pushboolean(L, if (b) 1 else 0);
                },
            }
            return true;
        }

        /// Pull the Lua return value off the stack and wrap as XLOPER12
        fn pullResult(L: *lua.lua_State) *xl.XLOPER12 {
            const lua_type = lua.lua_type(L, -1);
            switch (lua_type) {
                lua.LUA_TNUMBER => {
                    const num = lua.lua_tonumber(L, -1);
                    lua.lua_pop(L, 1);
                    return heapXloper(.{
                        .xltype = xl.xltypeNum,
                        .val = .{ .num = num },
                    });
                },
                lua.LUA_TSTRING => {
                    var len: usize = 0;
                    const ptr = lua.lua_tolstring(L, -1, &len);
                    if (ptr == null) {
                        lua.lua_pop(L, 1);
                        return makeErrorValue();
                    }
                    const str = ptr.?[0..len];
                    const xlval = XLValue.fromUtf8String(allocator, str) catch {
                        lua.lua_pop(L, 1);
                        return makeErrorValue();
                    };
                    lua.lua_pop(L, 1);
                    return heapXloper(xlval.m_val);
                },
                lua.LUA_TBOOLEAN => {
                    const b = lua.lua_toboolean(L, -1);
                    lua.lua_pop(L, 1);
                    return heapXloper(.{
                        .xltype = xl.xltypeBool,
                        .val = .{ .xbool = if (b != 0) 1 else 0 },
                    });
                },
                lua.LUA_TNIL => {
                    lua.lua_pop(L, 1);
                    return heapXloper(.{
                        .xltype = xl.xltypeNil,
                        .val = undefined,
                    });
                },
                else => {
                    lua.lua_pop(L, 1);
                    return makeErrorValue();
                },
            }
        }

        /// Execute a Lua call on a given state: push function + args, pcall, pull result.
        fn callOnState(L: *lua.lua_State, args: [param_count]*xl.XLOPER12) *xl.XLOPER12 {
            // Get the function from globals
            _ = lua.lua_getglobal(L, lua_name.ptr);
            if (lua.lua_type(L, -1) != lua.LUA_TFUNCTION) {
                lua.lua_pop(L, 1);
                xl_helpers.debugLog("LuaFunction '" ++ name ++ "': Lua function '" ++ lua_name ++ "' not found");
                return makeErrorValue();
            }

            // Push args
            inline for (0..param_count) |i| {
                if (!pushArg(L, args[i], i)) {
                    lua.lua_pop(L, @as(c_int, @intCast(i + 1))); // pop function + pushed args
                    xl_helpers.debugLog("LuaFunction '" ++ name ++ "': failed to convert argument " ++ std.fmt.comptimePrint("{d}", .{i + 1}));
                    return makeErrorValue();
                }
            }

            // Call
            if (lua.lua_pcall(L, @intCast(param_count), 1, 0) != lua.LUA_OK) {
                var err_len: usize = 0;
                const err_ptr = lua.lua_tolstring(L, -1, &err_len);
                if (err_ptr) |p| {
                    xl_helpers.debugLogRuntime(p[0..err_len]);
                } else {
                    xl_helpers.debugLog("LuaFunction '" ++ name ++ "': unknown Lua error");
                }
                lua.lua_pop(L, 1);
                return makeErrorValue();
            }

            return pullResult(L);
        }

        /// Sync call. Thread-safe functions acquire any pool state via CAS;
        /// non-thread-safe functions lock the main state (slot 0).
        fn callLua(args: [param_count]*xl.XLOPER12) *xl.XLOPER12 {
            if (thread_safe) {
                const L = lua.acquireState() orelse {
                    xl_helpers.debugLog("LuaFunction '" ++ name ++ "': no Lua state available");
                    return makeErrorValue();
                };
                defer lua.releaseState(L);
                return callOnState(L, args);
            } else {
                lua.lockMain();
                defer lua.unlockMain();
                const L = lua.getState() orelse {
                    xl_helpers.debugLog("LuaFunction '" ++ name ++ "': Lua state not initialized");
                    return makeErrorValue();
                };
                return callOnState(L, args);
            }
        }

        // ================================================================
        // Async support: topic key, worker, cache logic
        // ================================================================

        /// Argument pack duplicated for the async worker thread.
        const LuaAsyncArgs = struct {
            key: []const u8,
            /// Heap-allocated copies of XLOPER12 values (worker owns them).
            xloper_copies: [param_count]xl.XLOPER12,
        };

        /// Build a topic key from function name + serialized Lua args.
        fn buildLuaTopicKey(args: [param_count]*xl.XLOPER12) ![]const u8 {
            var buf: std.ArrayList(u8) = .empty;
            errdefer buf.deinit(allocator);

            try buf.appendSlice(allocator, name);
            inline for (0..param_count) |i| {
                try buf.append(allocator, '|');
                const val = XLValue.fromXLOPER12(allocator, args[i].*, false);
                switch (lua_params[i].type) {
                    .number => {
                        const num = val.as_double() catch 0;
                        try buf.print(allocator, "{d:.15}", .{num});
                    },
                    .string => {
                        const str = val.as_utf8str() catch "";
                        defer if (str.len > 0) allocator.free(str);
                        try buf.appendSlice(allocator, str);
                    },
                    .boolean => {
                        const b = val.as_bool() catch false;
                        try buf.appendSlice(allocator, if (b) "T" else "F");
                    },
                }
            }
            return buf.toOwnedSlice(allocator);
        }

        /// Duplicate XLOPER12 values for the worker thread.
        fn dupeXlopers(args: [param_count]*xl.XLOPER12) [param_count]xl.XLOPER12 {
            var copies: [param_count]xl.XLOPER12 = undefined;
            inline for (0..param_count) |i| {
                const src = args[i];
                copies[i] = src.*;
                // Deep-copy strings
                if ((src.xltype & xl.xltypeStr) != 0) {
                    if (src.val.str) |str_ptr| {
                        const len: usize = @intCast(str_ptr[0]);
                        const total = len + 2;
                        if (allocator.alloc(u16, total)) |new_buf| {
                            @memcpy(new_buf, str_ptr[0..total]);
                            copies[i].val = .{ .str = new_buf.ptr };
                        } else |_| {
                            copies[i].xltype = xl.xltypeErr;
                            copies[i].val = .{ .err = xl.xlerrValue };
                        }
                    }
                }
            }
            return copies;
        }

        /// Free duplicated XLOPER12 string data.
        fn freeXloperCopies(copies: *[param_count]xl.XLOPER12) void {
            inline for (0..param_count) |i| {
                if ((copies[i].xltype & xl.xltypeStr) != 0) {
                    if (copies[i].val.str) |str_ptr| {
                        const len: usize = @intCast(str_ptr[0]);
                        allocator.free(str_ptr[0 .. len + 2]);
                    }
                }
            }
        }

        /// Worker function: acquires a Lua state, runs the function, stores result.
        fn asyncWorker(pack: *LuaAsyncArgs) void {
            defer {
                freeXloperCopies(&pack.xloper_copies);
                allocator.free(pack.key);
                allocator.destroy(pack);
            }

            const L = lua.acquireState() orelse {
                xl_helpers.debugLog("LuaFunction '" ++ name ++ "': no Lua state available for async");
                async_infra.storeResult(pack.key, makeErrorValue());
                return;
            };
            defer lua.releaseState(L);

            // Build pointer array from our copies
            var ptrs: [param_count]*xl.XLOPER12 = undefined;
            inline for (0..param_count) |i| {
                ptrs[i] = &pack.xloper_copies[i];
            }

            const result = callOnState(L, ptrs);
            async_infra.storeResult(pack.key, result);
        }

        /// Async impl: check cache → return cached / spawn worker + subscribe RTD.
        fn asyncImpl(args: [param_count]*xl.XLOPER12) *xl.XLOPER12 {
            const key = buildLuaTopicKey(args) catch return makeErrorValue();
            defer allocator.free(key);

            const cache = async_cache.getGlobalCache();

            // Cache hit?
            if (cache.get(key)) |entry| {
                if (entry.completed) {
                    return async_infra.cloneXloper(entry.xloper);
                }
                return async_infra.rtdSubscribe(key) catch return makeErrorValue();
            }

            // Cache miss — spawn async work
            const pack = allocator.create(LuaAsyncArgs) catch return makeErrorValue();
            const worker_key = allocator.dupe(u8, key) catch {
                allocator.destroy(pack);
                return makeErrorValue();
            };
            pack.* = .{
                .key = worker_key,
                .xloper_copies = dupeXlopers(args),
            };

            // Mark in-progress
            async_infra.markInProgress(key);

            // Subscribe via RTD first (must happen before spawning worker)
            const rtd_result = async_infra.rtdSubscribe(key) catch {
                freeXloperCopies(&pack.xloper_copies);
                allocator.free(pack.key);
                allocator.destroy(pack);
                return makeErrorValue();
            };

            const handle = async_infra.createWorkerThread(LuaAsyncArgs, asyncWorker, pack);
            if (handle == null) {
                freeXloperCopies(&pack.xloper_copies);
                allocator.free(pack.key);
                allocator.destroy(pack);
            }

            return rtd_result;
        }

        // Generated C-callable impl — same arity switch as ExcelFunction
        // Dispatches to async path when is_async is true.
        const Impl = switch (param_count) {
            0 => struct {
                fn impl() callconv(.c) *xl.XLOPER12 {
                    if (is_async) return asyncImpl(.{});
                    return callLua(.{});
                }
            },
            1 => struct {
                fn impl(a1: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    if (is_async) return asyncImpl(.{a1});
                    return callLua(.{a1});
                }
            },
            2 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    if (is_async) return asyncImpl(.{ a1, a2 });
                    return callLua(.{ a1, a2 });
                }
            },
            3 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    if (is_async) return asyncImpl(.{ a1, a2, a3 });
                    return callLua(.{ a1, a2, a3 });
                }
            },
            4 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    if (is_async) return asyncImpl(.{ a1, a2, a3, a4 });
                    return callLua(.{ a1, a2, a3, a4 });
                }
            },
            5 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    if (is_async) return asyncImpl(.{ a1, a2, a3, a4, a5 });
                    return callLua(.{ a1, a2, a3, a4, a5 });
                }
            },
            6 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12, a6: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    if (is_async) return asyncImpl(.{ a1, a2, a3, a4, a5, a6 });
                    return callLua(.{ a1, a2, a3, a4, a5, a6 });
                }
            },
            7 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12, a6: *xl.XLOPER12, a7: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    if (is_async) return asyncImpl(.{ a1, a2, a3, a4, a5, a6, a7 });
                    return callLua(.{ a1, a2, a3, a4, a5, a6, a7 });
                }
            },
            8 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12, a6: *xl.XLOPER12, a7: *xl.XLOPER12, a8: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    if (is_async) return asyncImpl(.{ a1, a2, a3, a4, a5, a6, a7, a8 });
                    return callLua(.{ a1, a2, a3, a4, a5, a6, a7, a8 });
                }
            },
            else => @compileError("Unsupported parameter count (max 8)"),
        };

        pub const impl = Impl.impl;

        comptime {
            @export(&Impl.impl, .{ .name = export_name });
        }
    };
}
