// Lua RTD function registration — declares a Lua function that subscribes to an RTD server.
//
// The Lua function returns (prog_id, topic1, topic2, ...) via multi-return.
// The framework calls rtd_call.subscribeDynamic with those values.
// Always registered as non-thread-safe (xlfRtd must run on Excel's main thread).

const std = @import("std");
const xl_imports = @import("xl_imports.zig");
const xl = xl_imports.xl;
const XLValue = @import("xlvalue.zig").XLValue;
const lua = @import("lua.zig");
const rtd_call = @import("rtd_call.zig");
const xl_helpers = @import("xl_helpers.zig");

const allocator = std.heap.c_allocator;

const ParamMeta = @import("excel_function.zig").ParamMeta;
const LuaParam = @import("lua_function.zig").LuaParam;

fn sanitizeExportName(comptime len: usize, comptime input: *const [len]u8) *const [len]u8 {
    comptime {
        var result: [len]u8 = input.*;
        for (&result) |*ch| {
            if (ch.* == '.') ch.* = '_';
        }
        const final = result;
        return &final;
    }
}

pub fn LuaRtdFunction(comptime meta: anytype) type {
    const name = meta.name;
    const id_slice: []const u8 = if (@hasField(@TypeOf(meta), "id")) meta.id else name;
    const lua_name: [:0]const u8 = id_slice[0..id_slice.len :0];
    const description: []const u8 = if (@hasField(@TypeOf(meta), "description")) meta.description else "";
    const category: []const u8 = if (@hasField(@TypeOf(meta), "category")) meta.category else "Lua";
    const help_url: ?[]const u8 = if (@hasField(@TypeOf(meta), "help_url")) meta.help_url else null;
    const lua_params: []const LuaParam = if (@hasField(@TypeOf(meta), "params")) meta.params else &.{};

    const param_count = lua_params.len;

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

    // Always non-thread-safe (xlfRtd must be called on Excel's main thread)
    const type_string = comptime blk: {
        var result: []const u8 = "Q";
        for (0..param_count) |_| {
            result = result ++ "Q";
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
        pub const excel_thread_safe = false;
        pub const excel_is_async = false;
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

        /// Call Lua function, collect multi-return (prog_id, topics...), subscribe via RTD.
        fn callOnState(L: *lua.lua_State, args: [param_count]*xl.XLOPER12) *xl.XLOPER12 {
            _ = lua.lua_getglobal(L, lua_name.ptr);
            if (lua.lua_type(L, -1) != lua.LUA_TFUNCTION) {
                lua.lua_pop(L, 1);
                xl_helpers.debugLog("LuaRtdFunction '" ++ name ++ "': Lua function '" ++ lua_name ++ "' not found");
                return makeErrorValue();
            }

            // Push args
            inline for (0..param_count) |i| {
                if (!pushArg(L, args[i], i)) {
                    lua.lua_pop(L, @as(c_int, @intCast(i + 1)));
                    xl_helpers.debugLog("LuaRtdFunction '" ++ name ++ "': failed to convert argument " ++ std.fmt.comptimePrint("{d}", .{i + 1}));
                    return makeErrorValue();
                }
            }

            // Call with LUA_MULTRET to get all return values
            const c_api = @import("lua_c.zig").c;
            if (c_api.lua_pcallk(L, @intCast(param_count), c_api.LUA_MULTRET, 0, 0, null) != lua.LUA_OK) {
                var err_len: usize = 0;
                const err_ptr = lua.lua_tolstring(L, -1, &err_len);
                if (err_ptr) |p| {
                    xl_helpers.debugLogRuntime(p[0..err_len]);
                } else {
                    xl_helpers.debugLog("LuaRtdFunction '" ++ name ++ "': unknown Lua error");
                }
                lua.lua_pop(L, 1);
                return makeErrorValue();
            }

            // Stack now has: [prog_id, topic1, topic2, ...]
            const nresults = lua.lua_gettop(L);
            if (nresults < 1) {
                xl_helpers.debugLog("LuaRtdFunction '" ++ name ++ "': expected at least prog_id return value");
                return makeErrorValue();
            }

            // First return value (stack position 1) is prog_id
            if (lua.lua_type(L, 1) != lua.LUA_TSTRING) {
                lua.lua_settop(L, 0);
                xl_helpers.debugLog("LuaRtdFunction '" ++ name ++ "': first return value must be a string (prog_id)");
                return makeErrorValue();
            }

            var prog_id_len: usize = 0;
            const prog_id_ptr = lua.lua_tolstring(L, 1, &prog_id_len) orelse {
                lua.lua_settop(L, 0);
                return makeErrorValue();
            };
            const prog_id = prog_id_ptr[0..prog_id_len];

            // Remaining return values are topic strings
            const max_topics = 28;
            var topics_buf: [max_topics][]const u8 = undefined;
            var topic_count: usize = 0;

            var idx: c_int = 2;
            while (idx <= nresults and topic_count < max_topics) : (idx += 1) {
                if (lua.lua_type(L, idx) != lua.LUA_TSTRING) {
                    lua.lua_settop(L, 0);
                    xl_helpers.debugLog("LuaRtdFunction '" ++ name ++ "': topic return values must be strings");
                    return makeErrorValue();
                }
                var tlen: usize = 0;
                const tptr = lua.lua_tolstring(L, idx, &tlen) orelse continue;
                topics_buf[topic_count] = tptr[0..tlen];
                topic_count += 1;
            }

            // Subscribe via RTD
            const result = rtd_call.subscribeDynamic(prog_id, topics_buf[0..topic_count]) catch {
                lua.lua_settop(L, 0);
                xl_helpers.debugLog("LuaRtdFunction '" ++ name ++ "': rtd_call.subscribeDynamic failed");
                return makeErrorValue();
            };

            lua.lua_settop(L, 0);
            return result;
        }

        /// Always locks the main state (non-thread-safe).
        fn callLua(args: [param_count]*xl.XLOPER12) *xl.XLOPER12 {
            lua.lockMain();
            defer lua.unlockMain();
            const L = lua.getState() orelse {
                xl_helpers.debugLog("LuaRtdFunction '" ++ name ++ "': Lua state not initialized");
                return makeErrorValue();
            };
            return callOnState(L, args);
        }

        // Generated C-callable impl — same arity switch as LuaFunction
        const Impl = switch (param_count) {
            0 => struct {
                fn impl() callconv(.c) *xl.XLOPER12 {
                    return callLua(.{});
                }
            },
            1 => struct {
                fn impl(a1: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    return callLua(.{a1});
                }
            },
            2 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    return callLua(.{ a1, a2 });
                }
            },
            3 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    return callLua(.{ a1, a2, a3 });
                }
            },
            4 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    return callLua(.{ a1, a2, a3, a4 });
                }
            },
            5 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    return callLua(.{ a1, a2, a3, a4, a5 });
                }
            },
            6 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12, a6: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    return callLua(.{ a1, a2, a3, a4, a5, a6 });
                }
            },
            7 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12, a6: *xl.XLOPER12, a7: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    return callLua(.{ a1, a2, a3, a4, a5, a6, a7 });
                }
            },
            8 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12, a6: *xl.XLOPER12, a7: *xl.XLOPER12, a8: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
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
