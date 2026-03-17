// Native builtins registered on the xllify Lua global table.
// These provide fast JSON and regex operations that are impractical in pure Lua.
//
// Registered functions:
//   xllify.json_parse(str) -> table | nil, error
//   xllify.json_stringify(value) -> string
//   xllify.regex_match(text, pattern, occurrence?) -> string
//   xllify.regex_replace(text, pattern, replacement) -> string

const std = @import("std");
const c = @import("lua_c.zig").c;

const allocator = std.heap.c_allocator;

// ============================================================================
// JSON
// ============================================================================

/// xllify.json_parse(str) -> table | nil, error
fn luaJsonParse(L: ?*c.lua_State) callconv(.c) c_int {
    const state = L orelse return 0;
    var len: usize = 0;
    const ptr = c.lua_tolstring(state, 1, &len) orelse {
        c.lua_pushnil(state);
        _ = c.lua_pushlstring(state, "expected string", 15);
        return 2;
    };
    const input = ptr[0..len];

    const parsed = std.json.parseFromSlice(std.json.Value, allocator, input, .{}) catch {
        c.lua_pushnil(state);
        _ = c.lua_pushlstring(state, "invalid JSON", 12);
        return 2;
    };
    defer parsed.deinit();

    pushJsonValue(state, parsed.value);
    return 1;
}

/// Push a std.json.Value onto the Lua stack as a native Lua value.
fn pushJsonValue(L: *c.lua_State, value: std.json.Value) void {
    switch (value) {
        .null => c.lua_pushnil(L),
        .bool => |b| c.lua_pushboolean(L, if (b) 1 else 0),
        .integer => |n| c.lua_pushinteger(L, @intCast(n)),
        .float => |n| c.lua_pushnumber(L, n),
        .string => |s| _ = c.lua_pushlstring(L, s.ptr, s.len),
        .number_string => |s| {
            // Try to parse as number, fall back to string
            const n = std.fmt.parseFloat(f64, s) catch {
                _ = c.lua_pushlstring(L, s.ptr, s.len);
                return;
            };
            c.lua_pushnumber(L, n);
        },
        .array => |arr| {
            c.lua_createtable(L, @intCast(arr.items.len), 0);
            for (arr.items, 1..) |item, i| {
                pushJsonValue(L, item);
                c.lua_rawseti(L, -2, @intCast(i));
            }
        },
        .object => |obj| {
            c.lua_createtable(L, 0, @intCast(obj.count()));
            var it = obj.iterator();
            while (it.next()) |entry| {
                _ = c.lua_pushlstring(L, entry.key_ptr.*.ptr, entry.key_ptr.*.len);
                pushJsonValue(L, entry.value_ptr.*);
                c.lua_settable(L, -3);
            }
        },
    }
}

/// xllify.json_stringify(value) -> string
fn luaJsonStringify(L: ?*c.lua_State) callconv(.c) c_int {
    const state = L orelse return 0;
    var buf: std.ArrayList(u8) = .{};
    defer buf.deinit(allocator);

    luaValueToJson(state, 1, buf.writer(allocator), 0) catch {
        _ = c.lua_pushlstring(state, "null", 4);
        return 1;
    };

    _ = c.lua_pushlstring(state, buf.items.ptr, buf.items.len);
    return 1;
}

/// Serialise the Lua value at `idx` to JSON.
fn luaValueToJson(L: *c.lua_State, idx: c_int, writer: anytype, depth: u32) !void {
    if (depth > 50) return error.TooDeep;

    const abs_idx = if (idx < 0) c.lua_gettop(L) + idx + 1 else idx;
    const t = c.lua_type(L, abs_idx);

    switch (t) {
        c.LUA_TNIL => try writer.writeAll("null"),
        c.LUA_TBOOLEAN => {
            if (c.lua_toboolean(L, abs_idx) != 0)
                try writer.writeAll("true")
            else
                try writer.writeAll("false");
        },
        c.LUA_TNUMBER => {
            if (c.lua_isinteger(L, abs_idx) != 0) {
                const n = c.lua_tointegerx(L, abs_idx, null);
                try std.fmt.format(writer, "{d}", .{n});
            } else {
                const n = c.lua_tonumberx(L, abs_idx, null);
                if (std.math.isNan(n) or std.math.isInf(n)) {
                    try writer.writeAll("null");
                } else {
                    try std.fmt.format(writer, "{d}", .{n});
                }
            }
        },
        c.LUA_TSTRING => {
            var slen: usize = 0;
            const sptr = c.lua_tolstring(L, abs_idx, &slen) orelse {
                try writer.writeAll("null");
                return;
            };
            try writeJsonString(writer, sptr[0..slen]);
        },
        c.LUA_TTABLE => {
            // Detect array vs object: if key 1 exists, treat as array
            _ = c.lua_rawgeti(L, abs_idx, 1);
            const is_array = c.lua_type(L, -1) != c.LUA_TNIL;
            c.lua_settop(L, c.lua_gettop(L) - 1);

            if (is_array) {
                try writer.writeByte('[');
                const len = c.luaL_len(L, abs_idx);
                var i: c.lua_Integer = 1;
                while (i <= len) : (i += 1) {
                    if (i > 1) try writer.writeByte(',');
                    _ = c.lua_rawgeti(L, abs_idx, i);
                    try luaValueToJson(L, -1, writer, depth + 1);
                    c.lua_settop(L, c.lua_gettop(L) - 1);
                }
                try writer.writeByte(']');
            } else {
                try writer.writeByte('{');
                var first = true;
                c.lua_pushnil(L);
                while (c.lua_next(L, abs_idx) != 0) {
                    // Only string keys in JSON objects
                    if (c.lua_type(L, -2) == c.LUA_TSTRING) {
                        if (!first) try writer.writeByte(',');
                        first = false;
                        var klen: usize = 0;
                        const kptr = c.lua_tolstring(L, -2, &klen) orelse {
                            c.lua_settop(L, c.lua_gettop(L) - 1);
                            continue;
                        };
                        try writeJsonString(writer, kptr[0..klen]);
                        try writer.writeByte(':');
                        try luaValueToJson(L, -1, writer, depth + 1);
                    }
                    c.lua_settop(L, c.lua_gettop(L) - 1);
                }
                try writer.writeByte('}');
            }
        },
        else => try writer.writeAll("null"),
    }
}

fn writeJsonString(writer: anytype, s: []const u8) !void {
    try writer.writeByte('"');
    for (s) |ch| {
        switch (ch) {
            '"' => try writer.writeAll("\\\""),
            '\\' => try writer.writeAll("\\\\"),
            '\n' => try writer.writeAll("\\n"),
            '\r' => try writer.writeAll("\\r"),
            '\t' => try writer.writeAll("\\t"),
            else => {
                if (ch < 0x20) {
                    try std.fmt.format(writer, "\\u{x:0>4}", .{ch});
                } else {
                    try writer.writeByte(ch);
                }
            },
        }
    }
    try writer.writeByte('"');
}

// ============================================================================
// Registration
// ============================================================================

/// Register all xllify builtins onto an existing `xllify` table on the Lua stack.
/// Call this AFTER the boot script has created `xllify = {}`.
pub fn register(L: *c.lua_State) void {
    // Get the xllify global table
    _ = c.lua_getglobal(L, "xllify");
    if (c.lua_type(L, -1) != c.LUA_TTABLE) {
        // xllify doesn't exist yet — create it
        c.lua_settop(L, c.lua_gettop(L) - 1);
        c.lua_createtable(L, 0, 2);
        c.lua_pushvalue(L, -1);
        c.lua_setglobal(L, "xllify");
    }

    // Register functions on the xllify table
    c.lua_pushcclosure(L, luaJsonParse, 0);
    c.lua_setfield(L, -2, "json_parse");

    c.lua_pushcclosure(L, luaJsonStringify, 0);
    c.lua_setfield(L, -2, "json_stringify");

    c.lua_settop(L, c.lua_gettop(L) - 1); // pop xllify table
}
