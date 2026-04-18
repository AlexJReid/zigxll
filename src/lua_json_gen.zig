// Build-time code generator: reads a JSON file describing Lua-backed Excel functions
// and emits a Zig source file with LuaFunction declarations.
//
// This allows users to define Lua functions without writing any Zig.

const std = @import("std");

pub const LuaJsonError = error{
    ParseFailed,
    MissingName,
    MissingId,
    TooManyParams,
    InvalidParamType,
};

/// Generate Zig source from a JSON function definitions file.
/// Returns the generated source as a string owned by `allocator`.
pub fn generate(allocator: std.mem.Allocator, json_bytes: []const u8) ![]const u8 {
    var writer = std.Io.Writer.Allocating.init(allocator);
    errdefer writer.deinit();
    const w = &writer.writer;

    try w.writeAll(
        \\// Auto-generated from lua_functions.json — do not edit
        \\const xll = @import("xll");
        \\const LuaFunction = xll.LuaFunction;
        \\const LuaParam = xll.LuaParam;
        \\
        \\
    );

    const parsed = std.json.parseFromSlice(std.json.Value, allocator, json_bytes, .{}) catch
        return LuaJsonError.ParseFailed;
    defer parsed.deinit();

    const root = parsed.value;
    const functions = switch (root) {
        .array => |arr| arr.items,
        .object => |obj| blk: {
            const val = obj.get("functions") orelse return LuaJsonError.ParseFailed;
            break :blk switch (val) {
                .array => |arr| arr.items,
                else => return LuaJsonError.ParseFailed,
            };
        },
        else => return LuaJsonError.ParseFailed,
    };

    for (functions) |func_val| {
        const func = switch (func_val) {
            .object => |obj| obj,
            else => return LuaJsonError.ParseFailed,
        };

        const name = getStr(func, "name") orelse return LuaJsonError.MissingName;
        const id = getStr(func, "id") orelse return LuaJsonError.MissingId;
        const description = getStr(func, "description");
        const category = getStr(func, "category");
        const help_url = getStr(func, "help_url");
        const is_async = getBool(func, "async") orelse false;

        // Derive a valid Zig identifier from the id
        try w.writeAll("pub const ");
        try writeIdentifier(w, id);
        try w.writeAll(" = LuaFunction(.{\n");

        try w.print("    .name = \"{s}\",\n", .{name});
        try w.print("    .id = \"{s}\",\n", .{id});
        if (description) |d| try w.print("    .description = \"{s}\",\n", .{d});
        if (category) |c| try w.print("    .category = \"{s}\",\n", .{c});
        if (help_url) |u| try w.print("    .help_url = \"{s}\",\n", .{u});
        if (is_async) try w.writeAll("    .is_async = true,\n");

        // Parameters
        const params_val = func.get("parameters");
        if (params_val) |pv| {
            const params = switch (pv) {
                .array => |arr| arr.items,
                else => return LuaJsonError.ParseFailed,
            };
            if (params.len > 8) return LuaJsonError.TooManyParams;
            if (params.len > 0) {
                try w.writeAll("    .params = &[_]LuaParam{\n");
                for (params) |param_val| {
                    const param = switch (param_val) {
                        .object => |obj| obj,
                        else => return LuaJsonError.ParseFailed,
                    };
                    const pname = getStr(param, "name") orelse return LuaJsonError.ParseFailed;
                    const ptype = getStr(param, "type");
                    const pdesc = getStr(param, "description");

                    try w.print("        .{{ .name = \"{s}\"", .{pname});
                    if (ptype) |t| {
                        if (std.mem.eql(u8, t, "string")) {
                            try w.writeAll(", .type = .string");
                        } else if (std.mem.eql(u8, t, "boolean")) {
                            try w.writeAll(", .type = .boolean");
                        } else if (!std.mem.eql(u8, t, "number")) {
                            return LuaJsonError.InvalidParamType;
                        }
                    }
                    if (pdesc) |d| try w.print(", .description = \"{s}\"", .{d});
                    try w.writeAll(" },\n");
                }
                try w.writeAll("    },\n");
            }
        }

        try w.writeAll("});\n\n");
    }

    return writer.toOwnedSlice();
}

fn getStr(obj: std.json.ObjectMap, key: []const u8) ?[]const u8 {
    const val = obj.get(key) orelse return null;
    return switch (val) {
        .string => |s| s,
        else => null,
    };
}

fn getBool(obj: std.json.ObjectMap, key: []const u8) ?bool {
    const val = obj.get(key) orelse return null;
    return switch (val) {
        .bool => |b| b,
        else => null,
    };
}

fn writeIdentifier(w: anytype, name: []const u8) !void {
    for (name) |c| {
        if (c == '.' or c == '-' or c == ' ') {
            try w.writeByte('_');
        } else {
            try w.writeByte(c);
        }
    }
}

// ============================================================================
// Tests
// ============================================================================

test "generate basic functions" {
    const json =
        \\[
        \\  {
        \\    "name": "Lua.Add",
        \\    "id": "add",
        \\    "description": "Add two numbers",
        \\    "category": "Math",
        \\    "parameters": [
        \\      { "name": "x", "description": "First number" },
        \\      { "name": "y", "description": "Second number" }
        \\    ]
        \\  },
        \\  {
        \\    "name": "Lua.Greet",
        \\    "id": "greet",
        \\    "description": "Greet someone",
        \\    "parameters": [
        \\      { "name": "name", "type": "string", "description": "Name to greet" }
        \\    ]
        \\  }
        \\]
    ;

    const src = try generate(std.testing.allocator, json);
    defer std.testing.allocator.free(src);

    try std.testing.expect(std.mem.indexOf(u8, src, "pub const add = LuaFunction") != null);
    try std.testing.expect(std.mem.indexOf(u8, src, ".name = \"Lua.Add\"") != null);
    try std.testing.expect(std.mem.indexOf(u8, src, ".type = .string") != null);
    try std.testing.expect(std.mem.indexOf(u8, src, ".category = \"Math\"") != null);
}

test "generate async function" {
    const json =
        \\{ "functions": [
        \\  {
        \\    "name": "Lua.Slow",
        \\    "id": "slow_calc",
        \\    "async": true,
        \\    "parameters": [{ "name": "x" }]
        \\  }
        \\]}
    ;

    const src = try generate(std.testing.allocator, json);
    defer std.testing.allocator.free(src);

    try std.testing.expect(std.mem.indexOf(u8, src, ".is_async = true") != null);
    try std.testing.expect(std.mem.indexOf(u8, src, "pub const slow_calc = LuaFunction") != null);
}

test "generate no params" {
    const json =
        \\[{ "name": "Lua.Pi", "id": "pi" }]
    ;

    const src = try generate(std.testing.allocator, json);
    defer std.testing.allocator.free(src);

    try std.testing.expect(std.mem.indexOf(u8, src, "pub const pi = LuaFunction") != null);
    // Should NOT have .params
    try std.testing.expect(std.mem.indexOf(u8, src, ".params") == null);
}

test "generate help_url" {
    const json =
        \\[{ "name": "Lua.Add", "id": "add", "help_url": "https://example.com/help" }]
    ;

    const src = try generate(std.testing.allocator, json);
    defer std.testing.allocator.free(src);

    try std.testing.expect(std.mem.indexOf(u8, src, ".help_url = \"https://example.com/help\"") != null);
}

test "reject invalid param type" {
    const json =
        \\[{ "name": "F", "id": "f", "parameters": [{ "name": "x", "type": "matrix" }] }]
    ;

    const result = generate(std.testing.allocator, json);
    try std.testing.expectError(LuaJsonError.InvalidParamType, result);
}

test "reject too many params" {
    const json =
        \\[{ "name": "F", "id": "f", "parameters": [
        \\  {"name":"a"},{"name":"b"},{"name":"c"},{"name":"d"},
        \\  {"name":"e"},{"name":"f"},{"name":"g"},{"name":"h"},{"name":"i"}
        \\]}]
    ;

    const result = generate(std.testing.allocator, json);
    try std.testing.expectError(LuaJsonError.TooManyParams, result);
}
