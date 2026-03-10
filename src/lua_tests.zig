const std = @import("std");
const lua = @import("lua.zig");

test "init and deinit" {
    try lua.init();
    defer lua.deinit();

    try std.testing.expect(lua.getState() != null);
}

test "init is idempotent" {
    try lua.init();
    defer lua.deinit();

    const state1 = lua.getState();
    try lua.init(); // second init should be a no-op
    const state2 = lua.getState();
    try std.testing.expectEqual(state1, state2);
}

test "load and call a simple function" {
    try lua.init();
    defer lua.deinit();

    const L = lua.getState().?;

    try lua.loadScript(
        \\function add(x, y)
        \\    return x + y
        \\end
    , "test_add");

    // Verify the function exists
    _ = lua.lua_getglobal(L, "add");
    try std.testing.expectEqual(lua.LUA_TFUNCTION, lua.lua_type(L, -1));
    lua.lua_pop(L, 1);

    // Call it
    _ = lua.lua_getglobal(L, "add");
    lua.lua_pushnumber(L, 3.0);
    lua.lua_pushnumber(L, 4.0);
    try std.testing.expectEqual(lua.LUA_OK, lua.lua_pcall(L, 2, 1, 0));

    const result = lua.lua_tonumber(L, -1);
    try std.testing.expectEqual(@as(f64, 7.0), result);
    lua.lua_pop(L, 1);
}

test "load and call a string function" {
    try lua.init();
    defer lua.deinit();

    const L = lua.getState().?;

    try lua.loadScript(
        \\function greet(name)
        \\    return "hello " .. name
        \\end
    , "test_greet");

    _ = lua.lua_getglobal(L, "greet");
    lua.lua_pushlstring(L, "world", 5);
    try std.testing.expectEqual(lua.LUA_OK, lua.lua_pcall(L, 1, 1, 0));

    var len: usize = 0;
    const ptr = lua.lua_tolstring(L, -1, &len);
    try std.testing.expect(ptr != null);
    try std.testing.expectEqualStrings("hello world", ptr.?[0..len]);
    lua.lua_pop(L, 1);
}

test "load and call a boolean function" {
    try lua.init();
    defer lua.deinit();

    const L = lua.getState().?;

    try lua.loadScript(
        \\function is_positive(x)
        \\    return x > 0
        \\end
    , "test_bool");

    _ = lua.lua_getglobal(L, "is_positive");
    lua.lua_pushnumber(L, 5.0);
    try std.testing.expectEqual(lua.LUA_OK, lua.lua_pcall(L, 1, 1, 0));
    try std.testing.expectEqual(@as(c_int, 1), lua.lua_toboolean(L, -1));
    lua.lua_pop(L, 1);

    _ = lua.lua_getglobal(L, "is_positive");
    lua.lua_pushnumber(L, -3.0);
    try std.testing.expectEqual(lua.LUA_OK, lua.lua_pcall(L, 1, 1, 0));
    try std.testing.expectEqual(@as(c_int, 0), lua.lua_toboolean(L, -1));
    lua.lua_pop(L, 1);
}

test "multiple scripts share global state" {
    try lua.init();
    defer lua.deinit();

    const L = lua.getState().?;

    try lua.loadScript("PI = 3.14159", "test_constants");
    try lua.loadScript(
        \\function circle_area(r)
        \\    return PI * r * r
        \\end
    , "test_circle");

    _ = lua.lua_getglobal(L, "circle_area");
    lua.lua_pushnumber(L, 1.0);
    try std.testing.expectEqual(lua.LUA_OK, lua.lua_pcall(L, 1, 1, 0));

    const result = lua.lua_tonumber(L, -1);
    try std.testing.expectApproxEqRel(3.14159, result, 0.0001);
    lua.lua_pop(L, 1);
}

test "lua runtime error is captured" {
    try lua.init();
    defer lua.deinit();

    const L = lua.getState().?;

    try lua.loadScript(
        \\function bad_func()
        \\    error("something went wrong")
        \\end
    , "test_error");

    _ = lua.lua_getglobal(L, "bad_func");
    const status = lua.lua_pcall(L, 0, 1, 0);
    try std.testing.expect(status != lua.LUA_OK);

    var len: usize = 0;
    const ptr = lua.lua_tolstring(L, -1, &len);
    try std.testing.expect(ptr != null);
    const err_msg = ptr.?[0..len];
    try std.testing.expect(std.mem.indexOf(u8, err_msg, "something went wrong") != null);
    lua.lua_pop(L, 1);
}

test "syntax error in script is reported" {
    try lua.init();
    defer lua.deinit();

    const result = lua.loadScript("function bad(", "test_syntax");
    try std.testing.expectError(error.LoadFailed, result);
}

test "multiple functions in one script" {
    try lua.init();
    defer lua.deinit();

    const L = lua.getState().?;

    try lua.loadScript(
        \\function foo()
        \\    return 1
        \\end
        \\function bar()
        \\    return 2
        \\end
    , "test_multi");

    _ = lua.lua_getglobal(L, "foo");
    try std.testing.expectEqual(lua.LUA_TFUNCTION, lua.lua_type(L, -1));
    lua.lua_pop(L, 1);

    _ = lua.lua_getglobal(L, "bar");
    try std.testing.expectEqual(lua.LUA_TFUNCTION, lua.lua_type(L, -1));
    lua.lua_pop(L, 1);
}

test "sandbox removes dofile" {
    try lua.init();
    defer lua.deinit();

    const L = lua.getState().?;
    _ = lua.lua_getglobal(L, "dofile");
    try std.testing.expectEqual(lua.LUA_TNIL, lua.lua_type(L, -1));
    lua.lua_pop(L, 1);
}

test "sandbox removes loadfile" {
    try lua.init();
    defer lua.deinit();

    const L = lua.getState().?;
    _ = lua.lua_getglobal(L, "loadfile");
    try std.testing.expectEqual(lua.LUA_TNIL, lua.lua_type(L, -1));
    lua.lua_pop(L, 1);
}

test "sandbox removes load" {
    try lua.init();
    defer lua.deinit();

    const L = lua.getState().?;
    _ = lua.lua_getglobal(L, "load");
    try std.testing.expectEqual(lua.LUA_TNIL, lua.lua_type(L, -1));
    lua.lua_pop(L, 1);
}

test "sandbox removes require" {
    try lua.init();
    defer lua.deinit();

    const L = lua.getState().?;
    _ = lua.lua_getglobal(L, "require");
    try std.testing.expectEqual(lua.LUA_TNIL, lua.lua_type(L, -1));
    lua.lua_pop(L, 1);
}

test "sandbox removes io" {
    try lua.init();
    defer lua.deinit();

    const L = lua.getState().?;
    _ = lua.lua_getglobal(L, "io");
    try std.testing.expectEqual(lua.LUA_TNIL, lua.lua_type(L, -1));
    lua.lua_pop(L, 1);
}

test "sandbox removes os.execute but keeps os.time" {
    try lua.init();
    defer lua.deinit();

    const L = lua.getState().?;

    // os.execute should be nil
    _ = lua.lua_getglobal(L, "os");
    _ = lua.lua_getfield(L, -1, "execute");
    try std.testing.expectEqual(lua.LUA_TNIL, lua.lua_type(L, -1));
    lua.lua_pop(L, 2);

    // os.time should still work
    _ = lua.lua_getglobal(L, "os");
    _ = lua.lua_getfield(L, -1, "time");
    try std.testing.expectEqual(lua.LUA_TFUNCTION, lua.lua_type(L, -1));
    lua.lua_pop(L, 2);
}

test "stdlib is available" {
    try lua.init();
    defer lua.deinit();

    const L = lua.getState().?;

    try lua.loadScript(
        \\function use_math(x)
        \\    return math.sqrt(x)
        \\end
    , "test_stdlib");

    _ = lua.lua_getglobal(L, "use_math");
    lua.lua_pushnumber(L, 16.0);
    try std.testing.expectEqual(lua.LUA_OK, lua.lua_pcall(L, 1, 1, 0));

    const result = lua.lua_tonumber(L, -1);
    try std.testing.expectEqual(@as(f64, 4.0), result);
    lua.lua_pop(L, 1);
}
