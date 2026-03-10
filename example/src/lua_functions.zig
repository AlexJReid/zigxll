// Lua-backed Excel functions
const xll = @import("xll");
const LuaFunction = xll.LuaFunction;
const LuaParam = xll.LuaParam;

pub const lua_add = LuaFunction(.{
    .name = "Lua.Add",
    .lua_name = "add",
    .description = "Add two numbers (Lua)",
    .category = "Lua Functions",
    .params = &[_]LuaParam{
        .{ .name = "x", .description = "First number" },
        .{ .name = "y", .description = "Second number" },
    },
});

pub const lua_greet = LuaFunction(.{
    .name = "Lua.Greet",
    .lua_name = "greet",
    .description = "Greet someone by name (Lua)",
    .category = "Lua Functions",
    .params = &[_]LuaParam{
        .{ .name = "name", .type = .string, .description = "Name to greet" },
    },
});

pub const lua_hypotenuse = LuaFunction(.{
    .name = "Lua.Hypotenuse",
    .lua_name = "hypotenuse",
    .description = "Calculate hypotenuse (Lua)",
    .category = "Lua Functions",
    .params = &[_]LuaParam{
        .{ .name = "a", .description = "Side a" },
        .{ .name = "b", .description = "Side b" },
    },
});

pub const lua_is_even = LuaFunction(.{
    .name = "Lua.IsEven",
    .lua_name = "is_even",
    .description = "Check if a number is even (Lua)",
    .category = "Lua Functions",
    .params = &[_]LuaParam{
        .{ .name = "n", .description = "Number to check" },
    },
});

pub const lua_factorial = LuaFunction(.{
    .name = "Lua.Factorial",
    .lua_name = "factorial",
    .description = "Calculate factorial (Lua)",
    .category = "Lua Functions",
    .params = &[_]LuaParam{
        .{ .name = "n", .description = "Number" },
    },
});

pub const lua_fib = LuaFunction(.{
    .name = "Lua.Fib",
    .lua_name = "fib",
    .description = "Calculate Fibonacci number (Lua)",
    .category = "Lua Functions",
    .params = &[_]LuaParam{
        .{ .name = "n", .description = "Index" },
    },
});
