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

// -- Black-Scholes ------------------------------------------------------------

pub const lua_bs_call = LuaFunction(.{
    .name = "Lua.BS_CALL",
    .lua_name = "bs_call",
    .description = "Black-Scholes call option price (Lua)",
    .category = "Lua Functions",
    .params = &[_]LuaParam{
        .{ .name = "S", .description = "Current stock price" },
        .{ .name = "K", .description = "Strike price" },
        .{ .name = "T", .description = "Time to maturity (years)" },
        .{ .name = "r", .description = "Risk-free rate" },
        .{ .name = "sigma", .description = "Volatility" },
    },
});

pub const lua_bs_put = LuaFunction(.{
    .name = "Lua.BS_PUT",
    .lua_name = "bs_put",
    .description = "Black-Scholes put option price (Lua)",
    .category = "Lua Functions",
    .params = &[_]LuaParam{
        .{ .name = "S", .description = "Current stock price" },
        .{ .name = "K", .description = "Strike price" },
        .{ .name = "T", .description = "Time to maturity (years)" },
        .{ .name = "r", .description = "Risk-free rate" },
        .{ .name = "sigma", .description = "Volatility" },
    },
});

// -- Thread-safe Lua functions ------------------------------------------------
// These run on Excel's multi-threaded calc engine. Each thread gets its own
// Lua state from the pool, so there's no mutex contention.

pub const lua_is_prime = LuaFunction(.{
    .name = "Lua.IsPrime",
    .lua_name = "is_prime",
    .description = "Check if a number is prime (Lua, thread-safe)",
    .category = "Lua Functions",

    .params = &[_]LuaParam{
        .{ .name = "n", .description = "Number to check" },
    },
});

pub const lua_sum_range = LuaFunction(.{
    .name = "Lua.SumRange",
    .lua_name = "sum_range",
    .description = "Sum integers from lo to hi (Lua, thread-safe)",
    .category = "Lua Functions",

    .params = &[_]LuaParam{
        .{ .name = "lo", .description = "Start of range" },
        .{ .name = "hi", .description = "End of range" },
    },
});

// -- Async Lua functions ------------------------------------------------------
// These run on the framework's thread pool. The cell shows #N/A while the
// Lua function executes, then updates with the result via RTD.

pub const lua_slow_fib = LuaFunction(.{
    .name = "Lua.SlowFib",
    .lua_name = "slow_fib",
    .description = "Fibonacci with simulated delay (Lua, async)",
    .category = "Lua Functions",
    .is_async = true,
    .params = &[_]LuaParam{
        .{ .name = "n", .description = "Index" },
    },
});

pub const lua_slow_prime_count = LuaFunction(.{
    .name = "Lua.SlowPrimeCount",
    .lua_name = "slow_prime_count",
    .description = "Count primes up to limit (Lua, async)",
    .category = "Lua Functions",
    .is_async = true,
    .params = &[_]LuaParam{
        .{ .name = "limit", .description = "Upper bound" },
    },
});
