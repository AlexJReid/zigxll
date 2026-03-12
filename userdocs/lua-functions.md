# Lua Functions

ZigXLL can embed Lua scripts in your XLL, letting you write Excel functions in Lua instead of Zig. The framework handles all marshaling between Excel and Lua at compile time, with no runtime registry or stub pools needed. Lua functions support async execution and thread-safe parallel recalculation.

## Overview

You write Lua scripts with plain functions, embed them with `@embedFile`, and declare their Excel signatures using `LuaFunction()`. The framework generates C-callable wrappers at compile time (same pattern as `ExcelFunction()`) and manages a pool of Lua states at runtime.

Lua support is optional and off by default. It compiles Lua 5.4 from source as part of the build, so there are no external dependencies to manage.

## Quick start

### 1. Enable Lua in your build

In your `build.zig`, pass `.enable_lua = true` to `buildXll()`. This is the only change needed to your build file (the rest of your `build.zig` stays the same as shown in the main [README](../README.md#quick-start)):

```zig
const xll = xll_build.buildXll(b, .{
    .name = "my_functions",
    .user_module = user_module,
    .target = target,
    .optimize = optimize,
    .enable_lua = true,  // add this line
});
```

### 2. Write a Lua script

**`src/lua/functions.lua`:**

```lua
function add(x, y)
    return x + y
end

function greet(name)
    return "Hello, " .. name .. "!"
end

function hypotenuse(a, b)
    return math.sqrt(a * a + b * b)
end
```

### 3. Declare Excel function signatures

**`src/lua_functions.zig`:**

```zig
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
```

### 4. Wire up in main.zig

Add your Lua function module to `function_modules` as you would any other module, and add a `lua_scripts` tuple to embed your Lua source files:

```zig
pub const function_modules = .{
    @import("lua_functions.zig"),
};

pub const lua_scripts = .{
    .{ .name = "functions", .source = @embedFile("lua/functions.lua") },
};
```

Each script is embedded into the XLL binary and executed when the add-in loads (`xlAutoOpen`), making its global functions available for the `LuaFunction` wrappers to call.

## LuaFunction options

```zig
pub const my_func = LuaFunction(.{
    .name = "Lua.MyFunc",       // Excel function name (dots OK for namespacing)
    .lua_name = "my_func",      // Name of the Lua global function to call
    .description = "What it does",
    .category = "My Category",
    .params = &[_]LuaParam{ ... },
});
```

| Field | Required | Default | Notes |
|---|---|---|---|
| `name` | yes | | Function name in Excel. Dots allowed for namespacing. |
| `lua_name` | no | same as `name` | The Lua global function to call. Set this when the Excel name differs from the Lua name. |
| `description` | no | `""` | Shown in Excel's Insert Function dialog. |
| `category` | no | `"Lua"` | Groups the function in Excel's function list. |
| `params` | no | `&.{}` | Array of `LuaParam` structs. Must match the Lua function's arity. |
| `thread_safe` | no | `false` | When `true`, Excel can call from multiple threads. Each thread acquires its own Lua state from the pool. |
| `async` | no | `false` | When `true`, runs on a worker thread with result caching via RTD. Same pattern as Zig async functions. |

## Parameter types

Each `LuaParam` declares the expected type for marshaling between Excel and Lua:

| `LuaParamType` | Excel to Lua | Lua to Excel |
|---|---|---|
| `.number` (default) | XLOPER12 number to Lua number | Lua number to XLOPER12 number |
| `.string` | XLOPER12 string to Lua string (UTF-8) | Lua string to XLOPER12 string |
| `.boolean` | XLOPER12 boolean to Lua boolean | Lua boolean to XLOPER12 boolean |

Return types are detected automatically from whatever the Lua function returns. Numbers, strings, booleans, and nil are all supported.

## How it works

At compile time, `LuaFunction()` generates:

1. A C-callable `impl` function with the exact arity Excel expects (0-8 params)
2. An `@export` of the impl function (dots in names become underscores)
3. Excel registration metadata (type string, descriptions)

At runtime, the framework maintains a pool of independent Lua states (default 4, [configurable](#pool-size)), each loaded with identical scripts. When Excel calls the function:

1. **Non-thread-safe**: locks the main state (slot 0)
2. **Thread-safe**: acquires any free state from the pool via atomic CAS (no contention — each thread gets its own state)
3. **Async**: spawns a worker thread that acquires a pool state, runs the Lua function, stores the result in the async cache, and notifies Excel via RTD

For sync calls (both thread-safe and non-thread-safe), the wrapper:

1. Acquires a Lua state
2. Looks up the Lua function by name (`lua_name`)
3. Pushes each XLOPER12 argument onto the Lua stack, converting based on the declared `LuaParamType`
4. Calls the Lua function via `lua_pcall`
5. Pulls the return value off the Lua stack and wraps it as an XLOPER12
6. Releases the state and returns the result to Excel

For async calls, the same Lua call happens on a worker thread, and the result is cached so subsequent recalculations return instantly.

If anything goes wrong (no state available, function not found, type conversion failure, Lua runtime error), the wrapper returns `#VALUE!`.

## Async Lua functions

Add `.async = true` to run a Lua function on a worker thread. The first call returns `#N/A` while computing; once complete, the result is cached and returned instantly on recalculation.

```zig
pub const lua_slow = LuaFunction(.{
    .name = "Lua.SlowCalc",
    .lua_name = "slow_calc",
    .description = "A slow calculation (async)",
    .@"async" = true,
    .params = &[_]LuaParam{
        .{ .name = "x", .description = "Input value" },
    },
});
```

This uses the same async infrastructure as Zig `ExcelFunction(.{ .async = true })` — same cache, same RTD server, same fire-and-forget pattern.

## Thread-safe Lua functions

Add `.thread_safe = true` to allow Excel to call the function from multiple threads during parallel recalculation. Each thread acquires its own Lua state from the pool, so there is no contention.

```zig
pub const lua_fast = LuaFunction(.{
    .name = "Lua.FastCalc",
    .lua_name = "fast_calc",
    .thread_safe = true,
    .params = &[_]LuaParam{
        .{ .name = "x" },
    },
});
```

**Important**: since each pool state is independent, global variables set by one call may not be visible to the next (which may run on a different state). Don't rely on global mutation across calls. Use `xll.get`/`xll.set` for [shared state](#shared-state).

## Pool size

The number of Lua states defaults to 4. Override it in your `build.zig`:

```zig
const xll = xll_build.buildXll(b, .{
    .name = "my_functions",
    .user_module = user_module,
    .target = target,
    .optimize = optimize,
    .enable_lua = true,
    .lua_states = 8,
});
```

Or from the command line when building the framework directly:

```bash
zig build -Dlua_states=8
```

A value of 0 (the default) uses 4 states. Each state is an independent Lua VM with its own globals and GC, so memory usage scales linearly.

## Shared state

Since pool states are independent, global variables don't propagate between them. For state that needs to be visible across all states and threads, use the built-in `xll` library:

```lua
-- xll.set(key, value) — store a value (number, string, boolean, or nil to delete)
xll.set("counter", (xll.get("counter") or 0) + 1)
xll.set("status", "ready")

-- xll.get(key) — retrieve a value (returns nil if not set)
local count = xll.get("counter")
```

Access is serialized via a mutex on the Zig side — Lua execution stays parallel, only `xll.get`/`xll.set` calls block briefly. The store is shared across all pool states and persists for the lifetime of the add-in.

## Multiple scripts

You can embed multiple Lua scripts. They are all loaded into every pool state, so functions defined in one script are visible to others:

```zig
pub const lua_scripts = .{
    .{ .name = "helpers", .source = @embedFile("lua/helpers.lua") },
    .{ .name = "finance", .source = @embedFile("lua/finance.lua") },
};
```

Scripts execute in order, so later scripts can call functions defined in earlier ones.

## Mixing Lua and Zig functions

Lua functions and Zig functions coexist in the same XLL. List both kinds of modules in `function_modules`:

```zig
pub const function_modules = .{
    @import("my_functions.zig"),       // Zig ExcelFunction definitions
    @import("lua_functions.zig"),      // Lua LuaFunction definitions
};
```

Both are registered with Excel in `xlAutoOpen`. There is no difference from Excel's perspective.

## Sandbox

The Lua state is sandboxed by default. The following are removed before any user scripts run:

| Removed | Reason |
|---|---|
| `dofile`, `loadfile`, `load`, `require` | Prevents loading code from the filesystem or arbitrary bytecode |
| `io` (entire library) | No filesystem access |
| `os.execute`, `os.remove`, `os.rename`, `os.tmpname`, `os.getenv`, `os.exit` | No shell access or process control |

Safe functions are kept: `os.time`, `os.clock`, `os.date`, `os.difftime`, the full `math`, `string`, and `table` libraries, and all standard Lua builtins like `pairs`, `ipairs`, `tostring`, `tonumber`, `error`, `pcall`, `type`, `select`, `unpack`, etc.

Since `require` is removed, use globals or table namespaces to share code between scripts. Scripts in the `lua_scripts` tuple all run in the same state in order, so earlier scripts can define utilities that later scripts use:

```lua
-- utils.lua (loaded first)
utils = {}
function utils.format_price(x)
    return string.format("%.2f", x)
end
```

```lua
-- functions.lua (loaded second)
function my_func(x)
    return utils.format_price(x * 1.1)
end
```

## Limitations

- Maximum 8 parameters per function (same as `ExcelFunction`)
- No matrix/table parameter or return type support yet
- Pool states are independent — global variable mutations don't propagate between states (use `xll.get`/`xll.set` for shared state)
- `async` and `thread_safe` cannot both be `true` on the same function (compile error)
