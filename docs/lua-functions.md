---
layout: default
title: Lua Functions
---

# Lua Functions

ZigXLL can embed Lua scripts in your XLL, letting you write Excel functions in Lua instead of Zig. The framework handles all marshaling between Excel and Lua at compile time, with no runtime registry or stub pools needed. Lua functions support async execution and thread-safe parallel recalculation.

## Overview

You write Lua scripts with `---` annotations and list them in `build.zig`. The framework parses the annotations, generates `LuaFunction` declarations, embeds the scripts, and registers everything with Excel automatically. No manual code generation or wiring needed.

Lua support is optional and off by default. It compiles Lua 5.4 from source as part of the build, so there are no external dependencies to manage.

## Quick start

### 1. Write annotated Lua scripts

**`src/lua/functions.lua`:**

```lua
--- Add two numbers
-- @param x number First number
-- @param y number Second number
function add(x, y)
    return x + y
end

--- Greet someone by name
-- @param name string Name to greet
function greet(name)
    return "Hello, " .. name .. "!"
end

--- Calculate hypotenuse
-- @param a number Side a
-- @param b number Side b
function hypotenuse(a, b)
    return math.sqrt(a * a + b * b)
end
```

### 2. List them in `build.zig`

```zig
const xll = xll_build.buildXll(b, .{
    .name = "my_functions",
    .user_module = user_module,
    .target = target,
    .optimize = optimize,
    .enable_lua = true,
    .lua_scripts = &.{
        "src/lua/functions.lua",
    },
});
```

That's it. The framework handles parsing annotations, generating Excel function declarations, embedding scripts, and registering everything at startup. No changes to `main.zig` required.

To add more scripts later, just add them to `lua_scripts`:

```zig
    .lua_scripts = &.{
        "src/lua/functions.lua",
        "src/lua/finance.lua",
    },
```

## Annotations

Annotate Lua functions with `---` doc comments directly above the `function` declaration:

```lua
--- Description of the function
-- @param x number First number
-- @param y string Name to greet
-- @async
-- @thread_safe false
-- @category My Category
-- @name CustomExcelName
-- @help_url https://example.com/help
function my_func(x, y) ... end
```

| Tag | Description |
|---|---|
| `---` line | Function description (first `---` line) |
| `@param name [type] [description]` | Parameter. Type is `number` (default), `string`, or `boolean`. |
| `@rtd` | RTD subscription function. Return values are `prog_id, topic1, topic2, ...`. Automatically non-thread-safe. |
| `@async` | Run on worker thread with result caching via RTD |
| `@thread_safe false` | Disable multi-threaded recalculation (default: thread-safe) |
| `@category name` | Excel function category (default: `"Lua Functions"`) |
| `@name ExcelName` | Override the auto-generated Excel name |
| `@help_url url` | URL with help information |

Without `@name`, the Excel name is auto-generated from the Lua function name: `add` becomes `Lua.Add` (prefix + PascalCase).

## Build options

The `lua_scripts` option handles code generation and embedding automatically. You can also customize the prefix and category:

```zig
const xll = xll_build.buildXll(b, .{
    // ...
    .enable_lua = true,
    .lua_scripts = &.{ "src/lua/functions.lua" },
    .lua_prefix = "MyLib.",          // default: "Lua."
    .lua_category = "My Functions",  // default: "Lua Functions"
});
```

| Option | Default | Description |
|---|---|---|
| `lua_scripts` | `&.{}` | Lua script files to embed and generate declarations from |
| `lua_prefix` | `"Lua."` | Prefix for auto-generated Excel function names |
| `lua_category` | `"Lua Functions"` | Default category in Excel's function list |

## Generating functions.json

The `lua_introspect.lua` tool can also generate an Office JS-compatible `functions.json` for use with Excel web add-ins. Pass `--functions-json` with an output path:

```bash
lua tools/lua_introspect.lua --functions-json functions.json src/lua/*.lua
```

This produces a JSON file with the `$schema`, function metadata, parameter types, and `async`/`threadSafe` flags derived from the same `---` annotations used for the XLL build.

## LuaFunction options

If you need to write `LuaFunction` declarations by hand (instead of generating them), the full set of options is:

```zig
pub const my_func = LuaFunction(.{
    .name = "Lua.MyFunc",       // Excel function name (dots OK for namespacing)
    .id = "my_func",            // Lua global function to call
    .description = "What it does",
    .category = "My Category",
    .help_url = "https://example.com/help",
    .params = &[_]LuaParam{ ... },
    .thread_safe = true,        // default
    .is_async = false,          // default
});
```

| Field | Required | Default | Notes |
|---|---|---|---|
| `name` | yes | | Function name in Excel. Dots allowed for namespacing. |
| `id` | no | same as `name` | Stable identifier — the Lua global function to call. Set this when the Excel name differs from the Lua name. |
| `description` | no | `""` | Shown in Excel's Insert Function dialog. |
| `category` | no | `"Lua"` | Groups the function in Excel's function list. |
| `help_url` | no | | URL with help information about the function. |
| `params` | no | `&.{}` | Array of `LuaParam` structs. Must match the Lua function's arity. |
| `thread_safe` | no | `true` | Enables Multi-Threaded Recalculation. Each thread acquires its own Lua state from the pool. Set to `false` if your Lua function relies on global state. |
| `is_async` | no | `false` | When `true`, runs on a worker thread with result caching via RTD. Same pattern as Zig async functions. Automatically sets `thread_safe = false`. |

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

At runtime, the framework maintains a pool of independent Lua states (default 8, [configurable](#pool-size)), each loaded with identical scripts. When Excel calls the function:

1. **Non-thread-safe**: locks the main state (slot 0)
2. **Thread-safe**: acquires any free state from the pool via atomic CAS (no contention — each thread gets its own state)
3. **Async**: spawns a worker thread that acquires a pool state, runs the Lua function, stores the result in the async cache, and notifies Excel via RTD

For sync calls (both thread-safe and non-thread-safe), the wrapper:

1. Acquires a Lua state
2. Looks up the Lua function by name (`id`)
3. Pushes each XLOPER12 argument onto the Lua stack, converting based on the declared `LuaParamType`
4. Calls the Lua function via `lua_pcall`
5. Pulls the return value off the Lua stack and wraps it as an XLOPER12
6. Releases the state and returns the result to Excel

For async calls, the same Lua call happens on a worker thread, and the result is cached so subsequent recalculations return instantly.

If anything goes wrong (no state available, function not found, type conversion failure, Lua runtime error), the wrapper returns `#VALUE!`.

## Async Lua functions

Add `@async` to the annotation (or `.is_async = true` in hand-written Zig) to run a Lua function on a worker thread. The first call returns `#N/A` while computing; once complete, the result is cached and returned instantly on recalculation.

```lua
--- Fibonacci with simulated delay
-- @param n number Index
-- @async
function slow_fib(n)
    -- simulate slow work
    local a, b = 0, 1
    for i = 1, n do a, b = b, a + b end
    return a
end
```

This uses the same async infrastructure as Zig `ExcelFunction(.{ .is_async = true })` — same cache, same RTD server, same fire-and-forget pattern.

## Thread safety

Lua functions are **thread-safe by default** — Excel can call them from multiple threads during parallel recalculation. Each thread acquires its own Lua state from the pool, so there is no contention.

If your Lua function relies on global state that must be consistent across calls, add `@thread_safe false` to the annotation (or `.thread_safe = false` in hand-written Zig).

**Important**: since each pool state is independent, global variables set by one call may not be visible to the next (which may run on a different state). Don't rely on global mutation across calls. Use `xll.get`/`xll.set` for [shared state](#shared-state).

## Pool size

The number of Lua states defaults to 8. Override it in your `build.zig`:

```zig
const xll = xll_build.buildXll(b, .{
    .name = "my_functions",
    .user_module = user_module,
    .target = target,
    .optimize = optimize,
    .enable_lua = true,
    .lua_states = 12,
});
```

Or from the command line when building the framework directly:

```bash
zig build -Dlua_states=12
```

A value of 0 (the default) uses 8 states. Each state is an independent Lua VM with its own globals and GC, so memory usage scales linearly.

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

List multiple Lua files in `lua_scripts` — they'll all be embedded and their functions registered:

```zig
    .lua_scripts = &.{
        "src/lua/helpers.lua",
        "src/lua/finance.lua",
    },
```

All scripts are loaded into every pool state, so functions defined in one script are visible to others. Scripts execute in order, so later scripts can call functions defined in earlier ones.

## Mixing Lua and Zig functions

Lua functions and Zig functions coexist in the same XLL. Zig functions go in `function_modules` in `main.zig`, Lua functions come from `lua_scripts` in `build.zig`:

```zig
// main.zig — only Zig functions
pub const function_modules = .{
    @import("my_functions.zig"),
    @import("async_functions.zig"),
};
```

```zig
// build.zig — Lua functions handled by the framework
const xll = xll_build.buildXll(b, .{
    // ...
    .enable_lua = true,
    .lua_scripts = &.{ "src/lua/functions.lua" },
});
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

Since `require` is removed, use globals or table namespaces to share code between scripts. Scripts in `lua_scripts` all run in the same state in order, so earlier scripts can define utilities that later scripts use:

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

## RTD subscriptions from Lua

Lua functions can subscribe to an RTD server and return a live-updating cell value, just like Zig wrapper functions do. Add `@rtd` to the annotation — the function returns the prog_id and topic strings, and the framework handles the `xlfRtd` call:

```lua
--- Live price for a symbol
-- @param symbol string Ticker symbol
-- @rtd
function price(symbol)
    return "myprog.rtd", symbol
end
```

The function's return values are interpreted as: first = prog_id, rest = topic strings. The framework calls `xlfRtd` with these values and returns the live cell value to Excel. The cell updates automatically whenever the RTD server pushes a new value.

`@rtd` automatically sets `thread_safe = false` (the underlying `xlfRtd` call must run on Excel's main thread). You don't need to add `@thread_safe false` separately.

Multiple topic strings are just more return values:

```lua
--- Price on a specific exchange
-- @param exchange string Exchange code
-- @param symbol string Ticker symbol
-- @rtd
function price_on_exchange(exchange, symbol)
    return "myprog.rtd", exchange, symbol
end
```

If the Lua function errors or returns no values, the cell shows `#VALUE!`.

## Limitations

- Maximum 8 parameters per function (same as `ExcelFunction`)
- No matrix/table parameter or return type support yet
- Pool states are independent — global variable mutations don't propagate between states (use `xll.get`/`xll.set` for shared state)
- `is_async = true` automatically forces `thread_safe = false` (async functions use `xlfRtd` which must run on Excel's main thread)
- `@rtd` functions are always non-thread-safe (`xlfRtd` must run on Excel's main thread)
