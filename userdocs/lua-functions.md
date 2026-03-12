# Lua Functions

ZigXLL can embed Lua scripts in your XLL, letting you write Excel functions in Lua instead of Zig. The framework handles all marshaling between Excel and Lua at compile time, with no runtime registry or stub pools needed. Lua functions support async execution and thread-safe parallel recalculation.

## Overview

You write Lua scripts with plain functions, embed them with `@embedFile`, and declare their Excel signatures using `LuaFunction()`. The framework generates C-callable wrappers at compile time (same pattern as `ExcelFunction()`) and manages a pool of Lua states at runtime.

Lua support is optional and off by default. It compiles Lua 5.4 from source as part of the build, so there are no external dependencies to manage.

## Quick start

There are two ways to register Lua functions: **JSON** (no Zig required) and **Zig** (using `LuaFunction()`). Both can be used together in the same project.

### Option A: JSON definitions (no Zig required)

This approach lets you define Excel function signatures in a JSON file. You only need to touch `build.zig` once during initial setup.

#### 1. Enable Lua and point to your JSON file in `build.zig`

```zig
const xll = xll_build.buildXll(b, .{
    .name = "my_functions",
    .user_module = user_module,
    .target = target,
    .optimize = optimize,
    .enable_lua = true,
    .lua_json = b.path("src/lua_functions.json"),
});
```

#### 2. Write a Lua script

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

#### 3. Define function signatures in JSON

**`src/lua_functions.json`:**

```json
[
  {
    "name": "Lua.Add",
    "lua_name": "add",
    "description": "Add two numbers",
    "category": "Lua Functions",
    "params": [
      { "name": "x", "description": "First number" },
      { "name": "y", "description": "Second number" }
    ]
  },
  {
    "name": "Lua.Greet",
    "lua_name": "greet",
    "description": "Greet someone by name",
    "params": [
      { "name": "name", "type": "string", "description": "Name to greet" }
    ]
  },
  {
    "name": "Lua.Hypotenuse",
    "lua_name": "hypotenuse",
    "description": "Calculate hypotenuse",
    "category": "Lua Functions",
    "params": [
      { "name": "a", "description": "Side a" },
      { "name": "b", "description": "Side b" }
    ]
  }
]
```

#### 4. Embed scripts in main.zig

Your `main.zig` still needs to embed the Lua scripts (so they're compiled into the XLL binary):

```zig
pub const function_modules = .{};  // can be empty if all Lua funcs are in JSON

pub const lua_scripts = .{
    .{ .name = "functions", .source = @embedFile("lua/functions.lua") },
};
```

The JSON file is read at build time only — it is not embedded in the XLL. The build system generates `LuaFunction` declarations from it, which are compiled into the binary like hand-written Zig definitions.

#### JSON format reference

Each function object supports these fields:

| Field | Required | Default | Notes |
|---|---|---|---|
| `name` | yes | | Excel function name (dots OK for namespacing) |
| `lua_name` | yes | | Lua global function to call |
| `description` | no | `""` | Shown in Excel's Insert Function dialog |
| `category` | no | `"Lua"` | Groups the function in Excel's function list |
| `params` | no | `[]` | Array of parameter objects |
| `async` | no | `false` | Run on worker thread with result caching via RTD |

Each parameter object supports:

| Field | Required | Default | Notes |
|---|---|---|---|
| `name` | yes | | Parameter name shown in Excel |
| `type` | no | `"number"` | One of `"number"`, `"string"`, `"boolean"` |
| `description` | no | | Shown in Insert Function dialog |

The JSON can be a top-level array or an object with a `"functions"` key:

```json
{ "functions": [ ... ] }
```

### Option B: Zig definitions

If you prefer compile-time type safety or are already writing Zig functions, use `LuaFunction()` directly.

#### 1. Enable Lua in your build

```zig
const xll = xll_build.buildXll(b, .{
    .name = "my_functions",
    .user_module = user_module,
    .target = target,
    .optimize = optimize,
    .enable_lua = true,
});
```

#### 2. Write a Lua script

Same as Option A above.

#### 3. Declare Excel function signatures in Zig

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

#### 4. Wire up in main.zig

Add your Lua function module to `function_modules` and embed your scripts:

```zig
pub const function_modules = .{
    @import("lua_functions.zig"),
};

pub const lua_scripts = .{
    .{ .name = "functions", .source = @embedFile("lua/functions.lua") },
};
```

### Mixing both approaches

JSON-defined and Zig-defined Lua functions can coexist. Use `lua_json` in `build.zig` and `LuaFunction()` in `function_modules` at the same time — they call into the same Lua state pool and scripts. Just make sure the Excel function names don't collide.

Each script is embedded into the XLL binary and executed when the add-in loads (`xlAutoOpen`), making its global functions available for both JSON and Zig wrappers to call.

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

Add `.is_async = true` to run a Lua function on a worker thread. The first call returns `#N/A` while computing; once complete, the result is cached and returned instantly on recalculation.

```zig
pub const lua_slow = LuaFunction(.{
    .name = "Lua.SlowCalc",
    .lua_name = "slow_calc",
    .description = "A slow calculation (async)",
    .is_async = true,
    .params = &[_]LuaParam{
        .{ .name = "x", .description = "Input value" },
    },
});
```

This uses the same async infrastructure as Zig `ExcelFunction(.{ .is_async = true })` — same cache, same RTD server, same fire-and-forget pattern.

## Thread safety

Lua functions are **thread-safe by default** — Excel can call them from multiple threads during parallel recalculation. Each thread acquires its own Lua state from the pool, so there is no contention.

Set `.thread_safe = false` if your Lua function relies on global state that must be consistent across calls:

```zig
pub const lua_stateful = LuaFunction(.{
    .name = "Lua.Stateful",
    .lua_name = "stateful_calc",
    .thread_safe = false,
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

## Generating JSON from Lua scripts

If you already have Lua scripts and want to bootstrap a `functions.json`, the `tools/lua_introspect.lua` utility can introspect your scripts and emit a JSON skeleton:

```bash
lua tools/lua_introspect.lua src/lua/functions.lua > src/lua_functions.json
```

It uses `debug.getinfo` and `debug.getlocal` to extract function names and parameter names, converts `snake_case` to `PascalCase` for Excel names, and outputs JSON compatible with the `lua_json` build option.

Options:

```bash
lua tools/lua_introspect.lua --prefix "MyAddin." --category "Finance" src/lua/*.lua
```

| Flag | Default | Description |
|---|---|---|
| `--prefix` | `Lua.` | Prefix for Excel function names |
| `--category` | `Lua Functions` | Category in Excel's function list |

The generated JSON won't include `description`, `type`, or `async` fields — add those by hand after generation.

## Example: Black-Scholes option pricing

The example project includes a Black-Scholes implementation in Lua (`example/src/lua/functions.lua`) that prices European call and put options. It uses the Zelen & Severo polynomial approximation for the normal CDF, no external libraries needed.

The Lua functions `bs_call(S, K, T, r, sigma)` and `bs_put(S, K, T, r, sigma)` take five parameters: spot price, strike price, time to maturity in years, risk-free rate, and volatility. They are registered as both `Lua.BS_CALL` / `Lua.BS_PUT` (via Zig in `lua_functions.zig`) and `LuaFromJson.BS_CALL` / `LuaFromJson.BS_PUT` (via JSON in `lua_functions.json`).

Performance: the Lua version runs 2,000 calculations (1,000 rows x call + put) in about 11ms, roughly double the native Zig implementation, but still absolutely acceptable for interactive Excel use.

A test workbook `example/black_scholes_1000_rows_test_lua.xlsm` exercises these functions across 1,000 rows of randomized inputs. There is also a Zig-side counterpart workbook (`example/black_scholes_1000_rows_test.xlsm`) that tests the native `ExcelFunction` Black-Scholes.

## Limitations

- Maximum 8 parameters per function (same as `ExcelFunction`)
- No matrix/table parameter or return type support yet
- Pool states are independent — global variable mutations don't propagate between states (use `xll.get`/`xll.set` for shared state)
- `is_async = true` automatically forces `thread_safe = false` (async functions use `xlfRtd` which must run on Excel's main thread)
