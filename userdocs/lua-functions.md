# Lua Functions

ZigXLL can embed Lua scripts in your XLL, letting you write Excel functions in Lua instead of Zig. The framework handles all marshaling between Excel and Lua at compile time, with no runtime registry or stub pools needed.

## Overview

You write Lua scripts with plain functions, embed them with `@embedFile`, and declare their Excel signatures using `LuaFunction()`. The framework generates C-callable wrappers at compile time (same pattern as `ExcelFunction()`) and manages a shared Lua state at runtime.

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
| `thread_safe` | | `false` | Always false. Lua states are not thread-safe. Setting `true` is a compile error. |

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

At runtime, when Excel calls the function:

1. The wrapper acquires the global Lua state
2. Looks up the Lua function by name (`lua_name`)
3. Pushes each XLOPER12 argument onto the Lua stack, converting based on the declared `LuaParamType`
4. Calls the Lua function via `lua_pcall`
5. Pulls the return value off the Lua stack and wraps it as an XLOPER12
6. Returns the result to Excel

If anything goes wrong (Lua state not initialized, function not found, type conversion failure, Lua runtime error), the wrapper returns `#VALUE!`.

## Multiple scripts

You can embed multiple Lua scripts. They all execute in the same global Lua state, so functions defined in one script are visible to others:

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
- Always non-thread-safe (`thread_safe = true` is a compile error)
- No matrix/table parameter or return type support yet
- Lua functions share a single global state, so avoid global variable collisions across scripts
