# ZigXLL

A Zig package for implementing Excel custom functions against the C SDK.

There's a [standalone repo here](https://github.com/AlexJReid/zigxll-standalone/) which can be used as a template.

## Why

This exists because I wanted to see if it was possible to use Zig's C interop and comptime to make the Excel C SDK nicer to work with. I think it works quite nicely already, but I'd be glad of your feedback. I'm [@alexjreid](https://x.com/AlexJReid) on X.

This came about as I'm working on [xllify](https://xllify.com) in C++ and Luau. I have no complaints, other than curiosity over how this would look in Zig. One day maybe xllify will be Zig - it's too soon to tell as I'm still learning the language (it's a moving target.)

Claude helped a lot with the comptime stuff and the demos. Thanks, Claude.

Anyway, we end up with:

- **C performance but not C**: Higher level. No boilerplate. Memory rules enforced.
- **Zero boilerplate**: No need to export `xlAutoOpen`, `xlAutoClose`, etc. - the framework handles it all
- **Automatic discovery**: Just add an `ExcelFunction()` and reference your function
- **Type safety**: Zig types automatically convert to/from Excel values (support for ranges soon)
- **Thread-safe by default**: Functions marked thread-safe automatically for MTR
- **UTF-8 strings**: Write Zig code with normal `[]u8` strings, framework handles UTF-16 conversion
- **Error handling**: Zig errors automatically become `#VALUE!` in Excel
- **comptime**: Compilation-time code generation balances conciseness without affecting runtime performance

See [how it works](./HOW_IT_WORKS.md) for more details.

## Walkthrough

> See [example](./example) for a simple working example.

Add ZigXLL as a dependency in your `build.zig.zon`:

```zig
.dependencies = .{
    .xll = .{
        .url = "https://github.com/alexjreid/zigxll/archive/refs/tags/v0.2.0.tar.gz",
        .hash = "...",
    },
},
```

Create your `build.zig`:

```zig
const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.resolveTargetQuery(.{
        .cpu_arch = .x86_64,
        .os_tag = .windows,
        .abi = .msvc,
    });
    const optimize = b.standardOptimizeOption(.{ .preferred_optimize_mode = .ReleaseSmall });

    // Create a module for your functions
    const user_module = b.createModule(.{
        .root_source_file = b.path("src/main.zig"),
        .target = target,
        .optimize = optimize,
        .strip = true,
    });

    // Build the XLL using the framework helper
    const xll_build = @import("xll");
    const xll = xll_build.buildXll(b, .{
        .name = "my_functions",
        .user_module = user_module,
        .target = target,
        .optimize = optimize,
    });

    const install_xll = b.addInstallFile(xll.getEmittedBin(), "lib/my_functions.xll");
    b.getInstallStep().dependOn(&install_xll.step);
}
```

`src/main.zig` lists your function modules:

```zig
pub const function_modules = .{
    @import("my_functions.zig"),
};
```

`src/my_functions.zig` defines your Excel functions:

```zig
const std = @import("std");
const xll = @import("xll");
const ExcelFunction = xll.ExcelFunction;
const ParamMeta = xll.ParamMeta;

pub const add = ExcelFunction(.{
    .name = "add",
    .description = "Add two numbers",
    .category = "Zig Math",
    .params = &[_]ParamMeta{
        .{ .name = "a", .description = "First number" },
        .{ .name = "b", .description = "Second number" },
    },
    .func = addImpl,
});

fn addImpl(a: f64, b: f64) !f64 {
    return a + b;
}

// You can namespace functions using dots in the name
pub const bs_call = ExcelFunction(.{
    .name = "MyFunctions.BSCall",
    .description = "Black-Scholes European Call Option Price",
    .category = "Finance",
    .params = &[_]ParamMeta{
        .{ .name = "S", .description = "Current stock price" },
        .{ .name = "K", .description = "Strike price" },
        .{ .name = "T", .description = "Time to maturity (years)" },
        .{ .name = "r", .description = "Risk-free rate" },
        .{ .name = "sigma", .description = "Volatility" },
    },
    .func = blackScholesCall,
});

fn blackScholesCall(S: f64, K: f64, T: f64, r: f64, sigma: f64) !f64 {
    if (T <= 0) return error.InvalidMaturity;
    // ... implementation
    return call_price;
}
```

Build:

```bash
zig build
```

The XLL lands in `zig-out/lib/my_functions.xll`. Double click to load in Excel.

## Supported types

**Parameters:**
- `f64` - Numbers
- `bool` - Booleans (TRUE/FALSE)
- `[]const u8` - Strings (UTF-8)
- `[][]const f64` - 2D arrays/ranges of numbers (empty cells become 0.0)
- `*XLOPER12` - Raw Excel values (advanced)

**Return types:**
- `f64` - Numbers
- `bool` - Booleans (TRUE/FALSE)
- `[]const u8` / `[]u8` - Strings (automatically freed by Excel)
- `[][]const f64` / `[][]f64` - 2D arrays/ranges of numbers (automatically freed by Excel)
- `*XLOPER12` - Raw Excel values (advanced)

**Optional parameters:** Use `?T` types (like `?f64`, `?bool`, `?[]const u8`) for optional parameters.

Functions return `!T` - errors become `#VALUE!` in Excel.

## Available options for `ExcelFunction`

```zig
pub const myFunc = ExcelFunction(.{
    .name = "myFunc",
    .description = "My function",
    .category = "MyCategory",
    .params = &[_]ParamMeta{
        .{ .name = "x", .description = "Parameter help text" },
        .{ .name = "y", .description = "Optional parameter (default 10)" },
        .{ .description = "Name is optional" },
    },
    .func = myFuncImpl,
    .thread_safe = true, // Default is true
});

// Function with optional parameter
fn myFuncImpl(x: f64, y: ?f64, z: f64) !f64 {
    const y_val = y orelse 10.0; // Use default if not provided
    return x + y_val + z;
}
```

**Optional parameters**: Use `?T` types (like `?f64`, `?bool`) for optional parameters. When Excel passes a missing value, it becomes `null`. Use the `orelse` operator to provide default values in your implementation.

## Built-in Example Functions

The framework includes several example functions in `src/builtin_functions.zig` that demonstrate different features:

### ZigDouble
```
=ZigDouble(x, y, z)
```
Demonstrates basic numeric operations: returns `x * 2 + y - z`

**Example:** `=ZigDouble(5, 3, 1)` → 12

### ZigMatrix
```
=ZigMatrix([rows], [cols])
```
Returns a matrix filled with sequential numbers. Both parameters are optional.

- **Default:** 10×5 matrix (50 values: 1-50)
- **Max size:** 100×100
- **Examples:**
  - `=ZigMatrix()` → 10×5 matrix
  - `=ZigMatrix(3, 7)` → 3×7 matrix (21 values: 1-21)

Demonstrates: optional parameters, 2D array return values (`[][]f64`)

### ZigNot
```
=ZigNot(value)
```
Returns the logical NOT of a boolean value.

**Examples:**
- `=ZigNot(TRUE)` → FALSE
- `=ZigNot(A1>10)` → inverts the comparison

Demonstrates: boolean parameter and return type

### ZigPower
```
=ZigPower(base, [exponent])
```
Raises a number to a power. Exponent is optional (default: 2).

**Examples:**
- `=ZigPower(5)` → 25 (5²)
- `=ZigPower(5, 3)` → 125 (5³)

Demonstrates: optional numeric parameter with default value

### ZigSumRange
```
=ZigSumRange(data)
```
Sums all values in a range or 2D array.

**Examples:**
- `=ZigSumRange(A1:C5)` → sums all values in the range
- `=ZigSumRange(A:A)` → sums entire column A
- `=ZigSumRange(1:10)` → sums rows 1-10

Demonstrates: range/array input parameter (`[][]const f64`)

## Dependencies

This library uses the **Microsoft Excel 2013 XLL SDK** headers and libraries. These are included in the `excel/` directory and are required to build Excel add-ins.

- **Download**: https://www.microsoft.com/en-gb/download/details.aspx?id=35567
- **Files used**: `xlcall.h`, `FRAMEWRK.H`, `xlcall32.lib`, `frmwrk32.lib`

By using this software you agree to the EULA specified by Microsoft in the above download.

## Working on the framework

You can also clone this repo to improve the framework directly:

1. Edit `src/user_functions.zig` or create new modules
2. Run `zig build`
3. Your XLL is in `zig-out/lib/output.xll`

## License

[MIT](./LICENSE)
