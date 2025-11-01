# ZigXLL

A zero-boilerplate library for creating Excel XLL add-ins that provide custom functions in Zig.

## Why

This is all predicated on the possibility that someone would want to write fast functions for Excel in Zig. So that makes it very niche. 

It exists because I wanted to see if it was possible to use Zig's C interop and comptime to make the Excel C SDK nicer to work with. I think it works quite nicely already, but I'd be glad of your feedback. I'm [@alexjreid](https://x.com/AlexJReid) on X.

This came about as I'm working on [xllify](https://xllify.com) in C++ and Luau. I've no complaints here, other than curiousity over how a simple Zig framework would look. So I figured I'd just push the little Zig port here.

>Disclaimer: I'm still learning Zig (it's a moving target to put it mildly) so there will be gotchas. Claude helped a lot with the comptime stuff. Thanks, Claude.

Anyway, back to this implementation. We end up with:

- **C performance but not C**: Higher level. No boiler plate. Memory rules enforced.
- **Zero boilerplate**: No need to export `xlAutoOpen`, `xlAutoClose`, etc. - the framework handles it all
- **Automatic discovery**: Just add an `ExcelFunction()` and reference your function
- **Type safety**: Zig types automatically convert to/from Excel values (support for ranges soon)
- **Thread-safe by default**: Functions marked thread-safe automatically for MTR
- **UTF-8 strings**: Write Zig code with normal `[]u8` strings, framework handles UTF-16 conversion
- **Error handling**: Zig errors automatically become `#VALUE!` in Excel
- **comptime**: _Stuff_ happens at compile time to give the balance of concise code, without affecting runtime performance

See [HOW_IT_WORKS](./HOW_IT_WORKS.md) for technical details.


## Walkthrough

> See [example](./example) for a simple working example.

Add ZigXLL as a dependency in your `build.zig.zon`:

```zig
.dependencies = .{
    .xll = .{
        .url = "https://github.com/alexjreid/zigxll/archive/refs/tags/v0.1.0.tar.gz",
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
- `[]const u8` - Strings (UTF-8)
- `*XLOPER12` - Raw Excel values (advanced)

**Return types:**
- `f64` - Numbers
- `[]const u8` / `[]u8` - Strings (automatically freed by Excel)
- `*XLOPER12` - Raw Excel values (advanced)

Functions return `!T` - errors become `#VALUE!` in Excel. Support for ranges is next.

## Available options for `ExcelFunction`

```zig
pub const myFunc = ExcelFunction(.{
    .name = "myFunc",
    .description = "My function",
    .category = "MyCategory",
    .params = &[_]ParamMeta{
        .{ .name = "x", .description = "Parameter help text" },
        .{ .description = "Name is optional" },
    },
    .func = myFuncImpl,
    .thread_safe = true, // Default is true
});
```

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
