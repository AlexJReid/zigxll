# ZigXLL

A Zig package for creating Excel custom functions. Cross-compiles to Windows from Mac or Linux: no Windows install needed to build! This means you can use cheaper Linux CI runners.

There's a [standalone repo here](https://github.com/AlexJReid/zigxll-standalone/) which can be used as a template.

## Why

This exists because I wanted to see if it was possible to use Zig's C interop and comptime to make the Excel C SDK nicer to work with.

This came about as I'm working on [xllify](https://xllify.com) in C++ and Luau. I have no complaints, other than curiosity over how this would look in Zig. One day maybe xllify will be Zig - it's too soon to tell.

Anyway, we end up with:

- **C performance but not C**: Higher level. No boilerplate. Memory rules enforced.
- **Zero boilerplate**: No need to export `xlAutoOpen`, `xlAutoClose`, etc. - the framework handles it all
- **Automatic discovery**: Just add an `ExcelFunction()` and reference your function
- **Type safety**: Zig types automatically convert to/from Excel values (support for ranges soon)
- **Thread-safe by default**: Functions marked thread-safe automatically for MTR
- **UTF-8 strings**: Write Zig code with normal `[]u8` strings, framework handles UTF-16 conversion
- **Error handling**: Zig errors automatically become `#VALUE!` in Excel
- **comptime**: Compilation-time code generation balances conciseness without affecting runtime performance

## Quick start

> See [example](./example) for a complete working project.

Add ZigXLL as a dependency in your `build.zig.zon`:

```zig
.dependencies = .{
    .xll = .{
        .url = "https://github.com/alexjreid/zigxll/archive/refs/tags/v0.2.5.tar.gz",
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

    const user_module = b.createModule(.{
        .root_source_file = b.path("src/main.zig"),
        .target = target,
        .optimize = optimize,
    });

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

Define your functions:

```zig
// src/my_functions.zig
const xll = @import("xll");
const ExcelFunction = xll.ExcelFunction;
const ParamMeta = xll.ParamMeta;

pub const add = ExcelFunction(.{
    .name = "add",
    .description = "Add two numbers",
    .category = "My Functions",
    .params = &[_]ParamMeta{
        .{ .name = "a", .description = "First number" },
        .{ .name = "b", .description = "Second number" },
    },
    .func = addImpl,
});

fn addImpl(a: f64, b: f64) !f64 {
    return a + b;
}
```

Wire them up in `src/main.zig`:

```zig
pub const function_modules = .{
    @import("my_functions.zig"),
};
```

Build and load:

```bash
zig build
```

The XLL lands in `zig-out/lib/my_functions.xll`. Double click to load in Excel.

## Documentation

- [Creating functions](./userdocs/functions.md) — types, options, returning strings/arrays, namespacing
- [RTD servers](./userdocs/rtd-servers.md) — pushing live data to Excel, using RTD from UDFs
- [How it works](./userdocs/how-it-works.md) — comptime code generation, architecture

## Cross compiling on Mac or Linux

XLL add-ins can only run on Windows Excel. But thanks to Zig, you can build them on cheaper Linux-based CI runners or your Mac dev machine. You will still need Windows (or some VM, check out Azure for good value remote ones) to actually try out your XLL.

Tests can run natively without any Windows SDK:

```bash
zig build test
```

To cross-compile the XLL from Mac/Linux, install [xwin](https://jake-shadle.github.io/xwin/) to get the Windows SDK and CRT libraries:

**macOS:**
```bash
brew install xwin
xwin --accept-license splat --output ~/.xwin
```

**Linux:**
```bash
cargo install xwin
xwin --accept-license splat --output ~/.xwin
```

If you don't have Cargo, [install Rust](https://rustup.rs/) first, or download a prebuilt xwin binary from the [releases page](https://github.com/Jake-Shadle/xwin/releases).

Once set up, `zig build` will automatically detect `~/.xwin` and cross-compile the XLL.

## Dependencies

This library uses the **Microsoft Excel 2013 XLL SDK** headers and libraries, included in the `excel/` directory.

- **Download**: https://www.microsoft.com/en-gb/download/details.aspx?id=35567
- **Files used**: `xlcall.h`, `FRAMEWRK.H`, `xlcall32.lib`, `frmwrk32.lib`

By using this software you agree to the EULA specified by Microsoft in the above download.

## License

[MIT](./LICENSE)
