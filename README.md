# ZigXLL

A Zig framework for building Excel XLL add-ins. Cross-compiles from Mac/Linux to Windows without needing a Windows install.

[Standalone template repo](https://github.com/AlexJReid/zigxll-standalone/) | [Example project](./example)

## Why XLLs

XLL add-ins are native DLLs that run inside the Excel process with no serialization or IPC overhead. Excel calls your functions directly and can parallelize them across cores during recalculation.

The catch: the C SDK dates from the early 1990s. Memory management is manual, the type system is painful, and there's almost no tooling. Microsoft themselves call it "impractical for most users."

## Why Zig

Zig's C interop and comptime make the SDK usable. You write normal Zig functions with standard types. The framework generates all the Excel boilerplate at compile time: exports, type conversions, registration, COM vtables for RTD.

What you get:

- No boilerplate - define functions with `ExcelFunction()`, framework handles the rest
- Type-safe conversions between Zig types and XLOPER12
- UTF-8 strings (framework handles UTF-16 conversion)
- Zig errors become `#VALUE!` in Excel
- Thread-safe by default (MTR)
- Cross-compilation from Mac/Linux via [xwin](https://jake-shadle.github.io/xwin/)
- Async functions - add `.async = true` to run on a thread pool with automatic caching. See [function docs](./userdocs/functions.md#async-functions)
- Pure Zig COM RTD servers - no ATL/MFC. See [RTD docs](./userdocs/rtd-servers.md)

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

Build:

```bash
zig build
```

Output lands in `zig-out/lib/my_functions.xll`. Double-click to load in Excel.

## Documentation

- [Creating functions](./userdocs/functions.md) - types, options, returning strings/arrays, namespacing
- [RTD servers](./userdocs/rtd-servers.md) - pushing live data to Excel, using RTD from UDFs
- [How it works](./userdocs/how-it-works.md) - comptime code generation, architecture

## Cross-compilation

Tests run natively without any Windows SDK:

```bash
zig build test
```

To cross-compile the XLL, install [xwin](https://jake-shadle.github.io/xwin/) for Windows SDK/CRT libraries:

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

If you don't have Cargo, [install Rust](https://rustup.rs/) or grab a prebuilt binary from the [releases page](https://github.com/Jake-Shadle/xwin/releases).

Once set up, `zig build` auto-detects `~/.xwin` and cross-compiles.

## Dependencies

Uses the **Microsoft Excel 2013 XLL SDK** headers and libraries, included in `excel/`.

- **Download**: https://www.microsoft.com/en-gb/download/details.aspx?id=35567
- **Files**: `xlcall.h`, `FRAMEWRK.H`, `xlcall32.lib`, `frmwrk32.lib`

By using this software you agree to the EULA specified by Microsoft in the above download.

## License

[MIT](./LICENSE)
