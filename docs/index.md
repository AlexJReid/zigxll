---
layout: default
title: ZigXLL
---

# ZigXLL

A Zig framework for building Excel XLL add-ins. Define functions in Zig (or Lua), and the framework generates all Excel boilerplate at compile time: exports, type conversions, registration, COM vtables for RTD.

[GitHub](https://github.com/AlexJReid/zigxll) ·
[Standalone template](https://github.com/AlexJReid/zigxll-standalone/) ·
[Download example XLL](https://github.com/AlexJReid/zigxll/releases/latest)

## Why XLLs?

XLL add-ins are native DLLs that run inside the Excel process with no serialisation or IPC overhead. Excel calls your functions directly and can parallelise them across cores during recalculation.

The catch: the C SDK dates from the early 1990s. Memory management is manual, the type system is painful, and there is almost no tooling. Microsoft themselves call it "impractical for most users."

## Why Zig?

Zig's C interop and comptime make the SDK usable. You write normal Zig functions with standard types. The framework generates all the Excel boilerplate at compile time.

- **No boilerplate** -- define functions with `ExcelFunction()` and macros with `ExcelMacro()`, the framework handles the rest
- **Type-safe conversions** between Zig types and XLOPER12
- **UTF-8 strings** -- the framework handles UTF-16 conversion
- **Error mapping** -- Zig errors become `#VALUE!` in Excel, or return specific errors like `#N/A`, `#DIV/0!`
- **Thread-safe by default** (Multi-Threaded Recalculation)
- **Zero function call overhead** -- 2000 Black-Scholes calculations recalc in under 7ms on a basic PC
- **Cross-compile from Mac/Linux** via [xwin](https://jake-shadle.github.io/xwin/)
- **Async functions** -- add `.is_async = true` to run on a thread pool with automatic caching
- **RTD servers in pure Zig** -- no ATL/MFC needed
- **Embedded Lua scripting** -- write Excel functions in Lua with automatic type marshalling, async, and thread-safe support

## Quick start

Add ZigXLL as a dependency in your `build.zig.zon`:

```zig
.dependencies = .{
    .xll = .{
        .url = "https://github.com/alexjreid/zigxll/archive/refs/tags/v0.3.1.tar.gz",
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
    const optimize = b.standardOptimizeOption(.{
        .preferred_optimize_mode = .ReleaseSmall,
    });

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

    const install_xll = b.addInstallFile(
        xll.getEmittedBin(), "lib/my_functions.xll",
    );
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

## Cross-compilation

Tests run natively without any Windows SDK:

```bash
zig build test
```

To cross-compile the XLL from Mac or Linux, install [xwin](https://jake-shadle.github.io/xwin/) for Windows SDK/CRT libraries:

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

Once set up, `zig build` auto-detects `~/.xwin` and cross-compiles.

## Performance

The example project includes Black-Scholes option pricing. 2000 calculations (1000 rows, call + put) recalculate in around 4-7ms on a basic AMD Ryzen 5500U.

![2000 Black-Scholes recalculations in ~6.8ms](bs2000.png)

The Lua implementation of the same calculations is only 2-3ms slower, perfectly acceptable for interactive Excel use.

## Blog post

[zigxll: building Excel XLL add-ins in Zig](https://alexjreid.dev/posts/zigxll/) -- background, motivation, and how the framework came together.

## Documentation

- [Creating functions](functions) -- types, options, returning strings/arrays, async, macros
- [RTD servers](rtd-servers) -- pushing live data to Excel, using RTD from UDFs
- [Lua functions](lua-functions) -- writing Excel functions in Lua with JSON or Zig definitions
- [Using C/C++ libraries](c-libraries) -- calling existing native code from Excel functions
- [How it works](how-it-works) -- comptime code generation, architecture

## Example functions

The [example project](https://github.com/AlexJReid/zigxll/tree/main/example) includes:

| Function | Description |
|---|---|
| `ZigXLL.DOUBLE(x)` | Doubles a number |
| `ZigXLL.REVERSE(text)` | Reverses a string |
| `ZigXLL.BS_CALL(S, K, T, r, sigma)` | Black-Scholes call option price |
| `ZigXLL.BS_PUT(S, K, T, r, sigma)` | Black-Scholes put option price |
| `ZigXLL.TIMER()` | Live ticking counter (RTD wrapper) |
| `ZigXLL.SLOW_DOUBLE(x)` | Async doubled number (simulates slow computation) |
| `ZigXLL.MONTE_CARLO(batches, samples)` | Estimate pi via Monte Carlo with live progress |

## Alternatives

Using Zig for XLL development is a niche within a niche. Here are some alternatives to benchmark ZigXLL against to see which best fits your needs:

- **[xladd](https://github.com/MarcusRainbow/xladd)** (Rust) - Rust wrapper around the Excel C API. Proc macros generate registration boilerplate. Similar philosophy to ZigXLL but with Rust's ecosystem and crate support. See also [xladd-derive](https://github.com/ronniec95/xladd-derive).
- **[Excel-DNA](https://excel-dna.net/)** (.NET) - The most mature option. Write UDFs in C#, VB.NET, or F#, pack everything into a single .xll. Huge community, great docs, production-proven. If you're already in the .NET ecosystem, start here.
- **[PyXLL](https://www.pyxll.com/)** (Python) - Commercial. Runs Python inside Excel with full access to NumPy, Pandas, etc. Decorate functions to expose them as UDFs. Great if your logic is already in Python. Windows only.
- **[xlwings](https://www.xlwings.org/)** (Python) - Open-source core (BSD), commercial PRO and Server tiers. Call Python from Excel and vice versa. UDFs on Windows, automation on both Windows and Mac. Also supports Google Sheets and Excel on the web.

Honourable mention: **[xllify](https://xllify.com)** is not quite the same thing - it's a platform I built on ZigXLL that lets you create Excel function add-ins for Windows, Mac, and the web without writing Zig (or any code). Describe your functions in plain English or paste existing VBA, and it generates the add-in for you.

## Projects using ZigXLL

- [xllify](https://xllify.com) -- Platform for building custom Excel function add-ins for Windows, Mac, and the web
- [zigxll-nats](https://github.com/AlexJReid/zigxll-nats) -- Stream NATS messages into Excel as live data

## Commercial

ZigXLL is the MIT-licenced core behind [xllify.com](https://xllify.com), a platform for building custom Excel function add-ins for Windows, Mac, and the web.

Custom Excel add-in development is also available. Get in touch at [alex@lexvica.com](mailto:alex@lexvica.com).

## Licence

[MIT](https://github.com/AlexJReid/zigxll/blob/main/LICENSE)
