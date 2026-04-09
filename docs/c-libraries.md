---
layout: default
title: Using C/C++ Libraries
---

# Using C/C++ Libraries

Zig is a C compiler and has first-class C interop. You can call any C or C++ library from your Excel functions without any special framework support -- just import the library and wrap it with `ExcelFunction()`.

## Linking a C library

Add the library to your `build.zig`:

```zig
const xll = @import("xll");

pub fn build(b: *std.Build) void {
    const lib = xll.buildXll(b, .{
        .name = "my_addin",
        .xll_module = b.path("src/main.zig"),
    });

    // Link a system library
    lib.linkSystemLibrary("mylib");

    // Or link a static library built from C source
    lib.addCSourceFiles(.{
        .files = &.{ "vendor/mylib.c" },
        .flags = &.{ "-std=c99" },
    });
    lib.addIncludePath(b.path("vendor/"));
}
```

## Wrapping C functions

Import the C headers and write a thin Zig wrapper:

```zig
const xll = @import("xll");
const ExcelFunction = xll.ExcelFunction;
const ParamMeta = xll.ParamMeta;

const c = @cImport({
    @cInclude("mylib.h");
});

fn calcImpl(x: f64, y: f64) !f64 {
    return c.expensive_calculation(x, y);
}

pub const calc = ExcelFunction(.{
    .name = "Calc",
    .func = calcImpl,
    .params = &[_]ParamMeta{
        .{ .name = "x" },
        .{ .name = "y" },
    },
});
```

The Zig wrapper handles XLOPER12 type conversion as usual. The C function just sees normal C types.

## C++ libraries

For C++ libraries, create a C wrapper header that exposes the functions you need with `extern "C"` linkage, then import that header with `@cImport`. Zig cannot import C++ headers directly, but the C wrapper is typically straightforward.

## When to use Zig vs Lua vs C

| Approach | Best for |
|----------|----------|
| **Zig** | New code, performance-critical functions, full framework access |
| **Lua** | Rapid iteration, user-editable logic, simple calculations |
| **C library** | Reusing existing native code, vendor libraries, numerical libraries |

All three can coexist in the same XLL.
