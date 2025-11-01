# How ZigXLL Works

Some technical rambling about how this all works.

## Overview

ZigXLL uses Zig's _comptime_ to automatically generate Excel-compatible wrapper functions. The framework is a Zig package that mushes with user code at build time to produce a complete XLL, ready to run in Excel.

The user never writes any Excel boilerplate. The framework's `buildXll()` build helper creates a complete XLL by:

1. Taking the user's module containing `function_modules` tuple
2. Using the framework's `xll_builder.zig` as the root source file
3. The builder exports all Excel entry points (`xlAutoOpen`, etc.)
4. The builder discovers and registers user functions at compile time

## Core Components

### 1. ExcelFunction Wrapper (src/excel_function.zig)

The `ExcelFunction()` function is a compile-time code generator that takes function metadata and produces a struct containing:

- Excel registration metadata (name, description, type string)
- An `impl` function with the exact signature Excel expects
- Type conversion logic between Zig types and XLOPER12

**Exact Signature Generation**

The switch statement on `params.len` generates the exact number of parameters needed:

- 0 params: `fn impl() callconv(.c) *xl.XLOPER12`
- 1 param: `fn impl(a1: *xl.XLOPER12) callconv(.c) *xl.XLOPER12`
- 2 params: `fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12) callconv(.c) *xl.XLOPER12`

Up to 8 are permitted.

**Type Conversion**

`extractArg()` and `wrapResult()` handle conversion:

- `f64` ↔ `xltypeNum`
- `[]const u8` ↔ `xltypeStr` (UTF-8 conversion)
- `*XLOPER12` ↔ raw passthrough

More types (ranges!) are coming soon. All conversions use the XLValue wrapper for safety.

### 2. Function Discovery (src/function_discovery.zig)

`getAllFunctions()` uses comptime reflection to find Excel functions in a module. It scans declarations looking for structs with the `is_excel_function` marker and builds a comptime array of them.

### 3. XLL Builder (src/xll_builder.zig)

The XLL builder is the root source file for all generated XLLs. It:

1. Imports the framework as `"xll_framework"`
2. Imports the user module as `"user_module"`
3. Exports all Excel entry points
4. Exposes `user_functions` for framework discovery

The user never sees or edits this file. It just hooks up the framework to Excel's machinery.

### 4. Framework Entry (src/framework_entry.zig)

**Compile-Time Function Collection**

This runs at compile time to build the complete list of functions. It accesses `@import("root")` which is the `xll_builder.zig`, which exposes the user's modules.

This takes the zig defined metadata and calls Excel's `xlfRegister` function for n functions. Again this is mostly code generated at compile time.

### 5. Build Helper (build.zig)

`buildXll()` is called from user's `build.zig` and handles all the wiring. This is why user build files stay minimal.

### 6. XLValue Wrapper (src/xlvalue.zig)

XLValue wraps XLOPER12 with type safety.

**Memory Management**

XLValue tracks whether it owns memory via `m_owns_memory`. When Excel calls `xlAutoFree12()`, the framework deallocates any owned memory. Excel memory ownership rules are strict - time will tell if this is right.

**UTF-8 Conversion**

Excel uses UTF-16 wide strings. XLValue handles conversion:

- `fromUtf8String()`: UTF-8 → UTF-16 (allocates)
- `as_utf8str()`: UTF-16 → UTF-8 (allocates)

Note: the naming of these fns will be fixed!

## Sequence

### Execution of the ADD function in Excel

1. User enters `=ADD(1, 2)` in Excel
2. Excel looks up registered function "add_impl" in the XLL DLL
3. Excel calls `add_impl(XLOPER12*, XLOPER12*)` with C calling convention
4. The generated impl function:
   - Wraps each XLOPER12 in XLValue
   - Calls `extractArg()` to convert to f64
   - Calls user's `add(f64, f64)` function
   - Converts f64 result to XLOPER12 via `wrapResult()`
   - Returns pointer to XLOPER12
5. Excel displays the result
6. Later, Excel calls `xlAutoFree12()` to free the returned memory

### Comptime function registration

1. User defines function in their module (e.g., `my_functions.zig`):
   ```zig
   pub const add = ExcelFunction(.{
       .name = "add",
       .func = addImpl,
       // ...
   });
   ```

2. User lists that module in `main.zig`:
   ```zig
   pub const function_modules = .{
       @import("my_functions.zig"),
   };
   ```

3. User calls `buildXll()` in their `build.zig`, passing their module

4. Framework uses `xll_builder.zig` as the XLL root, which imports user module

5. `ExcelFunction()` runs at compile time, generating wrapper structs with `@export()`

6. `getAllFunctions()` scans modules and finds all wrappers

7. At runtime, `xlAutoOpen()` registers each discovered function with Excel

Nice and simple.

## Performance traits

**Compile Time**
- Function discovery
- Wrapper generation
- All metadata computed at compile time (no runtime overhead)

**Runtime**
- Function call: Direct C call, no reflection or indirection
- Type conversion: Minimal overhead (pointer deref and type check)
- Memory allocation: Only for strings and returned values
- Registration: One-time cost at XLL load (although you can safely load the XLL again without restarting Excel)

## Limitations

- Maximum 8 parameters per function (Excel limitation is 255, but framework currently supports 0-8)
- Supported types: f64, []const u8, *XLOPER12 - ranges are the big omission
- Windows x86_64 only, which makes sense as XLLs can only run on Windows
- Requires Zig 0.15.1 or later

## TODO

**Adding new parameter types:**

1. Update `extractArg()` in excel_function.zig
2. Add conversion logic using XLValue methods

**Adding new return types:**

1. Update `wrapResult()` in excel_function.zig
2. Add conversion logic using XLValue constructors

**Supporting more parameters:**

Add more cases to the switch statement in ExcelFunction().
