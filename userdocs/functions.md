# Creating Excel Functions

## Quick start

Define functions using `ExcelFunction()` and wire them up in your `main.zig`.

**`src/my_functions.zig`:**

```zig
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

**`src/main.zig`:**

```zig
pub const function_modules = .{
    @import("my_functions.zig"),
};
```

That's it. The framework discovers all `ExcelFunction` definitions at compile time, generates C-callable wrappers, and registers them with Excel when the XLL loads.

## ExcelFunction options

```zig
pub const myFunc = ExcelFunction(.{
    .name = "MyCategory.MyFunc",  // Dots are fine — namespaces the function
    .description = "What it does", // Shows in Excel's function wizard
    .category = "My Category",     // Groups functions in the wizard
    .params = &[_]ParamMeta{ ... },
    .func = myFuncImpl,
    .thread_safe = true,           // Default: true (enables MTR)
});
```

| Field | Required | Default | Notes |
|---|---|---|---|
| `name` | yes | — | Function name as it appears in Excel. Dots are allowed for namespacing (e.g. `"Finance.BSCall"`). |
| `func` | yes | — | The Zig function to wrap. |
| `params` | no | `&.{}` | Array of `ParamMeta` structs. Must match function arity if provided. |
| `description` | no | `""` | Shown in Excel's Insert Function dialog. |
| `category` | no | `"General"` | Groups the function in Excel's function list. |
| `thread_safe` | no | `true` | Enables Multi-Threaded Recalculation. Set to `false` if your function has side effects or shared state. |
| `is_async` | no | `false` | Runs the function on a background thread pool with result caching via RTD. Automatically sets `thread_safe = false`. |

## Supported types

### Parameters

| Zig type | Excel type | Notes |
|---|---|---|
| `f64` | Number | |
| `bool` | Boolean | TRUE/FALSE |
| `[]const u8` | String | Framework handles UTF-16 to UTF-8 conversion |
| `[][]const f64` | Range | 2D array of numbers. Empty cells become 0.0. |
| `*XLOPER12` | Any | Raw Excel value — for advanced use |

### Optional parameters

Use `?T` types for optional parameters. When Excel passes a missing value, it becomes `null`:

```zig
fn powerImpl(base: f64, exponent: ?f64) !f64 {
    const exp = exponent orelse 2.0;
    return std.math.pow(f64, base, exp);
}
```

Supported optional types: `?f64`, `?bool`, `?[]const u8`.

### Return types

| Zig type | Excel type | Notes |
|---|---|---|
| `f64` | Number | |
| `bool` | Boolean | |
| `[]const u8` / `[]u8` | String | Freed by Excel via `xlAutoFree12` |
| `[][]const f64` / `[][]f64` | Array | Spills into a range. Freed by Excel. |
| `*XLOPER12` | Any | Raw — you manage memory |

All return types are wrapped with `!T`. Errors automatically become `#VALUE!` in Excel.

### Returning specific Excel errors

For functions returning `!*xl.XLOPER12`, use the static error helpers on `XLValue`:

```zig
const XLValue = xll.XLValue;

fn safeDivideImpl(a: f64, b: f64) !*xl.XLOPER12 {
    if (b == 0) return XLValue.errDiv0();
    // ... normal return via wrapResult handled by framework
}
```

| Helper | Excel error |
|---|---|
| `XLValue.na()` | `#N/A` |
| `XLValue.errValue()` | `#VALUE!` |
| `XLValue.errDiv0()` | `#DIV/0!` |
| `XLValue.errRef()` | `#REF!` |
| `XLValue.errName()` | `#NAME?` |
| `XLValue.errNum()` | `#NUM!` |
| `XLValue.errNull()` | `#NULL!` |

These are static singletons — no allocation, safe to return from any code path.

## Returning strings

Allocate the result with `std.heap.c_allocator`. The framework marks it with `xlbitDLLFree` and Excel calls `xlAutoFree12` to free it:

```zig
const allocator = std.heap.c_allocator;

fn reverseImpl(text: []const u8) ![]const u8 {
    var result = try allocator.alloc(u8, text.len);
    for (text, 0..) |c, i| {
        result[text.len - 1 - i] = c;
    }
    return result;
}
```

## Returning arrays

Return a `[][]f64` or `[][]const f64`. Excel spills the result into adjacent cells:

```zig
fn matrixImpl(rows: ?f64, cols: ?f64) ![][]f64 {
    const r: usize = @intFromFloat(rows orelse 3);
    const c: usize = @intFromFloat(cols orelse 3);

    var matrix = try allocator.alloc([]f64, r);
    for (0..r) |i| {
        matrix[i] = try allocator.alloc(f64, c);
        for (0..c) |j| {
            matrix[i][j] = @floatFromInt(i * c + j + 1);
        }
    }
    return matrix;
}
```

## Returning raw XLOPER12

For advanced use (e.g. wrapping RTD calls), return `*xl.XLOPER12`:

```zig
const xl = xll.xl;
const rtd_call = xll.rtd_call;

fn livePriceImpl(symbol: []const u8) !*xl.XLOPER12 {
    return rtd_call.subscribe("myprog.rtd", &.{symbol});
}
```

Functions that call `rtd_call.subscribe()` must set `.thread_safe = false` — `xlfRtd` must run on Excel's main thread. See [rtd-servers.md](./rtd-servers.md) for more on RTD.

## Multiple modules

Split functions across files and list them all in `main.zig`:

```zig
pub const function_modules = .{
    @import("math_functions.zig"),
    @import("string_functions.zig"),
    @import("finance_functions.zig"),
};
```

Each module is scanned independently. Functions can use any name — there's no conflict between modules.

## Namespacing

Use dots in the function name to namespace it in Excel:

```zig
.name = "Finance.BSCall",
```

Excel shows this as `Finance.BSCall`. The exported DLL symbol uses underscores (`Finance_BSCall_impl`) since Windows `GetProcAddress` doesn't support dots.

## Async functions

Add `.is_async = true` to run a function on a background thread pool. The cell shows `#N/A` while computing, then updates with the final result. Once complete, the cell becomes a plain value (no ongoing overhead).

```zig
pub const slow_calc = ExcelFunction(.{
    .name = "SlowCalc",
    .description = "Expensive calculation",
    .is_async = true,
    .func = slowCalcImpl,
    .params = &[_]ParamMeta{
        .{ .name = "x", .description = "Input value" },
    },
});

fn slowCalcImpl(x: f64) !f64 {
    // This runs on a background thread — Excel stays responsive.
    doExpensiveWork();
    return x * 2.0;
}
```

The function signature is identical to a sync function. The framework handles all RTD plumbing, caching, and thread management automatically.

### How it works

1. First call → cache miss → spawns work on thread pool → cell shows `#N/A`
2. Worker finishes → result cached → Excel recalculates
3. Next recalc → cache hit → returns value directly → RTD subscription dropped
4. Subsequent calls with same args → instant cache hit (no re-computation)

### Intermediate values

To send progress updates to the cell before the final result, add `*AsyncContext` as the last parameter:

```zig
const AsyncContext = xll.AsyncContext;

pub const slow_calc = ExcelFunction(.{
    .name = "SlowCalc",
    .description = "Expensive calculation with progress",
    .is_async = true,
    .func = slowCalcImpl,
    .params = &[_]ParamMeta{
        .{ .name = "x", .description = "Input value" },
    },
});

fn slowCalcImpl(x: f64, ctx: *AsyncContext) !f64 {
    ctx.yield(.{ .string = "Computing..." });     // cell updates immediately
    doFirstPhase();

    ctx.yield(.{ .double = x * 0.5 });            // partial result
    doSecondPhase();

    ctx.yield(.{ .string = "Finalizing..." });
    doFinalPhase();

    return x * 2.0;  // final value — cell becomes a plain value cell
}
```

The `*AsyncContext` parameter is invisible to Excel — it is not counted as a function parameter and doesn't need a `ParamMeta` entry. Excel sees `=SlowCalc(42)` as a 1-parameter function.

`ctx.yield()` accepts an `AsyncValue`:

| Variant | Example |
|---|---|
| `.int` | `ctx.yield(.{ .int = 42 })` |
| `.double` | `ctx.yield(.{ .double = 3.14 })` |
| `.string` | `ctx.yield(.{ .string = "Loading..." })` |
| `.boolean` | `ctx.yield(.{ .boolean = true })` |

Each `yield` updates the cell immediately. When the function returns, the final value replaces the last yielded value and the cell stops being an RTD cell.

### Caching

Results are cached by function name and arguments. Two calls to `=SlowCalc(42)` in different cells share the same cached result — the computation runs only once. The cache persists for the lifetime of the XLL (until Excel closes or the add-in is unloaded).

To let users force a recalculation, expose a macro that clears the cache:

```zig
const xll = @import("xll");
const ExcelMacro = xll.ExcelMacro;

pub const clear_cache = ExcelMacro(.{
    .name = "ClearAsyncCache",
    .description = "Clear cached async results, forcing recalculation",
    .func = struct {
        fn f() void {
            xll.async_cache.getGlobalCache().clear();
        }
    }.f,
});
```

After running this macro (Alt+F8 → `ClearAsyncCache`), the next recalc will re-execute all async functions.

### Thread pool

Async functions run on a shared thread pool (4 workers). If all workers are busy, new tasks queue until a worker is free. The pool is created lazily on the first async call.

### Interaction with thread_safe

Async functions are always registered as non-thread-safe (`thread_safe` is forced to `false`). This is because the initial call uses `xlfRtd` which must run on Excel's main thread. The actual computation runs on the thread pool regardless.

## Macros (commands)

Excel macros are commands that perform actions (show dialogs, modify cells, etc.) rather than returning values. Unlike worksheet functions, macros can call Excel command-equivalent C API functions like `xlcAlert`, `xlcSelect`, etc.

Define macros using `ExcelMacro()`:

```zig
const xll = @import("xll");
const xl = xll.xl;
const XLValue = xll.XLValue;
const ExcelMacro = xll.ExcelMacro;

const allocator = @import("std").heap.c_allocator;

pub const hello = ExcelMacro(.{
    .name = "MyAddin.Hello",
    .description = "Show a greeting",
    .category = "My Macros",
    .func = helloImpl,
});

fn helloImpl() !void {
    var msg = try XLValue.fromUtf8String(allocator, "Hello from Zig!");
    defer msg.deinit();
    _ = xl.Excel12f(xl.xlcAlert, null, 1, &msg.m_val);
}
```

### ExcelMacro options

| Field | Required | Default | Notes |
|---|---|---|---|
| `name` | yes | — | Macro name as it appears in Excel. Dots allowed for namespacing. |
| `func` | yes | — | Must be `fn () !void` or `fn () void`. |
| `description` | no | `""` | Shown in Excel's macro dialog. |
| `category` | no | `"General"` | Groups the macro in Excel's UI. |

Macros are discovered alongside functions from the same `function_modules` tuple — no separate wiring needed.

### Inline functions

For simple macros (or functions), you can skip the separate `fn` declaration and inline it directly using an anonymous struct:

```zig
pub const hello = ExcelMacro(.{
    .name = "Hello",
    .description = "Show a greeting",
    .func = struct {
        fn f() !void {
            var msg = try XLValue.fromUtf8String(allocator, "Hello!");
            defer msg.deinit();
            _ = xl.Excel12f(xl.xlcAlert, null, 1, &msg.m_val);
        }
    }.f,
});
```

This works for `ExcelFunction` too — useful when the implementation is short and doesn't need to be referenced elsewhere.

### Running macros

Macros do **not** appear in Excel's macro dialog. They can be run via VBA `Application.Run "MacroName"`

## Limits

- Maximum 8 parameters per function (Excel supports 255, framework currently caps at 8).
- The `*AsyncContext` parameter (if used) does not count toward the 8-parameter limit.
- Functions must use the C allocator (`std.heap.c_allocator`) for returned strings and arrays.
