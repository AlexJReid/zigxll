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

See [rtd-servers.md](./rtd-servers.md) for more on RTD.

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

## Limits

- Maximum 8 parameters per function (Excel supports 255, framework currently caps at 8).
- Functions must use the C allocator (`std.heap.c_allocator`) for returned strings and arrays.
