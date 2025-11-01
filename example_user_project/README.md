# Example ZigXLL User Project

This is an example of how to use ZigXLL as a library dependency.

## Structure

```
src/
├── main.zig              ← Entry point (re-exports zigxll framework)
├── user_functions.zig    ← Register your function modules here
└── my_functions.zig      ← Your custom functions
```

## Building

```bash
zig build
```

Your XLL will be in `zig-out/lib/my_excel_functions.xll`

## Adding Functions

1. Create a new module in `src/` (e.g., `src/stats.zig`)
2. Add it to `src/user_functions.zig`:
   ```zig
   pub const stats = @import("stats");
   pub const function_modules = .{
       my_functions,
       stats,
   };
   ```
3. Rebuild: `zig build`

## Function Template

```zig
const zigxll = @import("zigxll");
const ExcelFunction = zigxll.@"src/excel_function.zig".ExcelFunction;
const ParamMeta = zigxll.@"src/excel_function.zig".ParamMeta;

pub const my_func = ExcelFunction(.{
    .name = "my_func",
    .description = "Does something cool",
    .category = "MyCategory",
    .params = &[_]ParamMeta{
        .{ .name = "x", .description = "First parameter" },
    },
    .func = myFuncImpl,
});

fn myFuncImpl(x: f64) !f64 {
    return x + 1;
}
```
