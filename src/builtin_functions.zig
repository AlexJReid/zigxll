const excel_function = @import("excel_function.zig");
const ExcelFunction = excel_function.ExcelFunction;
const ParamMeta = excel_function.ParamMeta;

pub const double = ExcelFunction(.{
    .name = "zigxll.Double",
    .description = "Doubles a number",
    .category = "Zig Functions",
    .params = &[_]ParamMeta{
        .{ .name = "x", .description = "Number to double" },
    },
    .func = doubleFunc,
});

fn doubleFunc(x: f64) !f64 {
    return x * 2;
}
