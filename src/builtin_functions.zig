// Text manipulation functions for Excel
const std = @import("std");
const excel_function = @import("excel_function.zig");
const ExcelFunction = excel_function.ExcelFunction;
const ParamMeta = excel_function.ParamMeta;

const allocator = std.heap.c_allocator;

pub const double = ExcelFunction(.{
    .name = "zigxll.Double",
    .description = "Double a number",
    .category = "Zig Functions",
    .params = &[_]ParamMeta{
        .{ .name = "x", .description = "Number to double" },
        .{ .name = "y", .description = "Number to add" },
        .{ .name = "z", .description = "Number to take off" },
    },
    .func = doubleFunc,
});

fn doubleFunc(x: f64, y: f64, z: f64) !f64 {
    return x * 2 + y - z;
}
