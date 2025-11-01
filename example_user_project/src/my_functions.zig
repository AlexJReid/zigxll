// My custom Excel functions
const std = @import("std");
const xll = @import("xll");
const ExcelFunction = xll.ExcelFunction;
const ParamMeta = xll.ParamMeta;

const allocator = std.heap.c_allocator;

// Example custom function
pub const double = ExcelFunction(.{
    .name = "double",
    .description = "Double a number",
    .category = "Zig Functions",
    .params = &[_]ParamMeta{
        .{ .name = "x", .description = "Number to double" },
    },
    .func = doubleFunc,
});

fn doubleFunc(x: f64) !f64 {
    return x * 2;
}

pub const reverse = ExcelFunction(.{
    .name = "reverse",
    .description = "Reverse a string",
    .category = "Zig Functions",
    .params = &[_]ParamMeta{
        .{ .name = "text", .description = "Text to reverse" },
    },
    .func = reverseFunc,
});

fn reverseFunc(text: []const u8) ![]const u8 {
    var result = try allocator.alloc(u8, text.len);
    for (text, 0..) |c, i| {
        result[text.len - 1 - i] = c;
    }
    return result;
}
