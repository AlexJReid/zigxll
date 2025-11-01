// Text manipulation functions for Excel
const std = @import("std");
const excel_function = @import("excel_function.zig");
const ExcelFunction = excel_function.ExcelFunction;
const ParamMeta = excel_function.ParamMeta;

const allocator = std.heap.c_allocator;

pub const uppercase = ExcelFunction(.{
    .name = "uppercase",
    .description = "Convert text to uppercase",
    .category = "Text",
    .params = &[_]ParamMeta{
        .{ .name = "text", .description = "Text to convert" },
    },
    .func = uppercaseFunc,
});

fn uppercaseFunc(text: []const u8) ![]const u8 {
    const result = try allocator.alloc(u8, text.len);
    for (text, 0..) |c, i| {
        result[i] = std.ascii.toUpper(c);
    }
    return result;
}

pub const concat = ExcelFunction(.{
    .name = "concat",
    .description = "Concatenate two strings",
    .category = "Text",
    .params = &[_]ParamMeta{
        .{ .name = "text1", .description = "First text" },
        .{ .name = "text2", .description = "Second text" },
    },
    .func = concatFunc,
});

fn concatFunc(text1: []const u8, text2: []const u8) ![]const u8 {
    return std.fmt.allocPrint(allocator, "{s}{s}", .{ text1, text2 });
}
