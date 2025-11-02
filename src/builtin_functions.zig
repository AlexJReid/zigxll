// Text manipulation functions for Excel
const std = @import("std");
const excel_function = @import("excel_function.zig");
const ExcelFunction = excel_function.ExcelFunction;
const ParamMeta = excel_function.ParamMeta;

const allocator = std.heap.c_allocator;

pub const double = ExcelFunction(.{
    .name = "zigxll.Double",
    .description = "Nonsense",
    .category = "Zig Functions",
    .params = &[_]ParamMeta{
        .{ .name = "x", .description = "Number to double" },
        .{ .name = "y", .description = "Number to add" },
        .{
            .name = "z",
        },
    },
    .func = doubleFunc,
});

fn doubleFunc(x: f64, y: f64, z: f64) !f64 {
    return x * 2 + y - z;
}

pub const zigmatrix = ExcelFunction(.{
    .name = "ZigMatrix",
    .description = "Returns a matrix filled with sequential numbers (max 100x100)",
    .category = "Zig Functions",
    .params = &[_]ParamMeta{
        .{ .name = "rows", .description = "Number of rows (max 100)" },
        .{ .name = "cols", .description = "Number of columns (max 100)" },
    },
    .func = zigMatrixFunc,
});

fn zigMatrixFunc(rows_param: f64, cols_param: f64) ![][]f64 {
    // Convert and validate parameters
    const rows_input = @as(i32, @intFromFloat(rows_param));
    const cols_input = @as(i32, @intFromFloat(cols_param));

    // Clamp to max 100, minimum 1
    const rows = @min(@max(rows_input, 1), 100);
    const cols = @min(@max(cols_input, 1), 100);

    const rows_usize = @as(usize, @intCast(rows));
    const cols_usize = @as(usize, @intCast(cols));

    // Allocate rows
    var matrix = try allocator.alloc([]f64, rows_usize);
    errdefer allocator.free(matrix);

    // Allocate and fill each row
    var row_idx: usize = 0;
    while (row_idx < rows_usize) : (row_idx += 1) {
        matrix[row_idx] = try allocator.alloc(f64, cols_usize);
        errdefer {
            // Clean up already allocated rows on error
            var i: usize = 0;
            while (i <= row_idx) : (i += 1) {
                allocator.free(matrix[i]);
            }
            allocator.free(matrix);
        }

        // Fill row with sequential values
        var col_idx: usize = 0;
        while (col_idx < cols_usize) : (col_idx += 1) {
            matrix[row_idx][col_idx] = @floatFromInt(row_idx * cols_usize + col_idx + 1);
        }
    }

    return matrix;
}

pub const not = ExcelFunction(.{
    .name = "ZigNot",
    .description = "Returns the logical NOT of a boolean value",
    .category = "Zig Functions",
    .params = &[_]ParamMeta{
        .{ .name = "value", .description = "Boolean value to invert" },
    },
    .func = notFunc,
});

fn notFunc(value: bool) !bool {
    return !value;
}
