const std = @import("std");

const xl_imports = @import("../xl_imports.zig");
const xl = xl_imports.xl;

// Excel uses 16-bit wide chars regardless of platform's native wchar_t
const ExcelWChar = u16;

pub const XLValue = struct {
    m_val: xl.XLOPER12,
    m_owns_memory: bool,
    allocator: std.mem.Allocator,

    // Constructors
    pub fn init(allocator: std.mem.Allocator) XLValue {
        return .{
            .m_val = undefined,
            .m_owns_memory = false,
            .allocator = allocator,
        };
    }

    pub fn fromXLOPER12(allocator: std.mem.Allocator, xloper: xl.XLOPER12, take_ownership: bool) XLValue {
        return .{
            .m_val = xloper,
            .m_owns_memory = take_ownership,
            .allocator = allocator,
        };
    }

    pub fn fromDouble(allocator: std.mem.Allocator, value: f64) XLValue {
        return .{
            .m_val = .{
                .xltype = xl.xltypeNum,
                .val = .{ .num = value },
            },
            .m_owns_memory = false,
            .allocator = allocator,
        };
    }

    pub fn fromInt(allocator: std.mem.Allocator, value: i32) XLValue {
        return fromDouble(allocator, @floatFromInt(value));
    }

    pub fn fromBool(allocator: std.mem.Allocator, value: bool) XLValue {
        return .{
            .m_val = .{
                .xltype = xl.xltypeBool,
                .val = .{ .xbool = if (value) 1 else 0 },
            },
            .m_owns_memory = false,
            .allocator = allocator,
        };
    }

    pub fn fromWString(allocator: std.mem.Allocator, str: []const u8) !XLValue {
        const len = str.len;
        var buf = try allocator.alloc(ExcelWChar, len + 2);
        buf[0] = @intCast(len);
        for (str, 0..) |ch, i| {
            buf[i + 1] = ch;
        }
        buf[len + 1] = 0;

        return .{
            .m_val = .{
                .xltype = xl.xltypeStr,
                .val = .{ .str = @ptrCast(buf.ptr) },
            },
            .m_owns_memory = true,
            .allocator = allocator,
        };
    }

    pub fn fromUtf8String(allocator: std.mem.Allocator, str: []const u8) !XLValue {
        // Calculate required UTF-16 length
        const utf16_len = try std.unicode.calcUtf16LeLen(str);

        // Allocate buffer: 1 wchar for length + utf16_len + 1 for null terminator
        var buf = try allocator.alloc(ExcelWChar, utf16_len + 2);
        errdefer allocator.free(buf);

        // Set length prefix
        buf[0] = @intCast(utf16_len);

        // Convert UTF-8 to UTF-16
        _ = try std.unicode.utf8ToUtf16Le(buf[1 .. utf16_len + 1], str);

        // Add null terminator
        buf[utf16_len + 1] = 0;

        return .{
            .m_val = .{
                .xltype = xl.xltypeStr,
                .val = .{ .str = @ptrCast(buf.ptr) },
            },
            .m_owns_memory = true,
            .allocator = allocator,
        };
    }

    pub fn fromMatrix(allocator: std.mem.Allocator, data: []const []const f64) !XLValue {
        const num_rows: i32 = @intCast(data.len);
        const num_cols: i32 = if (data.len > 0) @intCast(data[0].len) else 0;

        // Allocate multi array - be explicit about the type
        const array_size = @as(usize, @intCast(num_rows)) * @as(usize, @intCast(num_cols));
        var multi = try allocator.alloc(xl.XLOPER12, array_size);

        var idx: usize = 0;
        for (data) |row| {
            for (row) |val| {
                multi[idx] = .{
                    .xltype = xl.xltypeNum,
                    .val = .{ .num = val },
                };
                idx += 1;
            }
        }

        return .{
            .m_val = .{
                .xltype = xl.xltypeMulti,
                .val = .{ .array = .{
                    .lparray = multi.ptr,
                    .rows = num_rows,
                    .columns = num_cols,
                } },
            },
            .m_owns_memory = true,
            .allocator = allocator,
        };
    }

    pub fn err(allocator: std.mem.Allocator, err_code: c_int) XLValue {
        return .{
            .m_val = .{
                .xltype = xl.xltypeErr,
                .val = .{ .err = err_code },
            },
            .m_owns_memory = false,
            .allocator = allocator,
        };
    }

    pub fn missing(allocator: std.mem.Allocator) XLValue {
        return .{
            .m_val = .{
                .xltype = xl.xltypeMissing,
                .val = undefined,
            },
            .m_owns_memory = false,
            .allocator = allocator,
        };
    }

    // Type checking
    pub fn type_val(self: *const XLValue) xl.int {
        return self.m_val.xltype;
    }

    pub fn is_num(self: *const XLValue) bool {
        return self.m_val.xltype == xl.xltypeNum;
    }

    pub fn is_str(self: *const XLValue) bool {
        return self.m_val.xltype == xl.xltypeStr;
    }

    pub fn is_bool(self: *const XLValue) bool {
        return self.m_val.xltype == xl.xltypeBool;
    }

    pub fn is_err(self: *const XLValue) bool {
        return self.m_val.xltype == xl.xltypeErr;
    }

    pub fn is_multi(self: *const XLValue) bool {
        return self.m_val.xltype == xl.xltypeMulti;
    }

    pub fn is_missing(self: *const XLValue) bool {
        return self.m_val.xltype == xl.xltypeMissing;
    }

    pub fn is_nil(self: *const XLValue) bool {
        return self.m_val.xltype == xl.xltypeNil;
    }

    // Get values with type checking
    pub fn as_double(self: *const XLValue) !f64 {
        if (!self.is_num()) return error.NotANumber;
        return self.m_val.val.num;
    }

    pub fn as_int(self: *const XLValue) !i32 {
        const num = try self.as_double();
        return @intFromFloat(num);
    }

    pub fn as_bool(self: *const XLValue) !bool {
        if (!self.is_bool()) return error.NotABoolean;
        return self.m_val.val.xbool != 0;
    }

    pub fn as_wstring(self: *const XLValue) ![]u8 {
        if (!self.is_str()) return error.NotAString;
        const len = self.m_val.val.str[0];
        var result = try self.allocator.alloc(u8, len);
        for (0..len) |i| {
            result[i] = @intCast(self.m_val.val.str[i + 1]);
        }
        return result;
    }

    // Extract UTF-8 string from Excel's wide string XLOPER12
    // Excel stores strings as: first u16 = length, rest = wide chars
    // Returns allocated UTF-8 string that caller must free
    pub fn as_utf8str(self: *const XLValue) ![]u8 {
        if (!self.is_str()) return error.NotAString;

        const len = self.m_val.val.str[0];
        const wide_slice = self.m_val.val.str[1 .. len + 1];

        // Allocate buffer *3
        const buf = try self.allocator.alloc(u8, len * 3);
        errdefer self.allocator.free(buf);

        const utf8_len = try std.unicode.utf16LeToUtf8(buf, wide_slice);

        // Resize to actual length
        const result = try self.allocator.realloc(buf, utf8_len);
        return result;
    }

    // Matrix access
    pub fn rows(self: *const XLValue) usize {
        if (!self.is_multi()) return 0;
        return @intCast(self.m_val.val.array.rows);
    }

    pub fn columns(self: *const XLValue) usize {
        if (!self.is_multi()) return 0;
        return @intCast(self.m_val.val.array.columns);
    }

    pub fn get_cell(self: *const XLValue, row: usize, col: usize) !XLValue {
        if (!self.is_multi()) return error.NotAMatrix;
        if (row >= self.rows() or col >= self.columns()) return error.OutOfBounds;

        const idx = row * self.columns() + col;
        const cell = self.m_val.val.array.lparray[idx];

        return XLValue.fromXLOPER12(self.allocator, cell, false);
    }

    // Extract multi array as 2D array of f64
    // Caller must free the returned array and each row
    // Note: Empty/nil cells are converted to 0.0
    pub fn as_matrix(self: *const XLValue) ![][]const f64 {
        if (!self.is_multi()) return error.NotAMatrix;

        const num_rows = self.rows();
        const num_cols = self.columns();

        var result = try self.allocator.alloc([]const f64, num_rows);
        errdefer {
            for (result, 0..) |row, i| {
                if (i > 0) self.allocator.free(row);
            }
            self.allocator.free(result);
        }

        for (0..num_rows) |r| {
            var row = try self.allocator.alloc(f64, num_cols);
            errdefer if (r > 0) self.allocator.free(row);

            for (0..num_cols) |c| {
                const cell = try self.get_cell(r, c);

                // Handle different cell types - skip nil/missing/error cells
                if (cell.is_num()) {
                    row[c] = try cell.as_double();
                } else if (cell.is_nil() or cell.is_missing()) {
                    // Empty cells become 0.0
                    row[c] = 0.0;
                } else if (cell.is_err()) {
                    // Error cells become 0.0 (could alternatively propagate the error)
                    row[c] = 0.0;
                } else {
                    // Other types (string, bool) - try to convert or default to 0
                    row[c] = cell.as_double() catch 0.0;
                }
            }

            result[r] = row;
        }

        return result;
    }

    // Raw pointer access
    pub fn get(self: *XLValue) *xl.XLOPER12 {
        return &self.m_val;
    }

    pub fn get_const(self: *const XLValue) *const xl.XLOPER12 {
        return &self.m_val;
    }

    // Cleanup
    fn free_memory(self: *XLValue) void {
        if (!self.m_owns_memory) return;

        if (self.is_str()) {
            const str_ptr = self.m_val.val.str;
            const len = @as(usize, @intCast(str_ptr[0]));
            self.allocator.free(str_ptr[0 .. len + 2]); // Pascal string: 1 wchar for length prefix + len wchars for data + 1 null terminator
        } else if (self.is_multi()) {
            const rows_count = @as(usize, @intCast(self.m_val.val.array.rows));
            const cols_count = @as(usize, @intCast(self.m_val.val.array.columns));
            const total = rows_count * cols_count;
            self.allocator.free(self.m_val.val.array.lparray[0..total]);
        }

        self.m_owns_memory = false;
    }

    pub fn deinit(self: *XLValue) void {
        self.free_memory();
    }
};

pub fn format(
    self: *const XLValue,
    comptime fmt: []const u8,
    options: std.fmt.FormatOptions,
    writer: anytype,
) !void {
    _ = fmt;
    _ = options;

    if (self.is_num()) {
        try writer.print("XLValue(num: {})", .{self.m_val.val.num});
    } else if (self.is_str()) {
        try writer.print("XLValue(str)", .{});
    } else if (self.is_bool()) {
        try writer.print("XLValue(bool: {})", .{self.m_val.val.xbool != 0});
    } else {
        try writer.print("XLValue(type: {})", .{self.m_val.xltype});
    }
}

// Tests

test "XLValue basic types" {
    const allocator = std.testing.allocator;

    var val1 = XLValue.fromDouble(allocator, 42.0);
    defer val1.deinit();

    try std.testing.expect(val1.is_num());
    try std.testing.expectEqual(@as(f64, 42.0), try val1.as_double());

    var val2 = try XLValue.fromWString(allocator, "hello");
    defer val2.deinit();

    try std.testing.expect(val2.is_str());

    var val3 = XLValue.fromBool(allocator, true);
    defer val3.deinit();

    try std.testing.expect(val3.is_bool());
    try std.testing.expectEqual(true, try val3.as_bool());
}

test "XLValue matrix operations" {
    const allocator = std.testing.allocator;

    var matrix = try XLValue.fromMatrix(allocator, &.{
        &.{ 1.0, 2.0, 3.0 },
        &.{ 4.0, 5.0, 6.0 },
    });
    defer matrix.deinit();

    try std.testing.expect(matrix.is_multi());
    try std.testing.expectEqual(@as(usize, 2), matrix.rows());
    try std.testing.expectEqual(@as(usize, 3), matrix.columns());

    const cell = try matrix.get_cell(0, 1);
    const num = try cell.as_double();

    try std.testing.expectEqual(@as(f64, 2.0), num);
}

test "XLValue UTF-8 string round-trip" {
    const allocator = std.testing.allocator;

    // Test ASCII string
    {
        var val = try XLValue.fromUtf8String(allocator, "Hello");
        defer val.deinit();

        try std.testing.expect(val.is_str());
        const str = try val.as_utf8str();
        defer allocator.free(str);

        try std.testing.expectEqualStrings("Hello", str);
    }

    // Test empty string
    {
        var val = try XLValue.fromUtf8String(allocator, "");
        defer val.deinit();

        try std.testing.expect(val.is_str());
        const str = try val.as_utf8str();
        defer allocator.free(str);

        try std.testing.expectEqualStrings("", str);
    }

    // Test Unicode string (emoji, non-ASCII)
    {
        var val = try XLValue.fromUtf8String(allocator, "Hello 世界 🚀");
        defer val.deinit();

        try std.testing.expect(val.is_str());
        const str = try val.as_utf8str();
        defer allocator.free(str);

        try std.testing.expectEqualStrings("Hello 世界 🚀", str);
    }
}

test "XLValue UTF-8 string length calculation" {
    const allocator = std.testing.allocator;

    // ASCII: 1 byte per char
    {
        var val = try XLValue.fromUtf8String(allocator, "ABC");
        defer val.deinit();

        // Length should be stored in first wchar_t
        try std.testing.expectEqual(@as(u16, 3), val.m_val.val.str[0]);
    }

    // Multi-byte UTF-8 -> fewer UTF-16 code units
    {
        var val = try XLValue.fromUtf8String(allocator, "世界");
        defer val.deinit();

        // "世界" is 2 characters in UTF-16
        try std.testing.expectEqual(@as(u16, 2), val.m_val.val.str[0]);
    }
}

test "XLValue error types" {
    const allocator = std.testing.allocator;

    var val = XLValue.err(allocator, xl.xlerrValue);
    defer val.deinit();

    try std.testing.expect(val.is_err());
    try std.testing.expectEqual(@as(c_int, xl.xlerrValue), val.m_val.val.err);
}

test "XLValue type checking" {
    const allocator = std.testing.allocator;

    var num_val = XLValue.fromDouble(allocator, 42.0);
    defer num_val.deinit();

    // Should succeed
    _ = try num_val.as_double();

    // Should fail with type error
    try std.testing.expectError(error.NotAString, num_val.as_utf8str());
    try std.testing.expectError(error.NotABoolean, num_val.as_bool());
}

test "XLValue memory ownership" {
    const allocator = std.testing.allocator;

    // Numbers don't own memory
    {
        var val = XLValue.fromDouble(allocator, 123.0);
        defer val.deinit();
        try std.testing.expect(!val.m_owns_memory);
    }

    // Strings own memory
    {
        var val = try XLValue.fromUtf8String(allocator, "test");
        defer val.deinit();
        try std.testing.expect(val.m_owns_memory);
    }

    // Matrices own memory
    {
        var val = try XLValue.fromMatrix(allocator, &.{
            &.{ 1.0, 2.0 },
        });
        defer val.deinit();
        try std.testing.expect(val.m_owns_memory);
    }

    // Errors don't own memory
    {
        var val = XLValue.err(allocator, xl.xlerrValue);
        defer val.deinit();
        try std.testing.expect(!val.m_owns_memory);
    }
}

test "XLValue int conversion" {
    const allocator = std.testing.allocator;

    var val = XLValue.fromInt(allocator, 42);
    defer val.deinit();

    try std.testing.expect(val.is_num());
    try std.testing.expectEqual(@as(f64, 42.0), try val.as_double());
    try std.testing.expectEqual(@as(i32, 42), try val.as_int());
}

test "XLValue missing and nil types" {
    const allocator = std.testing.allocator;

    var missing_val = XLValue.missing(allocator);
    defer missing_val.deinit();
    try std.testing.expect(missing_val.is_missing());

    // Nil is different from missing in Excel
    var nil_val = XLValue.init(allocator);
    nil_val.m_val.xltype = xl.xltypeNil;
    defer nil_val.deinit();
    try std.testing.expect(nil_val.is_nil());
}

// Tests for fromXLOPER12 - simulating raw C structs from Excel

test "fromXLOPER12 with raw numeric" {
    const allocator = std.testing.allocator;

    const raw: xl.XLOPER12 = .{
        .xltype = xl.xltypeNum,
        .val = .{ .num = 99.5 },
    };

    var val = XLValue.fromXLOPER12(allocator, raw, false);
    defer val.deinit();

    try std.testing.expect(val.is_num());
    try std.testing.expect(!val.m_owns_memory);
    try std.testing.expectEqual(@as(f64, 99.5), try val.as_double());
}

test "fromXLOPER12 with raw bool" {
    const allocator = std.testing.allocator;

    const raw_true: xl.XLOPER12 = .{
        .xltype = xl.xltypeBool,
        .val = .{ .xbool = 1 },
    };

    var val_true = XLValue.fromXLOPER12(allocator, raw_true, false);
    defer val_true.deinit();

    try std.testing.expect(val_true.is_bool());
    try std.testing.expectEqual(true, try val_true.as_bool());

    const raw_false: xl.XLOPER12 = .{
        .xltype = xl.xltypeBool,
        .val = .{ .xbool = 0 },
    };

    var val_false = XLValue.fromXLOPER12(allocator, raw_false, false);
    defer val_false.deinit();

    try std.testing.expectEqual(false, try val_false.as_bool());
}

test "fromXLOPER12 with raw error" {
    const allocator = std.testing.allocator;

    const raw: xl.XLOPER12 = .{
        .xltype = xl.xltypeErr,
        .val = .{ .err = xl.xlerrDiv0 },
    };

    var val = XLValue.fromXLOPER12(allocator, raw, false);
    defer val.deinit();

    try std.testing.expect(val.is_err());
    try std.testing.expectEqual(@as(c_int, xl.xlerrDiv0), val.m_val.val.err);
}

test "fromXLOPER12 with raw string (no ownership)" {
    const allocator = std.testing.allocator;

    var str_buf = [_]u16{ 2, 'H', 'i', 0 };

    const raw: xl.XLOPER12 = .{
        .xltype = xl.xltypeStr,
        .val = .{ .str = &str_buf },
    };

    var val = XLValue.fromXLOPER12(allocator, raw, false);
    defer val.deinit();

    try std.testing.expect(val.is_str());
    try std.testing.expect(!val.m_owns_memory);

    const str = try val.as_utf8str();
    defer allocator.free(str);

    try std.testing.expectEqualStrings("Hi", str);
}

test "fromXLOPER12 with raw string (with ownership)" {
    const allocator = std.testing.allocator;

    var str_buf = try allocator.alloc(u16, 6);
    str_buf[0] = 4;
    str_buf[1] = 'T';
    str_buf[2] = 'e';
    str_buf[3] = 's';
    str_buf[4] = 't';
    str_buf[5] = 0;

    const raw: xl.XLOPER12 = .{
        .xltype = xl.xltypeStr,
        .val = .{ .str = str_buf.ptr },
    };

    var val = XLValue.fromXLOPER12(allocator, raw, true);
    defer val.deinit();

    try std.testing.expect(val.is_str());
    try std.testing.expect(val.m_owns_memory);

    const str = try val.as_utf8str();
    defer allocator.free(str);

    try std.testing.expectEqualStrings("Test", str);
}

test "fromXLOPER12 with raw multi array (no ownership)" {
    const allocator = std.testing.allocator;

    var cells = [_]xl.XLOPER12{
        .{ .xltype = xl.xltypeNum, .val = .{ .num = 1.0 } },
        .{ .xltype = xl.xltypeNum, .val = .{ .num = 2.0 } },
        .{ .xltype = xl.xltypeNum, .val = .{ .num = 3.0 } },
        .{ .xltype = xl.xltypeNum, .val = .{ .num = 4.0 } },
    };

    const raw: xl.XLOPER12 = .{
        .xltype = xl.xltypeMulti,
        .val = .{ .array = .{
            .lparray = &cells,
            .rows = 2,
            .columns = 2,
        } },
    };

    var val = XLValue.fromXLOPER12(allocator, raw, false);
    defer val.deinit();

    try std.testing.expect(val.is_multi());
    try std.testing.expect(!val.m_owns_memory);
    try std.testing.expectEqual(@as(usize, 2), val.rows());
    try std.testing.expectEqual(@as(usize, 2), val.columns());

    const cell_00 = try val.get_cell(0, 0);
    try std.testing.expectEqual(@as(f64, 1.0), try cell_00.as_double());

    const cell_01 = try val.get_cell(0, 1);
    try std.testing.expectEqual(@as(f64, 2.0), try cell_01.as_double());

    const cell_10 = try val.get_cell(1, 0);
    try std.testing.expectEqual(@as(f64, 3.0), try cell_10.as_double());

    const cell_11 = try val.get_cell(1, 1);
    try std.testing.expectEqual(@as(f64, 4.0), try cell_11.as_double());
}

test "fromXLOPER12 with raw multi array (with ownership)" {
    const allocator = std.testing.allocator;

    var cells = try allocator.alloc(xl.XLOPER12, 4);
    cells[0] = .{ .xltype = xl.xltypeNum, .val = .{ .num = 10.0 } };
    cells[1] = .{ .xltype = xl.xltypeNum, .val = .{ .num = 20.0 } };
    cells[2] = .{ .xltype = xl.xltypeNum, .val = .{ .num = 30.0 } };
    cells[3] = .{ .xltype = xl.xltypeNum, .val = .{ .num = 40.0 } };

    const raw: xl.XLOPER12 = .{
        .xltype = xl.xltypeMulti,
        .val = .{ .array = .{
            .lparray = cells.ptr,
            .rows = 2,
            .columns = 2,
        } },
    };

    var val = XLValue.fromXLOPER12(allocator, raw, true);
    defer val.deinit();

    try std.testing.expect(val.is_multi());
    try std.testing.expect(val.m_owns_memory);

    const cell = try val.get_cell(1, 1);
    try std.testing.expectEqual(@as(f64, 40.0), try cell.as_double());
}

test "fromXLOPER12 with mixed type array" {
    const allocator = std.testing.allocator;

    var str_buf = [_]u16{ 3, 'a', 'b', 'c', 0 };

    var cells = [_]xl.XLOPER12{
        .{ .xltype = xl.xltypeNum, .val = .{ .num = 42.0 } },
        .{ .xltype = xl.xltypeStr, .val = .{ .str = &str_buf } },
        .{ .xltype = xl.xltypeBool, .val = .{ .xbool = 1 } },
        .{ .xltype = xl.xltypeErr, .val = .{ .err = xl.xlerrNA } },
    };

    const raw: xl.XLOPER12 = .{
        .xltype = xl.xltypeMulti,
        .val = .{ .array = .{
            .lparray = &cells,
            .rows = 2,
            .columns = 2,
        } },
    };

    var val = XLValue.fromXLOPER12(allocator, raw, false);
    defer val.deinit();

    const cell_num = try val.get_cell(0, 0);
    try std.testing.expect(cell_num.is_num());
    try std.testing.expectEqual(@as(f64, 42.0), try cell_num.as_double());

    const cell_str = try val.get_cell(0, 1);
    try std.testing.expect(cell_str.is_str());
    const str = try cell_str.as_utf8str();
    defer allocator.free(str);
    try std.testing.expectEqualStrings("abc", str);

    const cell_bool = try val.get_cell(1, 0);
    try std.testing.expect(cell_bool.is_bool());
    try std.testing.expectEqual(true, try cell_bool.as_bool());

    const cell_err = try val.get_cell(1, 1);
    try std.testing.expect(cell_err.is_err());
    try std.testing.expectEqual(@as(c_int, xl.xlerrNA), cell_err.m_val.val.err);
}

test "XLValue matrix round-trip with as_matrix" {
    const allocator = std.testing.allocator;

    const input_data = &[_][]const f64{
        &.{ 1.0, 2.0, 3.0 },
        &.{ 4.0, 5.0, 6.0 },
        &.{ 7.0, 8.0, 9.0 },
    };

    var matrix = try XLValue.fromMatrix(allocator, input_data);
    defer matrix.deinit();

    try std.testing.expect(matrix.is_multi());
    try std.testing.expectEqual(@as(usize, 3), matrix.rows());
    try std.testing.expectEqual(@as(usize, 3), matrix.columns());

    const result = try matrix.as_matrix();
    defer {
        for (result) |row| {
            allocator.free(row);
        }
        allocator.free(result);
    }

    try std.testing.expectEqual(@as(usize, 3), result.len);
    try std.testing.expectEqual(@as(usize, 3), result[0].len);

    try std.testing.expectEqual(@as(f64, 1.0), result[0][0]);
    try std.testing.expectEqual(@as(f64, 2.0), result[0][1]);
    try std.testing.expectEqual(@as(f64, 3.0), result[0][2]);
    try std.testing.expectEqual(@as(f64, 4.0), result[1][0]);
    try std.testing.expectEqual(@as(f64, 5.0), result[1][1]);
    try std.testing.expectEqual(@as(f64, 6.0), result[1][2]);
    try std.testing.expectEqual(@as(f64, 7.0), result[2][0]);
    try std.testing.expectEqual(@as(f64, 8.0), result[2][1]);
    try std.testing.expectEqual(@as(f64, 9.0), result[2][2]);
}

test "XLValue as_matrix edge cases" {
    const allocator = std.testing.allocator;

    {
        var matrix = try XLValue.fromMatrix(allocator, &.{
            &.{42.0},
        });
        defer matrix.deinit();

        const result = try matrix.as_matrix();
        defer {
            for (result) |row| {
                allocator.free(row);
            }
            allocator.free(result);
        }

        try std.testing.expectEqual(@as(usize, 1), result.len);
        try std.testing.expectEqual(@as(usize, 1), result[0].len);
        try std.testing.expectEqual(@as(f64, 42.0), result[0][0]);
    }

    {
        var matrix = try XLValue.fromMatrix(allocator, &.{
            &.{ 1.0, 2.0, 3.0, 4.0 },
        });
        defer matrix.deinit();

        const result = try matrix.as_matrix();
        defer {
            for (result) |row| {
                allocator.free(row);
            }
            allocator.free(result);
        }

        try std.testing.expectEqual(@as(usize, 1), result.len);
        try std.testing.expectEqual(@as(usize, 4), result[0].len);
        try std.testing.expectEqual(@as(f64, 1.0), result[0][0]);
        try std.testing.expectEqual(@as(f64, 4.0), result[0][3]);
    }

    {
        var matrix = try XLValue.fromMatrix(allocator, &.{
            &.{1.0},
            &.{2.0},
            &.{3.0},
        });
        defer matrix.deinit();

        const result = try matrix.as_matrix();
        defer {
            for (result) |row| {
                allocator.free(row);
            }
            allocator.free(result);
        }

        try std.testing.expectEqual(@as(usize, 3), result.len);
        try std.testing.expectEqual(@as(usize, 1), result[0].len);
        try std.testing.expectEqual(@as(f64, 1.0), result[0][0]);
        try std.testing.expectEqual(@as(f64, 2.0), result[1][0]);
        try std.testing.expectEqual(@as(f64, 3.0), result[2][0]);
    }

    {
        var num_val = XLValue.fromDouble(allocator, 42.0);
        defer num_val.deinit();

        try std.testing.expectError(error.NotAMatrix, num_val.as_matrix());
    }
}
