const std = @import("std");

const xl_imports = @import("xl_imports.zig");
const win = xl_imports.win;
const xl = xl_imports.xl;

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
        var buf = try allocator.alloc(xl.wchar_t, len + 2);
        buf[0] = @intCast(len);
        for (str, 0..) |ch, i| {
            buf[i + 1] = ch;
        }
        buf[len + 1] = 0;

        return .{
            .m_val = .{
                .xltype = xl.xltypeStr,
                .val = .{ .str = buf.ptr },
            },
            .m_owns_memory = true,
            .allocator = allocator,
        };
    }

    pub fn fromUtf8String(allocator: std.mem.Allocator, str: []const u8) !XLValue {
        // Calculate required UTF-16 length
        const utf16_len = try std.unicode.calcUtf16LeLen(str);

        // Allocate buffer: 1 wchar_t for length + utf16_len + 1 for null terminator
        var buf = try allocator.alloc(xl.wchar_t, utf16_len + 2);
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
                .val = .{ .str = buf.ptr },
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
    pub fn as_matrix(self: *const XLValue) ![][]f64 {
        if (!self.is_multi()) return error.NotAMatrix;

        const num_rows = self.rows();
        const num_cols = self.columns();

        var result = try self.allocator.alloc([]f64, num_rows);
        errdefer {
            for (result, 0..) |row, i| {
                if (i > 0) self.allocator.free(row);
            }
            self.allocator.free(result);
        }

        for (0..num_rows) |r| {
            result[r] = try self.allocator.alloc(f64, num_cols);
            errdefer if (r > 0) self.allocator.free(result[r]);

            for (0..num_cols) |c| {
                const cell = try self.get_cell(r, c);
                result[r][c] = try cell.as_double();
            }
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

// The tests that follow have been proposed by Claude, they need sense checking of course.
// Note that the testing allocator used will complain if there are suspected leaks.

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

    // There is no spoon?
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

test "XLValue matrix round-trip with as_matrix" {
    const allocator = std.testing.allocator;

    // Create a matrix from 2D array
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

    // Extract it back as a 2D array
    const result = try matrix.as_matrix();
    defer {
        for (result) |row| {
            allocator.free(row);
        }
        allocator.free(result);
    }

    // Verify dimensions
    try std.testing.expectEqual(@as(usize, 3), result.len);
    try std.testing.expectEqual(@as(usize, 3), result[0].len);

    // Verify all values match
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

    // Test single cell (1x1 matrix)
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

    // Test single row
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

    // Test single column
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

    // Test error on non-matrix type
    {
        var num_val = XLValue.fromDouble(allocator, 42.0);
        defer num_val.deinit();

        try std.testing.expectError(error.NotAMatrix, num_val.as_matrix());
    }
}

test "XLOPER12 round-trip for numeric type" {
    const allocator = std.testing.allocator;

    // Create XLValue from number
    var val1 = XLValue.fromDouble(allocator, 123.456);
    defer val1.deinit();

    // Get raw XLOPER12
    const xloper = val1.m_val;

    // Verify XLOPER12 structure directly
    try std.testing.expectEqual(@as(@TypeOf(xloper.xltype), xl.xltypeNum), xloper.xltype);
    try std.testing.expectEqual(@as(f64, 123.456), xloper.val.num);

    // Create new XLValue from XLOPER12
    const val2 = XLValue.fromXLOPER12(allocator, xloper, false);

    // Extract and verify
    try std.testing.expect(val2.is_num());
    try std.testing.expectEqual(@as(f64, 123.456), try val2.as_double());
}

test "XLOPER12 round-trip for boolean type" {
    const allocator = std.testing.allocator;

    // Test true
    {
        var val1 = XLValue.fromBool(allocator, true);
        defer val1.deinit();

        const xloper = val1.m_val;
        try std.testing.expectEqual(@as(@TypeOf(xloper.xltype), xl.xltypeBool), xloper.xltype);
        try std.testing.expectEqual(@as(c_int, 1), xloper.val.xbool);

        const val2 = XLValue.fromXLOPER12(allocator, xloper, false);
        try std.testing.expect(val2.is_bool());
        try std.testing.expectEqual(true, try val2.as_bool());
    }

    // Test false
    {
        var val1 = XLValue.fromBool(allocator, false);
        defer val1.deinit();

        const xloper = val1.m_val;
        try std.testing.expectEqual(@as(@TypeOf(xloper.xltype), xl.xltypeBool), xloper.xltype);
        try std.testing.expectEqual(@as(c_int, 0), xloper.val.xbool);

        const val2 = XLValue.fromXLOPER12(allocator, xloper, false);
        try std.testing.expect(val2.is_bool());
        try std.testing.expectEqual(false, try val2.as_bool());
    }
}

test "XLOPER12 round-trip for string type" {
    const allocator = std.testing.allocator;

    // Create XLValue from UTF-8 string
    var val1 = try XLValue.fromUtf8String(allocator, "Hello Excel!");
    defer val1.deinit();

    // Get raw XLOPER12
    const xloper = val1.m_val;

    // Verify XLOPER12 structure
    try std.testing.expectEqual(@as(@TypeOf(xloper.xltype), xl.xltypeStr), xloper.xltype);
    try std.testing.expectEqual(@as(u16, 12), xloper.val.str[0]); // Length

    // Create new XLValue from XLOPER12 (not taking ownership)
    const val2 = XLValue.fromXLOPER12(allocator, xloper, false);

    // Extract and verify
    try std.testing.expect(val2.is_str());
    const str = try val2.as_utf8str();
    defer allocator.free(str);
    try std.testing.expectEqualStrings("Hello Excel!", str);
}

test "XLOPER12 round-trip for multi/matrix type" {
    const allocator = std.testing.allocator;

    // Create XLValue from matrix
    const input_data = &[_][]const f64{
        &.{ 1.0, 2.0 },
        &.{ 3.0, 4.0 },
    };
    var val1 = try XLValue.fromMatrix(allocator, input_data);
    defer val1.deinit();

    // Get raw XLOPER12
    const xloper = val1.m_val;

    // Verify XLOPER12 structure
    try std.testing.expectEqual(@as(@TypeOf(xloper.xltype), xl.xltypeMulti), xloper.xltype);
    try std.testing.expectEqual(@as(i32, 2), xloper.val.array.rows);
    try std.testing.expectEqual(@as(i32, 2), xloper.val.array.columns);

    // Verify individual cells in the raw array
    try std.testing.expectEqual(@as(@TypeOf(xloper.val.array.lparray[0].xltype), xl.xltypeNum), xloper.val.array.lparray[0].xltype);
    try std.testing.expectEqual(@as(f64, 1.0), xloper.val.array.lparray[0].val.num);
    try std.testing.expectEqual(@as(f64, 2.0), xloper.val.array.lparray[1].val.num);
    try std.testing.expectEqual(@as(f64, 3.0), xloper.val.array.lparray[2].val.num);
    try std.testing.expectEqual(@as(f64, 4.0), xloper.val.array.lparray[3].val.num);

    // Create new XLValue from XLOPER12 (not taking ownership)
    const val2 = XLValue.fromXLOPER12(allocator, xloper, false);

    // Verify type and dimensions
    try std.testing.expect(val2.is_multi());
    try std.testing.expectEqual(@as(usize, 2), val2.rows());
    try std.testing.expectEqual(@as(usize, 2), val2.columns());

    // Extract as matrix and verify values
    const result = try val2.as_matrix();
    defer {
        for (result) |row| {
            allocator.free(row);
        }
        allocator.free(result);
    }

    try std.testing.expectEqual(@as(f64, 1.0), result[0][0]);
    try std.testing.expectEqual(@as(f64, 2.0), result[0][1]);
    try std.testing.expectEqual(@as(f64, 3.0), result[1][0]);
    try std.testing.expectEqual(@as(f64, 4.0), result[1][1]);
}

test "XLOPER12 round-trip for error type" {
    const allocator = std.testing.allocator;

    // Test various error codes
    const error_codes = [_]c_int{
        xl.xlerrNull,
        xl.xlerrDiv0,
        xl.xlerrValue,
        xl.xlerrRef,
        xl.xlerrName,
        xl.xlerrNum,
        xl.xlerrNA,
    };

    for (error_codes) |err_code| {
        var val1 = XLValue.err(allocator, err_code);
        defer val1.deinit();

        const xloper = val1.m_val;
        try std.testing.expectEqual(@as(@TypeOf(xloper.xltype), xl.xltypeErr), xloper.xltype);
        try std.testing.expectEqual(err_code, xloper.val.err);

        const val2 = XLValue.fromXLOPER12(allocator, xloper, false);
        try std.testing.expect(val2.is_err());
    }
}

test "XLOPER12 round-trip for missing and nil types" {
    const allocator = std.testing.allocator;

    // Test missing type
    {
        var val1 = XLValue.missing(allocator);
        defer val1.deinit();

        const xloper = val1.m_val;
        try std.testing.expectEqual(@as(@TypeOf(xloper.xltype), xl.xltypeMissing), xloper.xltype);

        const val2 = XLValue.fromXLOPER12(allocator, xloper, false);
        try std.testing.expect(val2.is_missing());
    }

    // Test nil type
    {
        var val1 = XLValue.init(allocator);
        val1.m_val.xltype = xl.xltypeNil;
        defer val1.deinit();

        const xloper = val1.m_val;
        try std.testing.expectEqual(@as(@TypeOf(xloper.xltype), xl.xltypeNil), xloper.xltype);

        const val2 = XLValue.fromXLOPER12(allocator, xloper, false);
        try std.testing.expect(val2.is_nil());
    }
}
