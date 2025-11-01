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

    /// Extract UTF-8 string from Excel's wide string XLOPER12
    /// Excel stores strings as: first u16 = length, rest = wide chars
    /// Returns allocated UTF-8 string that caller must free
    pub fn as_utf8str(self: *const XLValue) ![]u8 {
        if (!self.is_str()) return error.NotAString;

        const len = self.m_val.val.str[0];
        const wide_slice = self.m_val.val.str[1 .. len + 1];

        // Allocate buffer (worst case: each UTF-16 char becomes 3 UTF-8 bytes)
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
