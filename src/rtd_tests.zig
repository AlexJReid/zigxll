const std = @import("std");
const builtin = @import("builtin");
const native_os = builtin.os.tag;

// rtd.zig uses std.os.windows types, so we can only import it on Windows.
const rtd = if (native_os == .windows) @import("rtd.zig") else struct {};
const rtd_registry = if (native_os == .windows) @import("rtd_registry.zig") else struct {};

// ============================================================================
// Cross-platform tests (comptime logic that doesn't need Windows types)
// ============================================================================

test "guid: parse basic GUID string" {
    const g = comptime parseGuid("A1B2C3D4-E5F6-7890-1234-567890ABCDEF");
    try std.testing.expectEqual(@as(u32, 0xA1B2C3D4), g.data1);
    try std.testing.expectEqual(@as(u16, 0xE5F6), g.data2);
    try std.testing.expectEqual(@as(u16, 0x7890), g.data3);
    try std.testing.expectEqualSlices(u8, &.{ 0x12, 0x34, 0x56, 0x78, 0x90, 0xAB, 0xCD, 0xEF }, &g.data4);
}

test "guid: parse with braces" {
    const g1 = comptime parseGuid("A1B2C3D4-E5F6-7890-1234-567890ABCDEF");
    const g2 = comptime parseGuid("{A1B2C3D4-E5F6-7890-1234-567890ABCDEF}");
    try std.testing.expectEqual(g1.data1, g2.data1);
    try std.testing.expectEqual(g1.data2, g2.data2);
    try std.testing.expectEqual(g1.data3, g2.data3);
    try std.testing.expectEqualSlices(u8, &g1.data4, &g2.data4);
}

test "guid: lowercase hex" {
    const g = comptime parseGuid("a1b2c3d4-e5f6-7890-abcd-ef0123456789");
    try std.testing.expectEqual(@as(u32, 0xA1B2C3D4), g.data1);
    try std.testing.expectEqualSlices(u8, &.{ 0xAB, 0xCD, 0xEF, 0x01, 0x23, 0x45, 0x67, 0x89 }, &g.data4);
}

test "guid: all zeros" {
    const g = comptime parseGuid("00000000-0000-0000-0000-000000000000");
    try std.testing.expectEqual(@as(u32, 0), g.data1);
    try std.testing.expectEqual(@as(u16, 0), g.data2);
    try std.testing.expectEqual(@as(u16, 0), g.data3);
    try std.testing.expectEqualSlices(u8, &.{ 0, 0, 0, 0, 0, 0, 0, 0 }, &g.data4);
}

test "guid: all F's" {
    const g = comptime parseGuid("FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF");
    try std.testing.expectEqual(@as(u32, 0xFFFFFFFF), g.data1);
    try std.testing.expectEqual(@as(u16, 0xFFFF), g.data2);
    try std.testing.expectEqual(@as(u16, 0xFFFF), g.data3);
    try std.testing.expectEqualSlices(u8, &.{ 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF }, &g.data4);
}

test "guid: IUnknown well-known GUID" {
    const g = comptime parseGuid("00000000-0000-0000-C000-000000000046");
    try std.testing.expectEqual(@as(u32, 0x00000000), g.data1);
    try std.testing.expectEqual(@as(u16, 0x0000), g.data2);
    try std.testing.expectEqual(@as(u16, 0x0000), g.data3);
    try std.testing.expectEqualSlices(u8, &.{ 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 }, &g.data4);
}

test "guidToString: round-trip" {
    const original = "A1B2C3D4-E5F6-7890-1234-567890ABCDEF";
    const g = comptime parseGuid(original);
    const result = comptime guidToStr(g);
    try std.testing.expectEqualStrings("{A1B2C3D4-E5F6-7890-1234-567890ABCDEF}", &result);
}

test "guidToString: round-trip all zeros" {
    const g = comptime parseGuid("00000000-0000-0000-0000-000000000000");
    const result = comptime guidToStr(g);
    try std.testing.expectEqualStrings("{00000000-0000-0000-0000-000000000000}", &result);
}

test "guidToString: round-trip IUnknown" {
    const g = comptime parseGuid("00000000-0000-0000-C000-000000000046");
    const result = comptime guidToStr(g);
    try std.testing.expectEqualStrings("{00000000-0000-0000-C000-000000000046}", &result);
}

test "RtdValue: construct int" {
    const v = RtdValue{ .int = 42 };
    try std.testing.expectEqual(@as(i32, 42), v.int);
}

test "RtdValue: construct double" {
    const v = RtdValue{ .double = 3.14 };
    try std.testing.expectEqual(@as(f64, 3.14), v.double);
}

test "RtdValue: construct boolean" {
    const v_true = RtdValue{ .boolean = true };
    const v_false = RtdValue{ .boolean = false };
    try std.testing.expect(v_true.boolean);
    try std.testing.expect(!v_false.boolean);
}

test "RtdValue: construct empty" {
    const v = RtdValue.empty;
    try std.testing.expectEqual(RtdValue.empty, v);
}

test "RtdValue: fromUtf8 comptime" {
    const v = comptime RtdValue.fromUtf8("hello");
    try std.testing.expectEqual(@as(usize, 5), v.string.len);
    try std.testing.expectEqual(@as(u16, 'h'), v.string[0]);
    try std.testing.expectEqual(@as(u16, 'o'), v.string[4]);
}

test "RtdValue: fromUtf8 unicode" {
    const v = comptime RtdValue.fromUtf8("€");
    // € is U+20AC, fits in a single UTF-16 code unit
    try std.testing.expectEqual(@as(usize, 1), v.string.len);
    try std.testing.expectEqual(@as(u16, 0x20AC), v.string[0]);
}

test "TopicEntry: default state" {
    const t = TopicEntry{};
    try std.testing.expect(!t.active);
    try std.testing.expectEqual(@as(i32, 0), t.topic_id);
    try std.testing.expect(!t.dirty);
}

test "topic tracking: add and remove" {
    var topics: [8]TopicEntry = .{TopicEntry{}} ** 8;
    var count: usize = 0;

    // Add topic 10
    for (&topics) |*t| {
        if (!t.active) {
            t.active = true;
            t.topic_id = 10;
            t.dirty = true;
            count += 1;
            break;
        }
    }
    try std.testing.expectEqual(@as(usize, 1), count);

    // Add topic 20
    for (&topics) |*t| {
        if (!t.active) {
            t.active = true;
            t.topic_id = 20;
            t.dirty = true;
            count += 1;
            break;
        }
    }
    try std.testing.expectEqual(@as(usize, 2), count);

    // Remove topic 10
    for (&topics) |*t| {
        if (t.active and t.topic_id == 10) {
            t.active = false;
            count -= 1;
            break;
        }
    }
    try std.testing.expectEqual(@as(usize, 1), count);

    // Verify topic 20 still active
    var found_20 = false;
    for (&topics) |*t| {
        if (t.active and t.topic_id == 20) {
            found_20 = true;
        }
    }
    try std.testing.expect(found_20);
}

test "topic tracking: dirty counting" {
    var topics: [8]TopicEntry = .{TopicEntry{}} ** 8;

    // Add 3 topics
    topics[0] = .{ .active = true, .topic_id = 1, .dirty = true };
    topics[1] = .{ .active = true, .topic_id = 2, .dirty = false };
    topics[2] = .{ .active = true, .topic_id = 3, .dirty = true };

    var dirty_count: usize = 0;
    for (&topics) |*t| {
        if (t.active and t.dirty) dirty_count += 1;
    }
    try std.testing.expectEqual(@as(usize, 2), dirty_count);
}

test "topic tracking: markAllDirty pattern" {
    var topics: [8]TopicEntry = .{TopicEntry{}} ** 8;

    topics[0] = .{ .active = true, .topic_id = 1, .dirty = false };
    topics[1] = .{ .active = false, .topic_id = 0, .dirty = false };
    topics[2] = .{ .active = true, .topic_id = 3, .dirty = false };

    // markAllDirty
    for (&topics) |*t| {
        if (t.active) t.dirty = true;
    }

    try std.testing.expect(topics[0].dirty);
    try std.testing.expect(!topics[1].dirty); // inactive, not dirtied
    try std.testing.expect(topics[2].dirty);
}

test "topic tracking: slot reuse after disconnect" {
    var topics: [4]TopicEntry = .{TopicEntry{}} ** 4;

    // Fill first slot
    topics[0] = .{ .active = true, .topic_id = 100, .dirty = true };

    // Disconnect it
    topics[0].active = false;

    // New topic should reuse slot 0
    for (&topics) |*t| {
        if (!t.active) {
            t.active = true;
            t.topic_id = 200;
            t.dirty = true;
            break;
        }
    }
    try std.testing.expectEqual(@as(i32, 200), topics[0].topic_id);
    try std.testing.expect(topics[0].active);
}

// ============================================================================
// Windows-only tests (need COM APIs at link time)
// ============================================================================

test "rtdValueToVariant: int" {
    if (native_os != .windows) return error.SkipZigTest;
    var v: rtd.VARIANT = undefined;
    rtd.rtdValueToVariant(.{ .int = 42 }, &v);
    try std.testing.expectEqual(@as(u16, 3), v.vt); // VT_I4
    try std.testing.expectEqual(@as(c_long, 42), v.data.lval);
}

test "rtdValueToVariant: double" {
    if (native_os != .windows) return error.SkipZigTest;
    var v: rtd.VARIANT = undefined;
    rtd.rtdValueToVariant(.{ .double = 3.14 }, &v);
    try std.testing.expectEqual(@as(u16, 5), v.vt); // VT_R8
    try std.testing.expectEqual(@as(f64, 3.14), v.data.dval);
}

test "rtdValueToVariant: boolean true" {
    if (native_os != .windows) return error.SkipZigTest;
    var v: rtd.VARIANT = undefined;
    rtd.rtdValueToVariant(.{ .boolean = true }, &v);
    try std.testing.expectEqual(@as(u16, 11), v.vt); // VT_BOOL
    try std.testing.expectEqual(@as(i16, -1), v.data.boolval); // VARIANT_TRUE
}

test "rtdValueToVariant: empty" {
    if (native_os != .windows) return error.SkipZigTest;
    var v: rtd.VARIANT = undefined;
    rtd.rtdValueToVariant(.empty, &v);
    try std.testing.expectEqual(@as(u16, 0), v.vt); // VT_EMPTY
}

// ============================================================================
// Portable re-implementations of comptime GUID logic for testing
// (avoids importing rtd.zig which needs std.os.windows)
// ============================================================================

const GuidData = struct {
    data1: u32,
    data2: u16,
    data3: u16,
    data4: [8]u8,
};

fn hexDigit(c: u8) u8 {
    return switch (c) {
        '0'...'9' => c - '0',
        'a'...'f' => c - 'a' + 10,
        'A'...'F' => c - 'A' + 10,
        else => @compileError("invalid hex digit in GUID"),
    };
}

fn hexU8(comptime s: *const [2]u8) u8 {
    return @as(u8, hexDigit(s[0])) << 4 | hexDigit(s[1]);
}

fn hexU16(comptime s: *const [4]u8) u16 {
    return @as(u16, hexU8(s[0..2])) << 8 | hexU8(s[2..4]);
}

fn hexU32(comptime s: *const [8]u8) u32 {
    return @as(u32, hexU16(s[0..4])) << 16 | hexU16(s[4..8]);
}

fn parseGuid(comptime s: []const u8) GuidData {
    comptime {
        const raw = if (s.len > 0 and s[0] == '{' and s[s.len - 1] == '}')
            s[1 .. s.len - 1]
        else
            s;

        if (raw.len != 36) @compileError("GUID string must be 36 chars");
        if (raw[8] != '-' or raw[13] != '-' or raw[18] != '-' or raw[23] != '-')
            @compileError("GUID string has wrong dash positions");

        return .{
            .data1 = hexU32(raw[0..8]),
            .data2 = hexU16(raw[9..13]),
            .data3 = hexU16(raw[14..18]),
            .data4 = .{
                hexU8(raw[19..21]), hexU8(raw[21..23]),
                hexU8(raw[24..26]), hexU8(raw[26..28]),
                hexU8(raw[28..30]), hexU8(raw[30..32]),
                hexU8(raw[32..34]), hexU8(raw[34..36]),
            },
        };
    }
}

fn hexCharOut(val: u4) u8 {
    const v: u8 = val;
    return if (v < 10) '0' + v else 'A' - 10 + v;
}

fn hexByteOut(b: u8) [2]u8 {
    return .{ hexCharOut(@truncate(b >> 4)), hexCharOut(@truncate(b & 0xF)) };
}

fn guidToStr(comptime g: GuidData) [38]u8 {
    comptime {
        var buf: [38]u8 = undefined;
        buf[0] = '{';

        const d1 = std.mem.toBytes(std.mem.nativeToBig(u32, g.data1));
        var pos: usize = 1;
        for (d1) |b| {
            const h = hexByteOut(b);
            buf[pos] = h[0];
            buf[pos + 1] = h[1];
            pos += 2;
        }
        buf[pos] = '-';
        pos += 1;

        const d2 = std.mem.toBytes(std.mem.nativeToBig(u16, g.data2));
        for (d2) |b| {
            const h = hexByteOut(b);
            buf[pos] = h[0];
            buf[pos + 1] = h[1];
            pos += 2;
        }
        buf[pos] = '-';
        pos += 1;

        const d3 = std.mem.toBytes(std.mem.nativeToBig(u16, g.data3));
        for (d3) |b| {
            const h = hexByteOut(b);
            buf[pos] = h[0];
            buf[pos + 1] = h[1];
            pos += 2;
        }
        buf[pos] = '-';
        pos += 1;

        for (g.data4[0..2]) |b| {
            const h = hexByteOut(b);
            buf[pos] = h[0];
            buf[pos + 1] = h[1];
            pos += 2;
        }
        buf[pos] = '-';
        pos += 1;

        for (g.data4[2..8]) |b| {
            const h = hexByteOut(b);
            buf[pos] = h[0];
            buf[pos + 1] = h[1];
            pos += 2;
        }

        buf[37] = '}';
        return buf;
    }
}

const TopicEntry = struct {
    active: bool = false,
    topic_id: i32 = 0,
    dirty: bool = false,
};

const RtdValue = union(enum) {
    int: i32,
    double: f64,
    string: []const u16,
    boolean: bool,
    empty,

    pub fn fromUtf8(comptime s: []const u8) RtdValue {
        return .{ .string = std.unicode.utf8ToUtf16LeStringLiteral(s) };
    }
};
