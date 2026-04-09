// rtd_call.zig — Helper for calling xlfRtd from within a UDF.
//
// Lets a UDF subscribe to an RTD topic and return the live value:
//
//   pub fn myPrice(symbol: []const u8) !*xl.XLOPER12 {
//       return rtd_call.subscribe("myprog.rtd", &.{symbol});
//   }
//
// Excel handles the subscription — the cell auto-updates when the
// RTD server pushes new values.

const std = @import("std");
const xl_imports = @import("xl_imports.zig");
const xl = xl_imports.xl;
const xlvalue = @import("xlvalue.zig");
const XLValue = xlvalue.XLValue;

const xl_helpers = @import("xl_helpers.zig");
const allocator = std.heap.c_allocator;

/// Call xlfRtd from a UDF to subscribe to an RTD topic.
///
/// `prog_id` is the RTD server's ProgID (e.g. "myprog.rtd").
/// `topics` is a slice of topic strings to pass to the RTD server.
///
/// Returns a heap-allocated XLOPER12 with xlbitDLLFree set.
/// Excel will call xlAutoFree12 to free it.
///
/// Use `subscribeDynamic` when `prog_id` is a runtime string (e.g. from Lua).
pub fn subscribe(comptime prog_id: []const u8, topics: []const []const u8) !*xl.XLOPER12 {
    // prog_id XLOPER12 string
    var prog_id_xl = try XLValue.fromUtf8String(allocator, prog_id);
    defer prog_id_xl.deinit();

    // Empty server string (local server)
    var server_xl = try XLValue.fromUtf8String(allocator, "");
    defer server_xl.deinit();

    // Build args array: prog_id, server, topic1, topic2, ...
    // xlfRtd supports up to 253 topic strings, we cap at 28 (plenty)
    const max_topics = 28;
    if (topics.len > max_topics) return error.TooManyTopics;

    var topic_xls: [max_topics]XLValue = undefined;
    var topic_count: usize = 0;
    defer for (topic_xls[0..topic_count]) |*t| t.deinit();

    for (topics) |topic| {
        topic_xls[topic_count] = try XLValue.fromUtf8String(allocator, topic);
        topic_count += 1;
    }

    // Excel12f is variadic — we need to pass the right number of args.
    // Build a pointer array and use Excel12v.
    const arg_count = 2 + topic_count;
    var args: [2 + max_topics][*c]xl.XLOPER12 = undefined;
    args[0] = &prog_id_xl.m_val;
    args[1] = &server_xl.m_val;
    for (0..topic_count) |i| {
        args[2 + i] = &topic_xls[i].m_val;
    }

    var result: xl.XLOPER12 = undefined;
    const ret = xl.Excel12v(xl.xlfRtd, &result, @intCast(arg_count), @ptrCast(&args));

    if (ret != xl.xlretSuccess) {
        return error.RtdCallFailed;
    }

    // xlfRtd returns an XLOPER12 with xlbitXLFree — Excel owns the data.
    // We must deep-copy any string data so xlAutoFree12 can safely free it.
    const heap_result = try allocator.create(xl.XLOPER12);
    const base_type = result.xltype & 0xFFF;

    if (base_type == xl.xltypeStr) {
        // Deep-copy the string buffer so we own it
        if (result.val.str) |str_ptr| {
            const len: usize = @intCast(str_ptr[0]);
            const total = len + 2; // length prefix + chars + null
            const new_buf = try allocator.alloc(u16, total);
            @memcpy(new_buf, str_ptr[0..total]);
            heap_result.* = .{
                .xltype = xl.xltypeStr | xl.xlbitDLLFree,
                .val = .{ .str = new_buf.ptr },
            };
        } else {
            heap_result.* = result;
            heap_result.xltype = xl.xltypeStr | xl.xlbitDLLFree;
        }
    } else {
        // Numeric/bool/error types have no pointers — shallow copy is fine
        heap_result.* = result;
        heap_result.xltype = base_type | xl.xlbitDLLFree;
    }

    // Free Excel's original allocation
    if ((result.xltype & xl.xlbitXLFree) != 0) {
        xl_helpers.xlFree(&result);
    }

    return heap_result;
}

/// Runtime-string variant of subscribe. Use this when prog_id is not a
/// comptime constant (e.g. when called from Lua via xllify.rtd_subscribe).
pub fn subscribeDynamic(prog_id: []const u8, topics: []const []const u8) !*xl.XLOPER12 {
    var prog_id_xl = try XLValue.fromUtf8String(allocator, prog_id);
    defer prog_id_xl.deinit();

    var server_xl = try XLValue.fromUtf8String(allocator, "");
    defer server_xl.deinit();

    const max_topics = 28;
    if (topics.len > max_topics) return error.TooManyTopics;

    var topic_xls: [max_topics]XLValue = undefined;
    var topic_count: usize = 0;
    defer for (topic_xls[0..topic_count]) |*t| t.deinit();

    for (topics) |topic| {
        topic_xls[topic_count] = try XLValue.fromUtf8String(allocator, topic);
        topic_count += 1;
    }

    const arg_count = 2 + topic_count;
    var args: [2 + max_topics][*c]xl.XLOPER12 = undefined;
    args[0] = &prog_id_xl.m_val;
    args[1] = &server_xl.m_val;
    for (0..topic_count) |i| {
        args[2 + i] = &topic_xls[i].m_val;
    }

    var result: xl.XLOPER12 = undefined;
    const ret = xl.Excel12v(xl.xlfRtd, &result, @intCast(arg_count), @ptrCast(&args));

    if (ret != xl.xlretSuccess) {
        return error.RtdCallFailed;
    }

    const heap_result = try allocator.create(xl.XLOPER12);
    const base_type = result.xltype & 0xFFF;

    if (base_type == xl.xltypeStr) {
        if (result.val.str) |str_ptr| {
            const len: usize = @intCast(str_ptr[0]);
            const total = len + 2;
            const new_buf = try allocator.alloc(u16, total);
            @memcpy(new_buf, str_ptr[0..total]);
            heap_result.* = .{
                .xltype = xl.xltypeStr | xl.xlbitDLLFree,
                .val = .{ .str = new_buf.ptr },
            };
        } else {
            heap_result.* = result;
            heap_result.xltype = xl.xltypeStr | xl.xlbitDLLFree;
        }
    } else {
        heap_result.* = result;
        heap_result.xltype = base_type | xl.xlbitDLLFree;
    }

    if ((result.xltype & xl.xlbitXLFree) != 0) {
        xl_helpers.xlFree(&result);
    }

    return heap_result;
}
