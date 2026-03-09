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

const allocator = std.heap.c_allocator;

/// Call xlfRtd from a UDF to subscribe to an RTD topic.
///
/// `prog_id` is the RTD server's ProgID (e.g. "myprog.rtd").
/// `topics` is a slice of topic strings to pass to the RTD server.
///
/// Returns a heap-allocated XLOPER12 with xlbitDLLFree set.
/// Excel will call xlAutoFree12 to free it.
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

    // Copy result to heap with DLLFree flag so xlAutoFree12 cleans it up
    const heap_result = try allocator.create(xl.XLOPER12);
    heap_result.* = result;
    heap_result.xltype |= xl.xlbitDLLFree;
    return heap_result;
}
