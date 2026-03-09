// Built-in RTD handler for async Excel functions.
//
// All async functions share a single RTD server ("zigxll.async").
// Topic keys are "FuncName|arg1|arg2|..." — the handler looks up
// completed values in the AsyncCache and returns them via RefreshData.
//
// The topic_id → topic_key mapping is set via a pending key mechanism:
// the UDF sets `pending_connect_key` right before calling xlfRtd,
// and ConnectData (called synchronously by Excel) picks it up.

const std = @import("std");
const rtd = @import("rtd.zig");
const async_cache = @import("async_cache.zig");
const async_infra = @import("async_infra.zig");
const xl_imports = @import("xl_imports.zig");
const xl = xl_imports.xl;

const allocator = std.heap.c_allocator;

// Set by the UDF right before calling xlfRtd.
// ConnectData runs synchronously on the same thread, so this is safe.
pub var pending_connect_key: ?[]const u8 = null;

pub const AsyncHandler = struct {
    // Map topic_id → cache key.  Parallel array to RtdContext.topics.
    topic_keys: [rtd.MAX_TOPICS]?[]const u8 = [_]?[]const u8{null} ** rtd.MAX_TOPICS,

    pub fn onStart(_: *AsyncHandler, ctx: *rtd.RtdContext) void {
        // Stash pointers so async workers can mark dirty + call UpdateNotify
        async_infra.global_update_event = ctx.update_event;
        async_infra.global_rtd_context = ctx;
    }

    pub fn onConnect(self: *AsyncHandler, ctx: *rtd.RtdContext, topic_id: i32, _: usize) void {
        // Pick up the pending key set by the UDF
        const key = pending_connect_key orelse return;
        pending_connect_key = null;

        // Find the slot that was just assigned this topic_id
        for (ctx.topics, 0..) |t, i| {
            if (t.active and t.topic_id == topic_id) {
                self.topic_keys[i] = allocator.dupe(u8, key) catch null;
                break;
            }
        }
    }

    pub fn onDisconnect(self: *AsyncHandler, ctx: *rtd.RtdContext, topic_id: i32, _: usize) void {
        for (ctx.topics, 0..) |t, i| {
            // The topic is still marked active at this point (cleared after this call)
            if (t.active and t.topic_id == topic_id) {
                if (self.topic_keys[i]) |k| allocator.free(k);
                self.topic_keys[i] = null;
                break;
            }
        }
    }

    pub fn onRefreshValue(self: *AsyncHandler, ctx: *rtd.RtdContext, topic_id: i32) rtd.RtdValue {
        const cache = async_cache.getGlobalCache();

        // Find the topic key for this topic_id
        for (ctx.topics, 0..) |t, i| {
            if (t.active and t.topic_id == topic_id) {
                if (self.topic_keys[i]) |key| {
                    if (cache.get(key)) |entry| {
                        // Return current value whether completed or intermediate
                        return xlopToRtdValue(entry.xloper);
                    }
                }
                break;
            }
        }
        // No cached value yet — return #N/A as loading indicator
        return rtd.RtdValue.na;
    }

    pub fn onTerminate(self: *AsyncHandler, _: *rtd.RtdContext) void {
        for (&self.topic_keys) |*k| {
            if (k.*) |key| {
                allocator.free(key);
                k.* = null;
            }
        }
    }
};

fn xlopToRtdValue(xlop: *xl.XLOPER12) rtd.RtdValue {
    const base_type = xlop.xltype & ~@as(u32, xl.xlbitDLLFree | xl.xlbitXLFree);
    return switch (base_type) {
        xl.xltypeNum => .{ .double = xlop.val.num },
        xl.xltypeBool => .{ .boolean = xlop.val.xbool != 0 },
        xl.xltypeStr => blk: {
            if (xlop.val.str) |str_ptr| {
                const len: usize = @intCast(str_ptr[0]);
                break :blk .{ .string = str_ptr[1 .. 1 + len] };
            }
            break :blk .empty;
        },
        xl.xltypeInt => .{ .int = @intCast(xlop.val.w) },
        xl.xltypeErr => rtd.RtdValue.na,
        else => .empty,
    };
}

// RTD config for the built-in async server
pub const async_rtd_clsid = rtd.guid("B5E45CC4-0530-4CEA-B6A2-F17C2E53A1D9");

pub const rtd_config = rtd.RtdConfig{
    .clsid = async_rtd_clsid,
    .prog_id = "zigxll.async",
};

pub const AsyncRtdServer = rtd.RtdServer(AsyncHandler, rtd_config);
