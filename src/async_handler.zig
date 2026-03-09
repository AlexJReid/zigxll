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

const TopicKeyMap = std.AutoHashMap(i32, []const u8);

pub const AsyncHandler = struct {
    // Map topic_id → cache key.
    topic_keys: TopicKeyMap = TopicKeyMap.init(allocator),

    pub fn onStart(_: *AsyncHandler, ctx: *rtd.RtdContext) void {
        // Stash pointers so async workers can mark dirty + call UpdateNotify
        async_infra.global_update_event = ctx.update_event;
        async_infra.global_rtd_context = ctx;
    }

    pub fn onConnect(self: *AsyncHandler, _: *rtd.RtdContext, topic_id: i32, _: usize) void {
        // Pick up the pending key set by the UDF
        const key = pending_connect_key orelse return;
        pending_connect_key = null;

        const owned = allocator.dupe(u8, key) catch return;
        self.topic_keys.put(topic_id, owned) catch {
            allocator.free(owned);
        };
    }

    pub fn onDisconnect(self: *AsyncHandler, _: *rtd.RtdContext, topic_id: i32, _: usize) void {
        if (self.topic_keys.fetchRemove(topic_id)) |entry| {
            allocator.free(@constCast(entry.value));
        }
    }

    pub fn onRefreshValue(self: *AsyncHandler, _: *rtd.RtdContext, topic_id: i32) rtd.RtdValue {
        const cache = async_cache.getGlobalCache();

        if (self.topic_keys.get(topic_id)) |key| {
            if (cache.get(key)) |entry| {
                return xlopToRtdValue(entry.xloper);
            }
        }
        // No cached value yet — return #N/A as loading indicator
        return rtd.RtdValue.na;
    }

    pub fn onTerminate(self: *AsyncHandler, _: *rtd.RtdContext) void {
        var it = self.topic_keys.iterator();
        while (it.next()) |entry| {
            allocator.free(@constCast(entry.value_ptr.*));
        }
        self.topic_keys.deinit();
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
