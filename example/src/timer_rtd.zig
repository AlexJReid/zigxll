// Timer RTD server — ticks a counter every ~2 seconds.
// Built on top of rtd.zig's generic RtdServer.
//
// Usage in Excel: =RTD("zigxll.rtd", , "anything")
// Registration:   automatic via xlAutoOpen (writes to HKCU, no admin needed)

const std = @import("std");
const xll = @import("xll");
const rtd = xll.rtd;

const TimerHandler = struct {
    counter: i32 = 0,
    timer_thread: ?std.Thread = null,
    running: std.atomic.Value(bool) = std.atomic.Value(bool).init(false),

    pub fn onStart(self: *TimerHandler, ctx: *rtd.RtdContext) void {
        self.running.store(true, .release);
        self.timer_thread = std.Thread.spawn(.{}, timerProc, .{ self, ctx }) catch return;
    }

    pub fn onConnect(_: *TimerHandler, _: *rtd.RtdContext, _: i32, _: usize) void {}

    pub fn onDisconnect(_: *TimerHandler, _: *rtd.RtdContext, _: i32, _: usize) void {}

    pub fn onRefreshValue(self: *TimerHandler, _: *rtd.RtdContext, _: i32) rtd.RtdValue {
        return .{ .int = self.counter };
    }

    pub fn onTerminate(self: *TimerHandler, _: *rtd.RtdContext) void {
        self.running.store(false, .release);
        if (self.timer_thread) |t| {
            t.join();
            self.timer_thread = null;
        }
    }

    fn timerProc(self: *TimerHandler, ctx: *rtd.RtdContext) void {
        rtd.debugLog("Timer thread started", .{});
        while (self.running.load(.acquire)) {
            std.Thread.sleep(2 * std.time.ns_per_s);
            if (!self.running.load(.acquire)) break;

            self.counter += 1;
            ctx.markAllDirty();
            ctx.notifyExcel();
        }
        rtd.debugLog("Timer thread stopped", .{});
    }
};

pub const rtd_config: rtd.RtdConfig = .{
    .clsid = rtd.guid("A1B2C3D4-E5F6-7890-1234-567890ABCDEF"),
    .prog_id = "zigxll.rtd",
};

const TimerRtd = rtd.RtdServer(TimerHandler, rtd_config);

comptime {
    TimerRtd.exportDllFunctions();
}
