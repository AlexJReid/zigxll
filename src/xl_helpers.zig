const std = @import("std");
const builtin = @import("builtin");
const xl_imports = @import("xl_imports.zig");
const xl = xl_imports.xl;


// Platform-specific debug output
const outputDebug = if (builtin.os.tag == .windows)
    struct {
        extern "kernel32" fn OutputDebugStringA(lpOutputString: [*:0]const u8) callconv(.winapi) void;
    }.OutputDebugStringA
else
    struct {
        fn log(msg: [*:0]const u8) void {
            std.debug.print("{s}", .{msg});
        }
    }.log;

/// Simple debug log (no formatting)
pub inline fn debugLog(comptime msg: []const u8) void {
    outputDebug(msg ++ "\n");
}

/// Debug log with formatting
/// Usage: debugLogFmt("Error: {s}", .{@errorName(err)});
pub fn debugLogFmt(comptime fmt: []const u8, args: anytype) void {
    var buf: [512]u8 = undefined;
    const msg = std.fmt.bufPrintZ(&buf, fmt ++ "\n", args) catch {
        outputDebug("debugLogFmt: buffer too small\n");
        return;
    };
    outputDebug(msg.ptr);
}

pub inline fn xlFree(oper: *xl.XLOPER12) void {
    _ = xl.Excel12f(xl.xlFree, null, 1, oper);
}
