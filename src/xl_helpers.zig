const std = @import("std");
const xl_imports = @import("xl_imports.zig");
const xl = xl_imports.xl;
const win = xl_imports.win;

/// Simple debug log (no formatting)
pub inline fn debugLog(comptime msg: []const u8) void {
    win.OutputDebugStringA(msg ++ "\n");
}

/// Debug log with formatting
/// Usage: debugLogFmt("Error: {s}", .{@errorName(err)});
pub fn debugLogFmt(comptime fmt: []const u8, args: anytype) void {
    var buf: [512]u8 = undefined;
    const msg = std.fmt.bufPrintZ(&buf, fmt ++ "\n", args) catch {
        win.OutputDebugStringA("debugLogFmt: buffer too small\n");
        return;
    };
    win.OutputDebugStringA(msg.ptr);
}

pub inline fn xlFree(oper: *xl.XLOPER12) void {
    _ = xl.Excel12f(xl.xlFree, null, 1, oper);
}
