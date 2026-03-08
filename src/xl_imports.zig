// Shared Excel imports to ensure type compatibility across modules
const builtin = @import("builtin");

// Import Excel types from xlcall.h (works cross-platform via win_compat.h)
const xl_c = @cImport({
    @cDefine("UNICODE", "1");
    @cDefine("_UNICODE", "1");
    if (builtin.os.tag == .windows) {
        @cInclude("windows.h");
    }
    @cInclude("xlcall.h");
});

// Re-export all C types directly
pub const xl = xl_c;

// Excel12v entry point - replaces framewrk32.lib
// Excel calls SetExcel12EntryPt during xlAutoOpen to give us the function pointer.
const EXCEL12PROC = *const fn (xlfn: c_int, coper: c_int, rgpxloper12: [*]?*xl_c.XLOPER12, xloper12Res: ?*xl_c.XLOPER12) callconv(.c) c_int;
var pexcel12: ?EXCEL12PROC = null;

pub export fn SetExcel12EntryPt(p: EXCEL12PROC) callconv(.c) void {
    pexcel12 = p;
}

pub fn Excel12v(xlfn: c_int, operRes: ?*xl_c.XLOPER12, count: c_int, opers: [*]?*xl_c.XLOPER12) c_int {
    if (builtin.os.tag != .windows) return xl_c.xlretSuccess;
    const proc = pexcel12 orelse return xl_c.xlretFailed;
    return proc(xlfn, count, opers, operRes);
}

/// Wrapper around Excel12v that accepts XLOPER12 pointer args as a tuple.
pub fn Excel12f(xlfn: c_int, operRes: ?*xl_c.XLOPER12, count: c_int, args: anytype) c_int {
    var opers: [16]?*xl_c.XLOPER12 = .{null} ** 16;
    inline for (0..args.len) |i| {
        opers[i] = args[i];
    }
    return Excel12v(xlfn, operRes, count, &opers);
}
