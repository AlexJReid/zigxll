// Shared Excel imports to ensure type compatibility across modules
const builtin = @import("builtin");

// Import Excel types from xlcall.h + framewrk.h (works cross-platform via win_compat.h)
pub const xl = @cImport({
    @cDefine("UNICODE", "1");
    @cDefine("_UNICODE", "1");
    if (builtin.os.tag == .windows) {
        @cInclude("windows.h");
    }
    @cInclude("xlcall.h");
    @cInclude("FRAMEWRK.H");
});
