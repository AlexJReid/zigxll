// Shared Excel imports to ensure type compatibility across modules
// win_compat.h provides minimal Windows types on all platforms,
// avoiding full windows.h which fails with Zig's @cImport on native MSVC.
pub const xl = @cImport({
    @cDefine("UNICODE", "1");
    @cDefine("_UNICODE", "1");
    @cInclude("xlcall.h");
    @cInclude("FRAMEWRK.H");
});
