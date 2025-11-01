// Shared Excel imports to ensure type compatibility across modules
pub const win = @cImport({
    @cDefine("UNICODE", "1");
    @cDefine("_UNICODE", "1");
    @cInclude("windows.h");
});

pub const xl = @cImport({
    @cDefine("UNICODE", "1");
    @cDefine("_UNICODE", "1");
    @cInclude("windows.h");
    @cInclude("xlcall.h");
    @cInclude("framewrk.h");
});
