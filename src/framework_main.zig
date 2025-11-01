// Simple entry point for building the framework's test XLL
// This is only used when building the framework itself for development

const framework_entry = @import("framework_entry.zig");
const xl_imports = @import("xl_imports.zig");
const xl = xl_imports.xl;

pub const user_functions = @import("user_functions.zig");

// Export all Excel entry points
export fn xlAutoOpen() callconv(.c) c_int {
    return framework_entry.xlAutoOpen();
}

export fn xlAutoClose() callconv(.c) c_int {
    return framework_entry.xlAutoClose();
}

export fn xlAutoFree12(p: ?*xl.XLOPER12) callconv(.c) void {
    framework_entry.xlAutoFree12(p);
}

export fn xlAutoFree(p: ?*xl.XLOPER12) callconv(.c) void {
    framework_entry.xlAutoFree12(p);
}

export fn xlAutoRegister(pxName: ?*xl.XLOPER12) callconv(.c) ?*xl.XLOPER12 {
    _ = pxName;
    return null;
}

export fn xlAutoAdd() callconv(.c) c_int {
    return xl.xlretSuccess;
}

export fn xlAutoRemove() callconv(.c) c_int {
    return xl.xlretSuccess;
}
