// XLL entry point - users never need to edit this file
// The framework uses this as the root source when building XLLs via buildXll()

const xll = @import("xll_framework");
const user_module = @import("user_module");

const xl = xll.xl;

pub const function_discovery = xll.function_discovery;
pub const user_functions = user_module;

// Export all XLL entry points - these call into the framework
export fn xlAutoOpen() callconv(.c) c_int {
    return xll.framework_entry.xlAutoOpen();
}

export fn xlAutoClose() callconv(.c) c_int {
    return xll.framework_entry.xlAutoClose();
}

export fn xlAutoFree12(p: ?*xl.XLOPER12) callconv(.c) void {
    xll.framework_entry.xlAutoFree12(p);
}

export fn xlAutoFree(p: ?*xl.XLOPER12) callconv(.c) void {
    xll.framework_entry.xlAutoFree12(p);
}

export fn xlAutoRegister(pxName: ?*xl.XLOPER12) callconv(.c) ?*xl.XLOPER12 {
    _ = pxName; // Required by Excel SDK but not supported
    return null;
}

export fn xlAutoAdd() callconv(.c) c_int {
    return xl.xlretSuccess;
}

export fn xlAutoRemove() callconv(.c) c_int {
    return xl.xlretSuccess;
}
