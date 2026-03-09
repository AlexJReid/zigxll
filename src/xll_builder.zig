// XLL entry point - users never need to edit this file
// The framework uses this as the root source when building XLLs via buildXll()

const xll = @import("xll_framework");
const user_module = @import("user_module");
const rtd = xll.rtd;

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
    _ = pxName;
    return null;
}

export fn xlAutoRegister12(pxName: ?*xl.XLOPER12) callconv(.c) ?*xl.XLOPER12 {
    _ = pxName;
    return null;
}

export fn xlAutoAdd() callconv(.c) c_int {
    return xl.xlretSuccess;
}

export fn xlAutoRemove() callconv(.c) c_int {
    return xl.xlretSuccess;
}

// ============================================================================
// Combined DLL exports for all RTD servers (user + async framework)
// ============================================================================

// Collect user RTD server types
const user_rtd_servers = if (@hasDecl(user_module, "rtd_servers")) user_module.rtd_servers else .{};

// Check if any user function uses async
const has_async_functions = blk: {
    const all_functions = xll.framework_entry.getAllFunctions();
    for (all_functions) |FuncType| {
        if (@hasDecl(FuncType, "excel_is_async") and FuncType.excel_is_async) {
            break :blk true;
        }
    }
    break :blk false;
};

// Force the async RTD server's COM vtable to be instantiated if needed
comptime {
    if (has_async_functions) {
        _ = xll.async_handler.AsyncRtdServer;
    }
}

fn dllGetClassObject(rclsid: *const rtd.GUID, riid: *const rtd.GUID, ppv: *?*anyopaque) callconv(.winapi) rtd.HRESULT {
    // Try user RTD servers first
    inline for (user_rtd_servers) |server_module| {
        const Server = server_module.RtdServerType;
        const hr = Server.tryGetClassObject(rclsid, riid, ppv);
        if (hr == rtd.S_OK) return rtd.S_OK;
    }

    // Try the built-in async RTD server
    if (has_async_functions) {
        const hr = xll.async_handler.AsyncRtdServer.tryGetClassObject(rclsid, riid, ppv);
        if (hr == rtd.S_OK) return rtd.S_OK;
    }

    ppv.* = null;
    return @bitCast(@as(u32, 0x80040111)); // CLASS_E_CLASSNOTAVAILABLE
}

fn dllCanUnloadNow() callconv(.winapi) rtd.HRESULT {
    // Check user RTD servers
    inline for (user_rtd_servers) |server_module| {
        const Server = server_module.RtdServerType;
        if (Server.hasActiveObjects()) return rtd.S_FALSE;
    }

    // Check async RTD server
    if (has_async_functions) {
        if (xll.async_handler.AsyncRtdServer.hasActiveObjects()) return rtd.S_FALSE;
    }

    return rtd.S_OK;
}

comptime {
    // Only export DLL functions if there are any RTD servers
    const has_any_rtd = user_rtd_servers.len > 0 or has_async_functions;
    if (has_any_rtd) {
        @export(&dllGetClassObject, .{ .name = "DllGetClassObject" });
        @export(&dllCanUnloadNow, .{ .name = "DllCanUnloadNow" });
    }
}
