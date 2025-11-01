const std = @import("std");

// Re-export all framework modules
pub const xl_imports = @import("xl_imports.zig");
pub const xl = xl_imports.xl;
pub const win = xl_imports.win;

pub const xlvalue = @import("xlvalue.zig");
pub const XLValue = xlvalue.XLValue;

pub const xl_helpers = @import("xl_helpers.zig");
pub const function_discovery = @import("function_discovery.zig");

const excel_allocator = std.heap.c_allocator;
var initialized = false;

const framework_version = "0.1.0";

// Discover all functions from user's modules at comptime
// Note: This expects @import("root") to have a `user_functions` declaration
pub fn getAllFunctions() []const type {
    const root = @import("root");
    if (!@hasDecl(root, "user_functions")) {
        @compileError("Root module must have a 'user_functions' declaration with 'function_modules' tuple");
    }
    const user_functions = root.user_functions;

    comptime var funcs: []const type = &.{};
    inline for (user_functions.function_modules) |module| {
        const module_funcs = function_discovery.getAllFunctions(module);
        funcs = funcs ++ module_funcs;
    }
    return funcs;
}

/// Initialize XLL - auto-discover and register
pub fn xlAutoOpen() callconv(.c) c_int {
    const build_options = @import("build_options");
    xl_helpers.debugLogFmt("zigxll v{s}: {s} loaded", .{ framework_version, build_options.xll_name });

    // Create arena for all registration strings - destroyed at end of this function
    var registration_arena = std.heap.ArenaAllocator.init(excel_allocator);
    defer registration_arena.deinit();
    const reg_allocator = registration_arena.allocator();

    // Get DLL path for function registration
    var xDLL: xl.XLOPER12 = undefined;
    const getName_result = xl.Excel12f(xl.xlGetName, &xDLL, 0);
    defer xl_helpers.xlFree(&xDLL);
    if (getName_result != xl.xlretSuccess) {
        xl_helpers.debugLog("Failed to get XLL path");
        return xl.xlretFailed;
    }

    // Register all discovered functions
    const all_functions = comptime getAllFunctions();
    inline for (all_functions) |FuncType| {
        registerFunction(FuncType, &xDLL, reg_allocator) catch |err| {
            xl_helpers.debugLogFmt("Failed to register function '{s}': {s}", .{ FuncType.excel_name, @errorName(err) });
            return xl.xlretFailed;
        };
    }
    xl_helpers.debugLogFmt("Successfully registered {d} functions", .{all_functions.len});

    initialized = true;
    return xl.xlretSuccess;
}

fn registerFunction(comptime FuncType: type, xll_path: *xl.XLOPER12, allocator: std.mem.Allocator) !void {
    // Convert metadata to XLOPER12 for registration (uses passed-in arena allocator)
    var proc_name_xl = try XLValue.fromUtf8String(allocator, FuncType.excel_export_name);
    var type_string_xl = try XLValue.fromUtf8String(allocator, FuncType.excel_type_string);
    var func_name_xl = try XLValue.fromUtf8String(allocator, FuncType.excel_name);

    // TODO: Build argument names from FuncType.excel_params
    var arg_names_xl = try XLValue.fromUtf8String(allocator, "");
    var func_type_xl = try XLValue.fromUtf8String(allocator, "1"); // 1 = normal function
    var category_xl = try XLValue.fromUtf8String(allocator, FuncType.excel_category);
    var description_xl = try XLValue.fromUtf8String(allocator, FuncType.excel_description);
    var empty_xl = try XLValue.fromUtf8String(allocator, "");

    // Call xlfRegister
    var result: xl.XLOPER12 = undefined;
    const ret = xl.Excel12f(
        xl.xlfRegister,
        &result,
        11,
        xll_path,
        &proc_name_xl.m_val,
        &type_string_xl.m_val,
        &func_name_xl.m_val,
        &arg_names_xl.m_val,
        &func_type_xl.m_val,
        &category_xl.m_val,
        &empty_xl.m_val, // Shortcut
        &empty_xl.m_val, // Help topic
        &description_xl.m_val, // Description
        &empty_xl.m_val, // arg
    );
    defer xl_helpers.xlFree(&result);

    if (ret != xl.xlretSuccess) {
        return error.RegistrationFailed;
    }

    xl_helpers.debugLogFmt("Registered Excel function: {s} ({s})", .{ FuncType.excel_name, FuncType.excel_type_string });
}

pub fn xlAutoClose() callconv(.c) c_int {
    xl_helpers.debugLog("xlAutoClose called");
    initialized = false;
    return xl.xlretSuccess;
}

pub fn xlAutoFree12(pxFree: ?*xl.XLOPER12) callconv(.c) void {
    if (pxFree) |oper| {
        // Only free if xlbitDLLFree is set (means we allocated it)
        if ((oper.xltype & xl.xlbitDLLFree) != 0) {
            // Free string data if present
            if ((oper.xltype & xl.xltypeStr) != 0) {
                if (oper.val.str) |str_ptr| {
                    const len = @as(usize, @intCast(str_ptr[0]));
                    // Free: length prefix (1) + string chars (len) + null terminator (1)
                    excel_allocator.free(str_ptr[0 .. len + 2]);
                }
            }
            // Free the XLOPER12 structure itself
            excel_allocator.destroy(oper);
        }
    }
}
