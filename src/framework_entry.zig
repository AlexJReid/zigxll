const std = @import("std");

// Re-export all framework modules
pub const xl_imports = @import("xl_imports.zig");
pub const xl = xl_imports.xl;

pub const xlvalue = @import("xlvalue.zig");
pub const XLValue = xlvalue.XLValue;

pub const xl_helpers = @import("xl_helpers.zig");
pub const function_discovery = @import("function_discovery.zig");
pub const rtd_registry = @import("rtd_registry.zig");
pub const async_handler = @import("async_handler.zig");

const excel_allocator = std.heap.c_allocator;
var initialized = false;

const framework_version = @import("build_options").framework_version;

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

// Discover all macros from user's modules at comptime
pub fn getAllMacros() []const type {
    const root = @import("root");
    if (!@hasDecl(root, "user_functions")) {
        @compileError("Root module must have a 'user_functions' declaration with 'function_modules' tuple");
    }
    const user_functions = root.user_functions;

    comptime var macros: []const type = &.{};
    inline for (user_functions.function_modules) |module| {
        const module_macros = function_discovery.getAllMacros(module);
        macros = macros ++ module_macros;
    }
    return macros;
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

    // Register all discovered macros
    const all_macros = comptime getAllMacros();
    inline for (all_macros) |MacroType| {
        registerMacro(MacroType, &xDLL, reg_allocator) catch |err| {
            xl_helpers.debugLogFmt("Failed to register macro '{s}': {s}", .{ MacroType.excel_name, @errorName(err) });
            return xl.xlretFailed;
        };
    }
    if (all_macros.len > 0) {
        xl_helpers.debugLogFmt("Successfully registered {d} macros", .{all_macros.len});
    }

    // Initialize Lua and load scripts if user module declares them
    const user_mod = @import("root").user_functions;
    if (comptime @hasDecl(user_mod, "lua_scripts")) {
        const lua = @import("lua.zig");
        lua.init() catch {
            xl_helpers.debugLog("Failed to initialize Lua");
            return xl.xlretFailed;
        };
        xl_helpers.debugLogFmt("Lua: initialized {d} state(s)", .{lua.getPoolSize()});
        inline for (user_mod.lua_scripts) |script| {
            lua.loadScript(script.source, script.name) catch {
                xl_helpers.debugLogFmt("Failed to load Lua script: {s}", .{script.name});
                return xl.xlretFailed;
            };
        }
        xl_helpers.debugLogFmt("Loaded {d} Lua scripts", .{user_mod.lua_scripts.len});
    }

    // Auto-register RTD servers if user module declares them
    if (comptime @hasDecl(user_mod, "rtd_servers")) {
        if (rtd_registry.getXllPathSlice(&xDLL)) |xll_path| {
            inline for (user_mod.rtd_servers) |server_module| {
                const cfg = server_module.rtd_config;
                rtd_registry.registerRtdServer(cfg.clsid, cfg.prog_id, xll_path);
            }
        }
    }

    // Auto-register the built-in async RTD server if any function uses async
    const has_async = comptime blk: {
        for (all_functions) |FuncType| {
            if (@hasDecl(FuncType, "excel_is_async") and FuncType.excel_is_async) {
                break :blk true;
            }
        }
        break :blk false;
    };
    if (has_async) {
        if (rtd_registry.getXllPathSlice(&xDLL)) |xll_path| {
            const acfg = async_handler.rtd_config;
            rtd_registry.registerRtdServer(acfg.clsid, acfg.prog_id, xll_path);
            xl_helpers.debugLog("Registered async RTD server (zigxll.async)");
        }
    }

    initialized = true;
    return xl.xlretSuccess;
}

fn registerFunction(comptime FuncType: type, xll_path: *xl.XLOPER12, allocator: std.mem.Allocator) !void {
    // Convert metadata to XLOPER12 for registration (uses passed-in arena allocator)
    var proc_name_xl = try XLValue.fromUtf8String(allocator, FuncType.excel_export_name);
    var type_string_xl = try XLValue.fromUtf8String(allocator, FuncType.excel_type_string);
    var func_name_xl = try XLValue.fromUtf8String(allocator, FuncType.excel_name);

    // Build argument names from FuncType.excel_params
    const arg_names_str = comptime blk: {
        var result: []const u8 = "";
        for (FuncType.excel_params, 0..) |param, i| {
            if (i > 0) result = result ++ ",";
            if (param.name) |name| {
                result = result ++ name;
            } else {
                // Default to arg1, arg2, arg3, etc.
                result = result ++ "arg" ++ std.fmt.comptimePrint("{d}", .{i + 1});
            }
        }
        break :blk result;
    };
    var arg_names_xl = try XLValue.fromUtf8String(allocator, arg_names_str);
    var func_type_xl = try XLValue.fromUtf8String(allocator, "1"); // 1 = normal function
    var category_xl = try XLValue.fromUtf8String(allocator, FuncType.excel_category);
    var description_xl = try XLValue.fromUtf8String(allocator, FuncType.excel_description);
    var empty_xl = try XLValue.fromUtf8String(allocator, "");

    // Build argument help descriptions
    const param_count = FuncType.excel_params.len;
    var arg_help: [8]XLValue = undefined; // Max 8 params
    inline for (0..param_count) |i| {
        const desc = FuncType.excel_params[i].description orelse "";
        arg_help[i] = try XLValue.fromUtf8String(allocator, desc);
    }

    // Build args array for Excel12v (only the ones we need)
    const arg_count = 10 + param_count + 1; // +1 for dummy trailing arg
    var args: [19][*c]xl.XLOPER12 = undefined;
    args[0] = xll_path;
    args[1] = &proc_name_xl.m_val;
    args[2] = &type_string_xl.m_val;
    args[3] = &func_name_xl.m_val;
    args[4] = &arg_names_xl.m_val;
    args[5] = &func_type_xl.m_val;
    args[6] = &category_xl.m_val;
    args[7] = &empty_xl.m_val; // Shortcut
    args[8] = &empty_xl.m_val; // Help topic
    args[9] = &description_xl.m_val; // Function description

    // Add argument descriptions
    inline for (0..param_count) |i| {
        args[10 + i] = &arg_help[i].m_val;
    }

    // Add dummy trailing empty string to prevent truncation
    args[10 + param_count] = &empty_xl.m_val;

    // Call xlfRegister using Excel12v
    var result: xl.XLOPER12 = undefined;
    const ret = xl.Excel12v(xl.xlfRegister, &result, @intCast(arg_count), @ptrCast(&args));
    defer xl_helpers.xlFree(&result);

    if (ret != xl.xlretSuccess) {
        return error.RegistrationFailed;
    }

    xl_helpers.debugLogFmt("Registered Excel function: {s} ({s})", .{ FuncType.excel_name, FuncType.excel_type_string });
}

fn registerMacro(comptime MacroType: type, xll_path: *xl.XLOPER12, allocator: std.mem.Allocator) !void {
    var proc_name_xl = try XLValue.fromUtf8String(allocator, MacroType.excel_export_name);
    var type_string_xl = try XLValue.fromUtf8String(allocator, MacroType.excel_type_string);
    var func_name_xl = try XLValue.fromUtf8String(allocator, MacroType.excel_name);
    var arg_names_xl = try XLValue.fromUtf8String(allocator, "");
    var func_type_xl = try XLValue.fromUtf8String(allocator, "2"); // 2 = macro/command
    var category_xl = try XLValue.fromUtf8String(allocator, MacroType.excel_category);
    var description_xl = try XLValue.fromUtf8String(allocator, MacroType.excel_description);
    var empty_xl = try XLValue.fromUtf8String(allocator, "");

    var args: [11][*c]xl.XLOPER12 = undefined;
    args[0] = xll_path;
    args[1] = &proc_name_xl.m_val;
    args[2] = &type_string_xl.m_val;
    args[3] = &func_name_xl.m_val;
    args[4] = &arg_names_xl.m_val;
    args[5] = &func_type_xl.m_val;
    args[6] = &category_xl.m_val;
    args[7] = &empty_xl.m_val; // Shortcut
    args[8] = &empty_xl.m_val; // Help topic
    args[9] = &description_xl.m_val;
    args[10] = &empty_xl.m_val; // Trailing empty

    var result: xl.XLOPER12 = undefined;
    const ret = xl.Excel12v(xl.xlfRegister, &result, 11, @ptrCast(&args));
    defer xl_helpers.xlFree(&result);

    if (ret != xl.xlretSuccess) {
        return error.RegistrationFailed;
    }

    xl_helpers.debugLogFmt("Registered Excel macro: {s}", .{MacroType.excel_name});
}

pub fn xlAutoClose() callconv(.c) c_int {
    xl_helpers.debugLog("xlAutoClose called");

    // Clean up Lua state if it was initialized
    const user_mod = @import("root").user_functions;
    if (comptime @hasDecl(user_mod, "lua_scripts")) {
        const lua = @import("lua.zig");
        lua.deinit();
    }

    // Call user-defined cleanup if present
    const root = @import("root");
    if (comptime @hasDecl(root, "deinit")) {
        root.deinit();
    }

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
