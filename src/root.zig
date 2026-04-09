// Re-export framework components for external users
pub const framework_entry = @import("framework_entry.zig");
pub const excel_function = @import("excel_function.zig");
pub const excel_macro = @import("excel_macro.zig");
pub const xl_imports = @import("xl_imports.zig");
pub const xlvalue = @import("xlvalue.zig");
pub const xl_helpers = @import("xl_helpers.zig");
pub const function_discovery = @import("function_discovery.zig");

pub const rtd = @import("rtd.zig");
pub const rtd_registry = @import("rtd_registry.zig");
pub const rtd_call = @import("rtd_call.zig");
pub const async_cache = @import("async_cache.zig");
pub const async_handler = @import("async_handler.zig");
pub const async_infra = @import("async_infra.zig");

// Convenience exports
pub const xl = xl_imports.xl;
pub const XLValue = xlvalue.XLValue;
pub const ExcelFunction = excel_function.ExcelFunction;
pub const ExcelMacro = excel_macro.ExcelMacro;
pub const ParamMeta = excel_function.ParamMeta;
pub const AsyncContext = async_infra.AsyncContext;
pub const AsyncValue = async_infra.AsyncValue;

// Lua scripting support
pub const lua = @import("lua.zig");
pub const lua_function = @import("lua_function.zig");
pub const LuaFunction = lua_function.LuaFunction;
pub const LuaParam = lua_function.LuaParam;
pub const LuaParamType = lua_function.LuaParamType;
pub const lua_rtd_function = @import("lua_rtd_function.zig");
pub const LuaRtdFunction = lua_rtd_function.LuaRtdFunction;
