// Re-export framework components for external users
pub const framework_entry = @import("framework_entry.zig");
pub const excel_function = @import("excel_function.zig");
pub const xl_imports = @import("xl_imports.zig");
pub const xlvalue = @import("core/xlvalue.zig");
pub const xl_helpers = @import("xl_helpers.zig");
pub const function_discovery = @import("function_discovery.zig");

pub const rtd = @import("rtd.zig");
pub const rtd_registry = @import("rtd_registry.zig");
pub const rtd_call = @import("rtd_call.zig");

// Convenience exports
pub const xl = xl_imports.xl;
pub const XLValue = xlvalue.XLValue;
pub const ExcelFunction = excel_function.ExcelFunction;
pub const ParamMeta = excel_function.ParamMeta;
