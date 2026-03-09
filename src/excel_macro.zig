// Excel macro (command) registration helper
const std = @import("std");
const xl_imports = @import("xl_imports.zig");
const xl = xl_imports.xl;

fn sanitizeExportName(comptime len: usize, comptime input: *const [len]u8) *const [len]u8 {
    comptime {
        var result: [len]u8 = input.*;
        for (&result) |*c| {
            if (c.* == '.') c.* = '_';
        }
        const final = result;
        return &final;
    }
}

pub fn ExcelMacro(comptime meta: anytype) type {
    const name = meta.name;
    const description = if (@hasField(@TypeOf(meta), "description")) meta.description else "";
    const category = if (@hasField(@TypeOf(meta), "category")) meta.category else "General";
    const func = meta.func;

    // Validate: macro function must take no parameters and return !void or void
    const func_info = @typeInfo(@TypeOf(func));
    switch (func_info) {
        .@"fn" => |f| {
            if (f.params.len != 0) {
                @compileError("Excel macros must take no parameters");
            }
        },
        else => @compileError("Expected function type"),
    }

    const export_name = comptime blk: {
        break :blk sanitizeExportName(name.len, name) ++ "_impl";
    };

    return struct {
        pub const excel_name = name;
        pub const excel_description = description;
        pub const excel_category = category;
        pub const excel_type_string = "A";
        pub const excel_export_name = export_name;
        pub const is_excel_macro = true;

        const Impl = struct {
            fn impl() callconv(.c) c_int {
                const maybe_error = @typeInfo(@typeInfo(@TypeOf(func)).@"fn".return_type.?);
                if (maybe_error == .error_union) {
                    func() catch return 0;
                } else {
                    func();
                }
                return 1;
            }
        };

        pub const impl = Impl.impl;

        comptime {
            @export(&Impl.impl, .{ .name = export_name });
        }
    };
}
