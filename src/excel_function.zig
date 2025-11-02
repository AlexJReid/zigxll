// Excel function registration helper with automatic type conversion
const std = @import("std");
const xl_imports = @import("xl_imports.zig");
const xl = xl_imports.xl;
const xlvalue = @import("xlvalue.zig");
const XLValue = xlvalue.XLValue;

const allocator = std.heap.c_allocator;

/// Parameter metadata for Excel function arguments
pub const ParamMeta = struct {
    name: ?[]const u8 = null,
    description: ?[]const u8 = null,
};

pub fn ExcelFunction(comptime meta: anytype) type {
    const name = meta.name;
    const description = if (@hasField(@TypeOf(meta), "description")) meta.description else "";
    const category = if (@hasField(@TypeOf(meta), "category")) meta.category else "General";
    const func = meta.func;
    const params_meta = if (@hasField(@TypeOf(meta), "params")) meta.params else &[_]ParamMeta{};
    const thread_safe = if (@hasField(@TypeOf(meta), "thread_safe")) meta.thread_safe else true;

    const func_info = @typeInfo(@TypeOf(func));
    const params = switch (func_info) {
        .@"fn" => |f| f.params,
        else => @compileError("Expected function type"),
    };

    // Validate params metadata matches function signature
    comptime {
        if (params_meta.len > 0 and params_meta.len != params.len) {
            @compileError(std.fmt.comptimePrint("Function '{s}' has {d} parameters but {d} parameter descriptions provided", .{ name, params.len, params_meta.len }));
        }
    }

    // Generate Excel type string at comptime
    const type_string = comptime blk: {
        // Return type (always Q for XLOPER12 pointer)
        var result: []const u8 = "Q";

        // Parameter types (all Q for XLOPER12 pointer)
        var i: usize = 0;
        while (i < params.len) : (i += 1) {
            result = result ++ "Q";
        }

        // Add $ suffix if thread safe
        if (thread_safe) {
            result = result ++ "$";
        }

        break :blk result;
    };

    // Generate unique export name based on function name
    const export_name = comptime blk: {
        break :blk name ++ "_impl"; // not sure if should be "more" unique
    };

    return struct {
        pub const excel_name = name;
        pub const excel_description = description;
        pub const excel_category = category;
        pub const excel_params = params_meta;
        pub const excel_param_count = params.len;
        pub const excel_type_string = type_string;
        pub const excel_thread_safe = thread_safe;
        pub const is_excel_function = true;
        pub const excel_export_name = export_name;

        fn makeErrorValue() *xl.XLOPER12 {
            var err_val = XLValue.err(allocator, xl.xlerrValue);
            return err_val.get();
        }

        const Impl = switch (params.len) {
            0 => struct {
                fn impl() callconv(.c) *xl.XLOPER12 {
                    const result = func() catch return makeErrorValue();
                    return wrapResult(result);
                }
            },
            1 => struct {
                fn impl(a1: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    defer freeArg(params[0].type.?, arg1);
                    const result = func(arg1) catch return makeErrorValue();
                    return wrapResult(result);
                }
            },
            2 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    defer freeArg(params[0].type.?, arg1);
                    const arg2 = extractArg(params[1].type.?, a2) catch return makeErrorValue();
                    defer freeArg(params[1].type.?, arg2);
                    const result = func(arg1, arg2) catch return makeErrorValue();
                    return wrapResult(result);
                }
            },
            3 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    defer freeArg(params[0].type.?, arg1);
                    const arg2 = extractArg(params[1].type.?, a2) catch return makeErrorValue();
                    defer freeArg(params[1].type.?, arg2);
                    const arg3 = extractArg(params[2].type.?, a3) catch return makeErrorValue();
                    defer freeArg(params[2].type.?, arg3);
                    const result = func(arg1, arg2, arg3) catch return makeErrorValue();
                    return wrapResult(result);
                }
            },
            4 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    defer freeArg(params[0].type.?, arg1);
                    const arg2 = extractArg(params[1].type.?, a2) catch return makeErrorValue();
                    defer freeArg(params[1].type.?, arg2);
                    const arg3 = extractArg(params[2].type.?, a3) catch return makeErrorValue();
                    defer freeArg(params[2].type.?, arg3);
                    const arg4 = extractArg(params[3].type.?, a4) catch return makeErrorValue();
                    defer freeArg(params[3].type.?, arg4);
                    const result = func(arg1, arg2, arg3, arg4) catch return makeErrorValue();
                    return wrapResult(result);
                }
            },
            5 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    defer freeArg(params[0].type.?, arg1);
                    const arg2 = extractArg(params[1].type.?, a2) catch return makeErrorValue();
                    defer freeArg(params[1].type.?, arg2);
                    const arg3 = extractArg(params[2].type.?, a3) catch return makeErrorValue();
                    defer freeArg(params[2].type.?, arg3);
                    const arg4 = extractArg(params[3].type.?, a4) catch return makeErrorValue();
                    defer freeArg(params[3].type.?, arg4);
                    const arg5 = extractArg(params[4].type.?, a5) catch return makeErrorValue();
                    defer freeArg(params[4].type.?, arg5);
                    const result = func(arg1, arg2, arg3, arg4, arg5) catch return makeErrorValue();
                    return wrapResult(result);
                }
            },
            6 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12, a6: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    defer freeArg(params[0].type.?, arg1);
                    const arg2 = extractArg(params[1].type.?, a2) catch return makeErrorValue();
                    defer freeArg(params[1].type.?, arg2);
                    const arg3 = extractArg(params[2].type.?, a3) catch return makeErrorValue();
                    defer freeArg(params[2].type.?, arg3);
                    const arg4 = extractArg(params[3].type.?, a4) catch return makeErrorValue();
                    defer freeArg(params[3].type.?, arg4);
                    const arg5 = extractArg(params[4].type.?, a5) catch return makeErrorValue();
                    defer freeArg(params[4].type.?, arg5);
                    const arg6 = extractArg(params[5].type.?, a6) catch return makeErrorValue();
                    defer freeArg(params[5].type.?, arg6);
                    const result = func(arg1, arg2, arg3, arg4, arg5, arg6) catch return makeErrorValue();
                    return wrapResult(result);
                }
            },
            7 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12, a6: *xl.XLOPER12, a7: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    defer freeArg(params[0].type.?, arg1);
                    const arg2 = extractArg(params[1].type.?, a2) catch return makeErrorValue();
                    defer freeArg(params[1].type.?, arg2);
                    const arg3 = extractArg(params[2].type.?, a3) catch return makeErrorValue();
                    defer freeArg(params[2].type.?, arg3);
                    const arg4 = extractArg(params[3].type.?, a4) catch return makeErrorValue();
                    defer freeArg(params[3].type.?, arg4);
                    const arg5 = extractArg(params[4].type.?, a5) catch return makeErrorValue();
                    defer freeArg(params[4].type.?, arg5);
                    const arg6 = extractArg(params[5].type.?, a6) catch return makeErrorValue();
                    defer freeArg(params[5].type.?, arg6);
                    const arg7 = extractArg(params[6].type.?, a7) catch return makeErrorValue();
                    defer freeArg(params[6].type.?, arg7);
                    const result = func(arg1, arg2, arg3, arg4, arg5, arg6, arg7) catch return makeErrorValue();
                    return wrapResult(result);
                }
            },
            8 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12, a6: *xl.XLOPER12, a7: *xl.XLOPER12, a8: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    defer freeArg(params[0].type.?, arg1);
                    const arg2 = extractArg(params[1].type.?, a2) catch return makeErrorValue();
                    defer freeArg(params[1].type.?, arg2);
                    const arg3 = extractArg(params[2].type.?, a3) catch return makeErrorValue();
                    defer freeArg(params[2].type.?, arg3);
                    const arg4 = extractArg(params[3].type.?, a4) catch return makeErrorValue();
                    defer freeArg(params[3].type.?, arg4);
                    const arg5 = extractArg(params[4].type.?, a5) catch return makeErrorValue();
                    defer freeArg(params[4].type.?, arg5);
                    const arg6 = extractArg(params[5].type.?, a6) catch return makeErrorValue();
                    defer freeArg(params[5].type.?, arg6);
                    const arg7 = extractArg(params[6].type.?, a7) catch return makeErrorValue();
                    defer freeArg(params[6].type.?, arg7);
                    const arg8 = extractArg(params[7].type.?, a8) catch return makeErrorValue();
                    defer freeArg(params[7].type.?, arg8);
                    const result = func(arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8) catch return makeErrorValue();
                    return wrapResult(result);
                }
            },
            else => @compileError("Unsupported parameter count (max 8)"),
        };

        // Store reference to the impl function
        pub const impl = Impl.impl;

        fn extractArg(comptime T: type, xloper: *xl.XLOPER12) !T {
            const val = XLValue.fromXLOPER12(allocator, xloper.*, false);

            // Handle different types
            if (T == f64) {
                return try val.as_double();
            } else if (T == []const u8) {
                return try val.as_utf8str();
            } else if (T == *xl.XLOPER12) {
                return xloper;
            } else {
                @compileError("Unsupported parameter type: " ++ @typeName(T));
            }
        }

        fn freeArg(comptime T: type, arg: T) void {
            if (T == []const u8) {
                allocator.free(arg);
            }
        }

        fn wrapResult(result: anytype) *xl.XLOPER12 {
            const T = @TypeOf(result);
            if (T == f64) {
                const val = XLValue.fromDouble(allocator, result);
                // Allocate XLOPER12 on heap so Excel can access it after function returns
                const ret_ptr = allocator.create(xl.XLOPER12) catch {
                    const err_val = XLValue.err(allocator, xl.xlerrValue);
                    const err_ptr = allocator.create(xl.XLOPER12) catch unreachable;
                    err_ptr.* = err_val.m_val;
                    err_ptr.xltype |= xl.xlbitDLLFree;
                    return err_ptr;
                };
                ret_ptr.* = val.m_val;
                ret_ptr.xltype |= xl.xlbitDLLFree;
                return ret_ptr;
            } else if (T == []const u8 or T == []u8) {
                defer allocator.free(result);
                const val = XLValue.fromUtf8String(allocator, result) catch {
                    const err_val = XLValue.err(allocator, xl.xlerrValue);
                    const err_ptr = allocator.create(xl.XLOPER12) catch unreachable;
                    err_ptr.* = err_val.m_val;
                    err_ptr.xltype |= xl.xlbitDLLFree;
                    return err_ptr;
                };
                // Allocate XLOPER12 on heap so Excel can access it after function returns
                const ret_ptr = allocator.create(xl.XLOPER12) catch {
                    const err_val = XLValue.err(allocator, xl.xlerrValue);
                    const err_ptr = allocator.create(xl.XLOPER12) catch unreachable;
                    err_ptr.* = err_val.m_val;
                    err_ptr.xltype |= xl.xlbitDLLFree;
                    return err_ptr;
                };
                ret_ptr.* = val.m_val;
                ret_ptr.xltype |= xl.xlbitDLLFree;
                return ret_ptr;
            } else if (T == *xl.XLOPER12) {
                return result;
            } else {
                @compileError("Unsupported return type: " ++ @typeName(T));
            }
        }

        // Use @export to give it the custom name
        comptime {
            @export(&Impl.impl, .{ .name = export_name });
        }
    };
}
