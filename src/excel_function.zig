// Excel function registration helper with automatic type conversion
const std = @import("std");
const xl_imports = @import("xl_imports.zig");
const xl = xl_imports.xl;
const XLValue = @import("xlvalue.zig").XLValue;

const allocator = std.heap.c_allocator;
const xl_helpers = @import("xl_helpers.zig");

// Async support
const async_cache = @import("async_cache.zig");
const async_infra = @import("async_infra.zig");

/// Parameter metadata for Excel function arguments
pub const ParamMeta = struct {
    name: ?[]const u8 = null,
    description: ?[]const u8 = null,
};

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

pub fn ExcelFunction(comptime meta: anytype) type {
    const name = meta.name;
    const description = if (@hasField(@TypeOf(meta), "description")) meta.description else "";
    const category = if (@hasField(@TypeOf(meta), "category")) meta.category else "General";
    const func = meta.func;
    const params_meta = if (@hasField(@TypeOf(meta), "params")) meta.params else &[_]ParamMeta{};
    const is_async = if (@hasField(@TypeOf(meta), "async")) meta.async else false;
    // Async functions must NOT be thread-safe (they call xlfRtd which isn't thread-safe)
    const thread_safe = if (is_async) false else (if (@hasField(@TypeOf(meta), "thread_safe")) meta.thread_safe else true);

    const func_info = @typeInfo(@TypeOf(func));
    const all_params = switch (func_info) {
        .@"fn" => |f| f.params,
        else => @compileError("Expected function type"),
    };

    // Check if the last parameter is *AsyncContext (yield support)
    const has_yield = comptime blk: {
        if (!is_async or all_params.len == 0) break :blk false;
        const LastType = all_params[all_params.len - 1].type.?;
        break :blk LastType == *async_infra.AsyncContext;
    };

    // Excel-visible params exclude the trailing *AsyncContext if present
    const excel_param_len = if (has_yield) all_params.len - 1 else all_params.len;
    const params = all_params[0..excel_param_len];

    // Validate params metadata matches Excel-visible parameters
    comptime {
        if (params_meta.len > 0 and params_meta.len != params.len) {
            @compileError(std.fmt.comptimePrint("Function '{s}' has {d} Excel parameters but {d} parameter descriptions provided", .{ name, params.len, params_meta.len }));
        }
    }

    // Build an array of Excel-visible parameter types for async helpers
    const ParamTypes = comptime blk: {
        var types: [params.len]type = undefined;
        for (0..params.len) |i| {
            types[i] = params[i].type.?;
        }
        break :blk types;
    };

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
    // Replace dots with underscores to avoid Windows GetProcAddress issues
    const export_name = comptime blk: {
        break :blk sanitizeExportName(name.len, name) ++ "_impl";
    };

    return struct {
        pub const excel_name = name;
        pub const excel_description = description;
        pub const excel_category = category;
        pub const excel_params = params_meta;
        pub const excel_param_count = excel_param_len;
        pub const excel_type_string = type_string;
        pub const excel_thread_safe = thread_safe;
        pub const excel_is_async = is_async;
        pub const is_excel_function = true;
        pub const excel_export_name = export_name;

        fn makeErrorValue() *xl.XLOPER12 {
            const err_ptr = allocator.create(xl.XLOPER12) catch unreachable;
            err_ptr.* = .{
                .xltype = xl.xltypeErr | xl.xlbitDLLFree,
                .val = .{ .err = xl.xlerrValue },
            };
            return err_ptr;
        }

        // ================================================================
        // Sync implementation (same as before)
        // ================================================================

        fn callSync(args: anytype) *xl.XLOPER12 {
            const result = @call(.auto, func, args) catch return makeErrorValue();
            return wrapResult(result);
        }

        // ================================================================
        // Async support: worker task + cache logic
        // ================================================================

        /// Heap-allocated argument pack for the worker thread.
        /// Only contains the Excel-visible args (not *AsyncContext).
        const ExcelArgsTuple = blk: {
            var fields: [ParamTypes.len]type = undefined;
            for (0..ParamTypes.len) |i| {
                fields[i] = ParamTypes[i];
            }
            break :blk std.meta.Tuple(&fields);
        };

        const AsyncArgs = struct {
            key: []const u8,
            args: ExcelArgsTuple,
        };

        /// Worker function spawned on the thread pool.
        /// Calls the user function, stores result in cache, notifies Excel.
        fn asyncWorker(pack: *AsyncArgs) void {
            defer {
                inline for (0..ParamTypes.len) |i| {
                    async_infra.freeOwnedArg(ParamTypes[i], pack.args[i]);
                }
                allocator.free(pack.key);
                allocator.destroy(pack);
            }

            if (has_yield) {
                var ctx = async_infra.AsyncContext{ .key = pack.key };
                var call_args: std.meta.ArgsTuple(@TypeOf(func)) = undefined;
                inline for (0..ParamTypes.len) |i| {
                    call_args[i] = pack.args[i];
                }
                call_args[ParamTypes.len] = &ctx;
                const result = @call(.auto, func, call_args) catch {
                    async_infra.storeResult(pack.key, makeErrorValue());
                    return;
                };
                async_infra.storeResult(pack.key, wrapResult(result));
            } else {
                var call_args: std.meta.ArgsTuple(@TypeOf(func)) = undefined;
                inline for (0..ParamTypes.len) |i| {
                    call_args[i] = pack.args[i];
                }
                const result = @call(.auto, func, call_args) catch {
                    async_infra.storeResult(pack.key, makeErrorValue());
                    return;
                };
                async_infra.storeResult(pack.key, wrapResult(result));
            }
        }

        /// Async impl: check cache → return cached / spawn + subscribe RTD.
        /// Takes ownership of the extracted args (frees them before returning).
        fn asyncImpl(extracted: ExcelArgsTuple) *xl.XLOPER12 {
            // Build topic key from function name + serialized args
            const key = async_infra.buildTopicKey(name, &ParamTypes, extracted) catch {
                freeExtracted(extracted);
                return makeErrorValue();
            };
            defer allocator.free(key);

            const cache = async_cache.getGlobalCache();

            // Cache hit?
            if (cache.get(key)) |entry| {
                freeExtracted(extracted);
                if (entry.completed) {
                    return async_infra.cloneXloper(entry.xloper);
                }
                return async_infra.rtdSubscribe(key) catch return makeErrorValue();
            }

            // Cache miss — spawn async work.
            // Dupe args for the worker thread, then free originals.
            const pack = allocator.create(AsyncArgs) catch {
                freeExtracted(extracted);
                return makeErrorValue();
            };
            var duped_args: ExcelArgsTuple = undefined;
            inline for (0..ParamTypes.len) |i| {
                duped_args[i] = async_infra.dupeArg(ParamTypes[i], extracted[i]);
            }
            // Worker gets its own copy of the key
            const worker_key = allocator.dupe(u8, key) catch {
                inline for (0..ParamTypes.len) |i| {
                    async_infra.freeOwnedArg(ParamTypes[i], duped_args[i]);
                }
                allocator.destroy(pack);
                freeExtracted(extracted);
                return makeErrorValue();
            };
            pack.* = .{
                .key = worker_key,
                .args = duped_args,
            };

            freeExtracted(extracted);

            // Mark in-progress in cache (cache dupes the key internally)
            async_infra.markInProgress(key);

            // Subscribe via RTD FIRST — this triggers ServerStart which
            // sets up the global update_event pointer.  Must happen before
            // spawning the worker, otherwise the worker could call
            // UpdateNotify while Excel is still inside xlfRtd.
            const rtd_result = async_infra.rtdSubscribe(key) catch {
                inline for (0..ParamTypes.len) |i| {
                    async_infra.freeOwnedArg(ParamTypes[i], pack.args[i]);
                }
                allocator.free(pack.key);
                allocator.destroy(pack);
                return makeErrorValue();
            };
            const handle = async_infra.createWorkerThread(AsyncArgs, asyncWorker, pack);
            if (handle == null) {
                inline for (0..ParamTypes.len) |i| {
                    async_infra.freeOwnedArg(ParamTypes[i], pack.args[i]);
                }
                allocator.free(pack.key);
                allocator.destroy(pack);
            }

            return rtd_result;
        }

        fn freeExtracted(extracted: ExcelArgsTuple) void {
            inline for (0..ParamTypes.len) |i| {
                freeArg(ParamTypes[i], extracted[i]);
            }
        }

        // ================================================================
        // Generated C-callable impl (dispatches to sync or async)
        // ================================================================

        const Impl = switch (params.len) {
            0 => struct {
                fn impl() callconv(.c) *xl.XLOPER12 {
                    if (is_async) {
                        return asyncImpl(.{});
                    }
                    return callSync(.{});
                }
            },
            1 => struct {
                fn impl(a1: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    if (is_async) {
                        return asyncImpl(.{arg1});
                    }
                    defer freeArg(params[0].type.?, arg1);
                    return callSync(.{arg1});
                }
            },
            2 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    const arg2 = extractArg(params[1].type.?, a2) catch |e| {
                        freeArg(params[0].type.?, arg1);
                        return if (e == error.OutOfMemory) makeErrorValue() else makeErrorValue();
                    };
                    if (is_async) return asyncImpl(.{ arg1, arg2 });
                    defer freeArg(params[0].type.?, arg1);
                    defer freeArg(params[1].type.?, arg2);
                    return callSync(.{ arg1, arg2 });
                }
            },
            3 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    const arg2 = extractArg(params[1].type.?, a2) catch {
                        freeArg(params[0].type.?, arg1);
                        return makeErrorValue();
                    };
                    const arg3 = extractArg(params[2].type.?, a3) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        return makeErrorValue();
                    };
                    if (is_async) return asyncImpl(.{ arg1, arg2, arg3 });
                    defer freeArg(params[0].type.?, arg1);
                    defer freeArg(params[1].type.?, arg2);
                    defer freeArg(params[2].type.?, arg3);
                    return callSync(.{ arg1, arg2, arg3 });
                }
            },
            4 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    const arg2 = extractArg(params[1].type.?, a2) catch {
                        freeArg(params[0].type.?, arg1);
                        return makeErrorValue();
                    };
                    const arg3 = extractArg(params[2].type.?, a3) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        return makeErrorValue();
                    };
                    const arg4 = extractArg(params[3].type.?, a4) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        freeArg(params[2].type.?, arg3);
                        return makeErrorValue();
                    };
                    if (is_async) return asyncImpl(.{ arg1, arg2, arg3, arg4 });
                    defer freeArg(params[0].type.?, arg1);
                    defer freeArg(params[1].type.?, arg2);
                    defer freeArg(params[2].type.?, arg3);
                    defer freeArg(params[3].type.?, arg4);
                    return callSync(.{ arg1, arg2, arg3, arg4 });
                }
            },
            5 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    const arg2 = extractArg(params[1].type.?, a2) catch {
                        freeArg(params[0].type.?, arg1);
                        return makeErrorValue();
                    };
                    const arg3 = extractArg(params[2].type.?, a3) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        return makeErrorValue();
                    };
                    const arg4 = extractArg(params[3].type.?, a4) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        freeArg(params[2].type.?, arg3);
                        return makeErrorValue();
                    };
                    const arg5 = extractArg(params[4].type.?, a5) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        freeArg(params[2].type.?, arg3);
                        freeArg(params[3].type.?, arg4);
                        return makeErrorValue();
                    };
                    if (is_async) return asyncImpl(.{ arg1, arg2, arg3, arg4, arg5 });
                    defer freeArg(params[0].type.?, arg1);
                    defer freeArg(params[1].type.?, arg2);
                    defer freeArg(params[2].type.?, arg3);
                    defer freeArg(params[3].type.?, arg4);
                    defer freeArg(params[4].type.?, arg5);
                    return callSync(.{ arg1, arg2, arg3, arg4, arg5 });
                }
            },
            6 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12, a6: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    const arg2 = extractArg(params[1].type.?, a2) catch {
                        freeArg(params[0].type.?, arg1);
                        return makeErrorValue();
                    };
                    const arg3 = extractArg(params[2].type.?, a3) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        return makeErrorValue();
                    };
                    const arg4 = extractArg(params[3].type.?, a4) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        freeArg(params[2].type.?, arg3);
                        return makeErrorValue();
                    };
                    const arg5 = extractArg(params[4].type.?, a5) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        freeArg(params[2].type.?, arg3);
                        freeArg(params[3].type.?, arg4);
                        return makeErrorValue();
                    };
                    const arg6 = extractArg(params[5].type.?, a6) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        freeArg(params[2].type.?, arg3);
                        freeArg(params[3].type.?, arg4);
                        freeArg(params[4].type.?, arg5);
                        return makeErrorValue();
                    };
                    if (is_async) return asyncImpl(.{ arg1, arg2, arg3, arg4, arg5, arg6 });
                    defer freeArg(params[0].type.?, arg1);
                    defer freeArg(params[1].type.?, arg2);
                    defer freeArg(params[2].type.?, arg3);
                    defer freeArg(params[3].type.?, arg4);
                    defer freeArg(params[4].type.?, arg5);
                    defer freeArg(params[5].type.?, arg6);
                    return callSync(.{ arg1, arg2, arg3, arg4, arg5, arg6 });
                }
            },
            7 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12, a6: *xl.XLOPER12, a7: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    const arg2 = extractArg(params[1].type.?, a2) catch {
                        freeArg(params[0].type.?, arg1);
                        return makeErrorValue();
                    };
                    const arg3 = extractArg(params[2].type.?, a3) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        return makeErrorValue();
                    };
                    const arg4 = extractArg(params[3].type.?, a4) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        freeArg(params[2].type.?, arg3);
                        return makeErrorValue();
                    };
                    const arg5 = extractArg(params[4].type.?, a5) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        freeArg(params[2].type.?, arg3);
                        freeArg(params[3].type.?, arg4);
                        return makeErrorValue();
                    };
                    const arg6 = extractArg(params[5].type.?, a6) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        freeArg(params[2].type.?, arg3);
                        freeArg(params[3].type.?, arg4);
                        freeArg(params[4].type.?, arg5);
                        return makeErrorValue();
                    };
                    const arg7 = extractArg(params[6].type.?, a7) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        freeArg(params[2].type.?, arg3);
                        freeArg(params[3].type.?, arg4);
                        freeArg(params[4].type.?, arg5);
                        freeArg(params[5].type.?, arg6);
                        return makeErrorValue();
                    };
                    if (is_async) return asyncImpl(.{ arg1, arg2, arg3, arg4, arg5, arg6, arg7 });
                    defer freeArg(params[0].type.?, arg1);
                    defer freeArg(params[1].type.?, arg2);
                    defer freeArg(params[2].type.?, arg3);
                    defer freeArg(params[3].type.?, arg4);
                    defer freeArg(params[4].type.?, arg5);
                    defer freeArg(params[5].type.?, arg6);
                    defer freeArg(params[6].type.?, arg7);
                    return callSync(.{ arg1, arg2, arg3, arg4, arg5, arg6, arg7 });
                }
            },
            8 => struct {
                fn impl(a1: *xl.XLOPER12, a2: *xl.XLOPER12, a3: *xl.XLOPER12, a4: *xl.XLOPER12, a5: *xl.XLOPER12, a6: *xl.XLOPER12, a7: *xl.XLOPER12, a8: *xl.XLOPER12) callconv(.c) *xl.XLOPER12 {
                    const arg1 = extractArg(params[0].type.?, a1) catch return makeErrorValue();
                    const arg2 = extractArg(params[1].type.?, a2) catch {
                        freeArg(params[0].type.?, arg1);
                        return makeErrorValue();
                    };
                    const arg3 = extractArg(params[2].type.?, a3) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        return makeErrorValue();
                    };
                    const arg4 = extractArg(params[3].type.?, a4) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        freeArg(params[2].type.?, arg3);
                        return makeErrorValue();
                    };
                    const arg5 = extractArg(params[4].type.?, a5) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        freeArg(params[2].type.?, arg3);
                        freeArg(params[3].type.?, arg4);
                        return makeErrorValue();
                    };
                    const arg6 = extractArg(params[5].type.?, a6) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        freeArg(params[2].type.?, arg3);
                        freeArg(params[3].type.?, arg4);
                        freeArg(params[4].type.?, arg5);
                        return makeErrorValue();
                    };
                    const arg7 = extractArg(params[6].type.?, a7) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        freeArg(params[2].type.?, arg3);
                        freeArg(params[3].type.?, arg4);
                        freeArg(params[4].type.?, arg5);
                        freeArg(params[5].type.?, arg6);
                        return makeErrorValue();
                    };
                    const arg8 = extractArg(params[7].type.?, a8) catch {
                        freeArg(params[0].type.?, arg1);
                        freeArg(params[1].type.?, arg2);
                        freeArg(params[2].type.?, arg3);
                        freeArg(params[3].type.?, arg4);
                        freeArg(params[4].type.?, arg5);
                        freeArg(params[5].type.?, arg6);
                        freeArg(params[6].type.?, arg7);
                        return makeErrorValue();
                    };
                    if (is_async) return asyncImpl(.{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 });
                    defer freeArg(params[0].type.?, arg1);
                    defer freeArg(params[1].type.?, arg2);
                    defer freeArg(params[2].type.?, arg3);
                    defer freeArg(params[3].type.?, arg4);
                    defer freeArg(params[4].type.?, arg5);
                    defer freeArg(params[5].type.?, arg6);
                    defer freeArg(params[6].type.?, arg7);
                    defer freeArg(params[7].type.?, arg8);
                    return callSync(.{ arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8 });
                }
            },
            else => @compileError("Unsupported parameter count (max 8)"),
        };

        // Store reference to the impl function
        pub const impl = Impl.impl;

        fn extractArg(comptime T: type, xloper: *xl.XLOPER12) !T {
            const val = XLValue.fromXLOPER12(allocator, xloper.*, false);

            // Check if T is an optional type
            const type_info = @typeInfo(T);
            if (type_info == .optional) {
                if (val.is_missing()) {
                    return null;
                }
                const Child = type_info.optional.child;
                return try extractNonOptional(Child, val);
            }

            return try extractNonOptional(T, val);
        }

        fn extractNonOptional(comptime T: type, val: XLValue) !T {
            if (T == f64) {
                return try val.as_double();
            } else if (T == bool) {
                return try val.as_bool();
            } else if (T == []const u8) {
                return try val.as_utf8str();
            } else if (T == [][]const f64) {
                return try val.as_matrix();
            } else if (T == *xl.XLOPER12) {
                @compileError("Optional XLOPER12 pointers are not supported");
            } else {
                @compileError("Unsupported parameter type: " ++ @typeName(T));
            }
        }

        fn freeArg(comptime T: type, arg: T) void {
            // Check if T is an optional type
            const type_info = @typeInfo(T);
            if (type_info == .optional) {
                if (arg) |value| {
                    // Free the unwrapped value if needed
                    const Child = type_info.optional.child;
                    freeNonOptional(Child, value);
                }
                return;
            }

            // Handle non-optional types
            freeNonOptional(T, arg);
        }

        fn freeNonOptional(comptime T: type, arg: T) void {
            if (T == []const u8) {
                allocator.free(arg);
            } else if (T == [][]const f64) {
                for (arg) |row| {
                    allocator.free(row);
                }
                allocator.free(arg);
            }
            // Other types don't need explicit freeing
        }

        fn heapXloper(xloper: xl.XLOPER12) *xl.XLOPER12 {
            const ptr = allocator.create(xl.XLOPER12) catch return makeErrorValue();
            ptr.* = xloper;
            ptr.xltype |= xl.xlbitDLLFree;
            return ptr;
        }

        fn wrapResult(result: anytype) *xl.XLOPER12 {
            const T = @TypeOf(result);
            if (T == f64) {
                return heapXloper(.{
                    .xltype = xl.xltypeNum,
                    .val = .{ .num = result },
                });
            } else if (T == bool) {
                return heapXloper(.{
                    .xltype = xl.xltypeBool,
                    .val = .{ .xbool = if (result) 1 else 0 },
                });
            } else if (T == []const u8 or T == []u8) {
                defer allocator.free(result);
                const val = XLValue.fromUtf8String(allocator, result) catch return makeErrorValue();
                return heapXloper(val.m_val);
            } else if (T == [][]const f64 or T == [][]f64) {
                defer {
                    for (result) |row| {
                        allocator.free(row);
                    }
                    allocator.free(result);
                }
                const val = XLValue.fromMatrix(allocator, result) catch return makeErrorValue();
                return heapXloper(val.m_val);
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

// Tests for optional parameter support
test "optional parameter - f64 with value" {
    // Create an XLOPER12 with a number
    var xloper: xl.XLOPER12 = .{
        .xltype = xl.xltypeNum,
        .val = .{ .num = 42.0 },
    };

    const TestFunc = ExcelFunction(.{
        .name = "TestOptional",
        .func = struct {
            fn impl(x: ?f64) !?f64 {
                return x;
            }
        }.impl,
    });

    const result = try TestFunc.extractArg(?f64, &xloper);
    try std.testing.expect(result != null);
    try std.testing.expectEqual(@as(f64, 42.0), result.?);
}

test "optional parameter - f64 missing" {
    // Create an XLOPER12 with missing type
    var xloper: xl.XLOPER12 = .{
        .xltype = xl.xltypeMissing,
        .val = undefined,
    };

    const TestFunc = ExcelFunction(.{
        .name = "TestOptional",
        .func = struct {
            fn impl(x: ?f64) !?f64 {
                return x;
            }
        }.impl,
    });

    const result = try TestFunc.extractArg(?f64, &xloper);
    try std.testing.expect(result == null);
}

test "optional parameter - bool with value" {
    var xloper: xl.XLOPER12 = .{
        .xltype = xl.xltypeBool,
        .val = .{ .xbool = 1 },
    };

    const TestFunc = ExcelFunction(.{
        .name = "TestOptional",
        .func = struct {
            fn impl(x: ?bool) !?bool {
                return x;
            }
        }.impl,
    });

    const result = try TestFunc.extractArg(?bool, &xloper);
    try std.testing.expect(result != null);
    try std.testing.expectEqual(true, result.?);
}

test "optional parameter - bool missing" {
    var xloper: xl.XLOPER12 = .{
        .xltype = xl.xltypeMissing,
        .val = undefined,
    };

    const TestFunc = ExcelFunction(.{
        .name = "TestOptional",
        .func = struct {
            fn impl(x: ?bool) !?bool {
                return x;
            }
        }.impl,
    });

    const result = try TestFunc.extractArg(?bool, &xloper);
    try std.testing.expect(result == null);
}
