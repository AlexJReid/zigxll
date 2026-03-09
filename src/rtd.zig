// COM RTD server base in pure Zig with no ATL or C++ arrghs.
//
// Usage:
//   const MyRtd = rtd.RtdServer(MyHandler, .{
//       .clsid = my_clsid,
//       .prog_id = "myapp.rtd",
//   });
//   // in main.zig: comptime { _ = @import("my_rtd.zig"); }
//
// Your Handler should provide the following event handlers/lifecycle hooks:
//   fn onStart(ctx: *RtdContext) void
//   fn onConnect(ctx: *RtdContext, topic_id: i32, topic_count: usize) void
//   fn onDisconnect(ctx: *RtdContext, topic_id: i32, topic_count: usize) void
//   fn onRefreshValue(ctx: *RtdContext, topic_id: i32) RtdValue
//   fn onTerminate(ctx: *RtdContext) void
//
// RtdContext gives access to: update_event (for UpdateNotify), topics, and
// a user_data pointer for your own state.

const std = @import("std");
const windows = std.os.windows;

// ============================================================================
// Windows COM types (pub so handlers can use them if needed)
// Magic values ahoy
// ============================================================================

pub const HRESULT = windows.HRESULT;
pub const ULONG = windows.ULONG;
pub const LONG = c_long;
pub const GUID = windows.GUID;

/// Parse a GUID from a string at comptime.
/// Accepts: "A1B2C3D4-E5F6-7890-1234-567890ABCDEF"
///     or:  "{A1B2C3D4-E5F6-7890-1234-567890ABCDEF}"
pub fn guid(comptime s: []const u8) GUID {
    comptime {
        // Strip optional braces
        const raw = if (s.len > 0 and s[0] == '{' and s[s.len - 1] == '}')
            s[1 .. s.len - 1]
        else
            s;

        if (raw.len != 36) @compileError("GUID string must be 36 chars (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)");
        if (raw[8] != '-' or raw[13] != '-' or raw[18] != '-' or raw[23] != '-')
            @compileError("GUID string has wrong dash positions");

        return .{
            .Data1 = hexU32(raw[0..8]),
            .Data2 = hexU16(raw[9..13]),
            .Data3 = hexU16(raw[14..18]),
            .Data4 = .{
                hexU8(raw[19..21]), hexU8(raw[21..23]),
                hexU8(raw[24..26]), hexU8(raw[26..28]),
                hexU8(raw[28..30]), hexU8(raw[30..32]),
                hexU8(raw[32..34]), hexU8(raw[34..36]),
            },
        };
    }
}

fn hexDigit(c: u8) u8 {
    return switch (c) {
        '0'...'9' => c - '0',
        'a'...'f' => c - 'a' + 10,
        'A'...'F' => c - 'A' + 10,
        else => @compileError("invalid hex digit in GUID"),
    };
}

fn hexU8(comptime s: *const [2]u8) u8 {
    return @as(u8, hexDigit(s[0])) << 4 | hexDigit(s[1]);
}

fn hexU16(comptime s: *const [4]u8) u16 {
    return @as(u16, hexU8(s[0..2])) << 8 | hexU8(s[2..4]);
}

fn hexU32(comptime s: *const [8]u8) u32 {
    return @as(u32, hexU16(s[0..4])) << 16 | hexU16(s[4..8]);
}

pub const S_OK: HRESULT = @bitCast(@as(u32, 0));
pub const S_FALSE: HRESULT = @bitCast(@as(u32, 1));
const E_NOINTERFACE: HRESULT = @bitCast(@as(u32, 0x80004002));
const E_FAIL: HRESULT = @bitCast(@as(u32, 0x80004005));
const CLASS_E_CLASSNOTAVAILABLE: HRESULT = @bitCast(@as(u32, 0x80040111));
const CLASS_E_NOAGGREGATION: HRESULT = @bitCast(@as(u32, 0x80040110));
const E_NOTIMPL: HRESULT = @bitCast(@as(u32, 0x80004001));
const DISP_E_UNKNOWNNAME: HRESULT = @bitCast(@as(u32, 0x80020006));
const DISP_E_MEMBERNOTFOUND: HRESULT = @bitCast(@as(u32, 0x80020003));

const VT_EMPTY: u16 = 0;
const VT_I4: u16 = 3;
const VT_R8: u16 = 5;
const VT_BSTR: u16 = 8;
const VT_DISPATCH: u16 = 9;
const VT_ERROR: u16 = 10;
const VT_BOOL: u16 = 11;
const VT_VARIANT: u16 = 12;
const VT_UNKNOWN: u16 = 13;
const VT_ARRAY: u16 = 0x2000;
const VT_BYREF: u16 = 0x4000;

const VARIANT_TRUE: i16 = -1;
const VARIANT_FALSE: i16 = 0;

// Well-known IIDs
const IID_IUnknown = GUID{
    .Data1 = 0x00000000,
    .Data2 = 0x0000,
    .Data3 = 0x0000,
    .Data4 = .{ 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 },
};
const IID_IClassFactory = GUID{
    .Data1 = 0x00000001,
    .Data2 = 0x0000,
    .Data3 = 0x0000,
    .Data4 = .{ 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 },
};
const IID_IDispatch = GUID{
    .Data1 = 0x00020400,
    .Data2 = 0x0000,
    .Data3 = 0x0000,
    .Data4 = .{ 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 },
};
const IID_IRtdServer = GUID{
    .Data1 = 0xEC0E6191,
    .Data2 = 0xDB51,
    .Data3 = 0x11D3,
    .Data4 = .{ 0x8F, 0x3E, 0x00, 0xC0, 0x4F, 0x36, 0x51, 0xB8 },
};

// ============================================================================
// VARIANT / SAFEARRAY
// ============================================================================

pub const VARIANT = extern struct {
    vt: u16,
    reserved1: u16 = 0,
    reserved2: u16 = 0,
    reserved3: u16 = 0,
    data: extern union {
        lval: LONG,
        dval: f64,
        bstrval: ?[*:0]u16,
        boolval: i16,
        scode: c_long,
        ptr: ?*anyopaque,
    } = .{ .ptr = null },
    _pad: u64 = 0,
};

const SAFEARRAYBOUND = extern struct {
    cElements: ULONG,
    lLbound: LONG,
};

const SAFEARRAY = opaque {};

const DISPPARAMS = extern struct {
    rgvarg: ?[*]VARIANT,
    rgdispidNamedArgs: ?[*]c_long,
    cArgs: c_uint,
    cNamedArgs: c_uint,
};

extern "oleaut32" fn SafeArrayCreate(vt: u16, cDims: c_uint, rgsabound: [*]SAFEARRAYBOUND) callconv(.c) ?*SAFEARRAY;
extern "oleaut32" fn SafeArrayPutElement(psa: *SAFEARRAY, rgIndices: [*]LONG, pv: *anyopaque) callconv(.c) HRESULT;
extern "oleaut32" fn VariantInit(pvarg: *VARIANT) callconv(.c) void;
extern "oleaut32" fn SysAllocStringLen(psz: ?[*]const u16, len: c_uint) callconv(.c) ?[*:0]u16;

extern "kernel32" fn OutputDebugStringA(lpOutputString: [*:0]const u8) callconv(.c) void;

pub fn debugLog(comptime fmt_str: []const u8, args: anytype) void {
    var buf: [512]u8 = undefined;
    const msg = std.fmt.bufPrint(&buf, "[ZigRTD] " ++ fmt_str ++ "\x00", args) catch return;
    OutputDebugStringA(@ptrCast(msg.ptr));
}

// ============================================================================
// RtdValue — tagged union for values returned by onRefreshValue
// ============================================================================

pub const RtdValue = union(enum) {
    int: i32,
    double: f64,
    string: []const u16, // UTF-16 slice (not BSTR — framework converts)
    boolean: bool,
    err: i32, // SCODE error code (e.g. DISP_E_PARAMNOTFOUND for #N/A)
    empty,

    /// Convenience: create from a Zig string literal or slice at comptime.
    pub fn fromUtf8(comptime s: []const u8) RtdValue {
        return .{ .string = std.unicode.utf8ToUtf16LeStringLiteral(s) };
    }

    /// #N/A error value.
    pub const na = RtdValue{ .err = @bitCast(@as(u32, 0x80020004)) }; // DISP_E_PARAMNOTFOUND
};

pub fn rtdValueToVariant(val: RtdValue, out: *VARIANT) void {
    VariantInit(out);
    switch (val) {
        .int => |v| {
            out.vt = VT_I4;
            out.data = .{ .lval = v };
        },
        .double => |v| {
            out.vt = VT_R8;
            out.data = .{ .dval = v };
        },
        .string => |v| {
            const bstr = SysAllocStringLen(v.ptr, @intCast(v.len));
            out.vt = VT_BSTR;
            out.data = .{ .bstrval = bstr };
        },
        .boolean => |v| {
            out.vt = VT_BOOL;
            out.data = .{ .boolval = if (v) VARIANT_TRUE else VARIANT_FALSE };
        },
        .err => |v| {
            out.vt = VT_ERROR;
            out.data = .{ .scode = v };
        },
        .empty => {
            out.vt = VT_EMPTY;
        },
    }
}

// ============================================================================
// IRTDUpdateEvent (Excel's callback interface)
// ============================================================================

const IRTDUpdateEvent_VTable = extern struct {
    QueryInterface: *const fn (*anyopaque, *const GUID, *?*anyopaque) callconv(.winapi) HRESULT,
    AddRef: *const fn (*anyopaque) callconv(.winapi) ULONG,
    Release: *const fn (*anyopaque) callconv(.winapi) ULONG,
    GetTypeInfoCount: *const fn (*anyopaque, *c_uint) callconv(.winapi) HRESULT,
    GetTypeInfo: *const fn (*anyopaque, c_uint, u32, *?*anyopaque) callconv(.winapi) HRESULT,
    GetIDsOfNames: *const fn (*anyopaque, *const GUID, [*]const [*:0]const u16, c_uint, u32, [*]c_long) callconv(.winapi) HRESULT,
    Invoke: *const fn (*anyopaque, c_long, *const GUID, u32, u16, *anyopaque, *anyopaque, *anyopaque, *anyopaque) callconv(.winapi) HRESULT,
    UpdateNotify: *const fn (*anyopaque) callconv(.winapi) HRESULT,
    HeartbeatInterval_get: *const fn (*anyopaque, *LONG) callconv(.winapi) HRESULT,
    HeartbeatInterval_put: *const fn (*anyopaque, LONG) callconv(.winapi) HRESULT,
    Disconnect: *const fn (*anyopaque) callconv(.winapi) HRESULT,
};

pub const IRTDUpdateEvent = extern struct {
    vtable: *const IRTDUpdateEvent_VTable,

    pub fn updateNotify(self: *IRTDUpdateEvent) void {
        _ = self.vtable.UpdateNotify(@ptrCast(self));
    }
};

// ============================================================================
// Topic tracking
// ============================================================================

pub const TopicEntry = struct {
    active: bool = false,
    topic_id: LONG = 0,
    dirty: bool = false,
};

pub const MAX_TOPICS = 16384;

// ============================================================================
// RtdContext — passed to handler callbacks
// ============================================================================

pub const RtdContext = struct {
    update_event: ?*IRTDUpdateEvent = null,
    topics: [MAX_TOPICS]TopicEntry = [_]TopicEntry{.{}} ** MAX_TOPICS,
    topic_count: usize = 0,
    user_data: ?*anyopaque = null,

    /// Call UpdateNotify on Excel's callback to trigger RefreshData.
    pub fn notifyExcel(self: *RtdContext) void {
        if (self.update_event) |evt| {
            evt.updateNotify();
        }
    }

    /// Mark all active topics as dirty.
    pub fn markAllDirty(self: *RtdContext) void {
        for (&self.topics) |*t| {
            if (t.active) t.dirty = true;
        }
    }
};

// ============================================================================
// RtdServer — generic COM RTD server parameterized by Handler
// ============================================================================

pub const RtdConfig = struct {
    clsid: GUID,
    prog_id: [:0]const u8,
};

pub fn RtdServer(comptime Handler: type, comptime config: RtdConfig) type {
    return struct {
        const Self = @This();

        // ---- server state (non-extern, allocated on heap) ----
        const ServerState = struct {
            ref_count: ULONG = 1,
            ctx: RtdContext = .{},
            handler: Handler = .{},
        };

        // ---- COM object layout ----
        const ComObject = extern struct {
            vtable: *const IRtdServer_VTable = &vtable_instance,
            state: *anyopaque = undefined,

            fn getState(self: *ComObject) *ServerState {
                return @ptrCast(@alignCast(self.state));
            }
        };

        // ---- vtable types ----
        const IRtdServer_VTable = extern struct {
            QueryInterface: *const fn (*anyopaque, *const GUID, *?*anyopaque) callconv(.winapi) HRESULT,
            AddRef: *const fn (*anyopaque) callconv(.winapi) ULONG,
            Release: *const fn (*anyopaque) callconv(.winapi) ULONG,
            GetTypeInfoCount: *const fn (*anyopaque, *c_uint) callconv(.winapi) HRESULT,
            GetTypeInfo: *const fn (*anyopaque, c_uint, u32, *?*anyopaque) callconv(.winapi) HRESULT,
            GetIDsOfNames: *const fn (*anyopaque, *const GUID, [*]const [*:0]const u16, c_uint, u32, [*]c_long) callconv(.winapi) HRESULT,
            Invoke: *const fn (*anyopaque, c_long, *const GUID, u32, u16, *anyopaque, *anyopaque, *anyopaque, *anyopaque) callconv(.winapi) HRESULT,
            ServerStart: *const fn (*anyopaque, *IRTDUpdateEvent, *LONG) callconv(.winapi) HRESULT,
            ConnectData: *const fn (*anyopaque, LONG, **SAFEARRAY, *c_short, *VARIANT) callconv(.winapi) HRESULT,
            RefreshData: *const fn (*anyopaque, *LONG, *?*SAFEARRAY) callconv(.winapi) HRESULT,
            DisconnectData: *const fn (*anyopaque, LONG) callconv(.winapi) HRESULT,
            Heartbeat: *const fn (*anyopaque, *LONG) callconv(.winapi) HRESULT,
            ServerTerminate: *const fn (*anyopaque) callconv(.winapi) HRESULT,
        };

        const IClassFactory_VTable = extern struct {
            QueryInterface: *const fn (*anyopaque, *const GUID, *?*anyopaque) callconv(.winapi) HRESULT,
            AddRef: *const fn (*anyopaque) callconv(.winapi) ULONG,
            Release: *const fn (*anyopaque) callconv(.winapi) ULONG,
            CreateInstance: *const fn (*anyopaque, ?*anyopaque, *const GUID, *?*anyopaque) callconv(.winapi) HRESULT,
            LockServer: *const fn (*anyopaque, c_int) callconv(.winapi) HRESULT,
        };

        // ---- globals ----
        var g_object_count: i32 = 0;
        var g_class_factory: extern struct { vtable: *const IClassFactory_VTable } = .{ .vtable = &cf_vtable_instance };

        // ---- helpers ----
        fn getObj(self_opaque: *anyopaque) *ComObject {
            return @ptrCast(@alignCast(self_opaque));
        }

        // ---- IUnknown ----
        fn queryInterface(self_opaque: *anyopaque, riid: *const GUID, ppv: *?*anyopaque) callconv(.winapi) HRESULT {
            if (guidEql(riid, &IID_IUnknown) or guidEql(riid, &IID_IDispatch) or guidEql(riid, &IID_IRtdServer)) {
                ppv.* = self_opaque;
                _ = addRef(self_opaque);
                return S_OK;
            }
            ppv.* = null;
            return E_NOINTERFACE;
        }

        fn addRef(self_opaque: *anyopaque) callconv(.winapi) ULONG {
            const s = getObj(self_opaque).getState();
            s.ref_count += 1;
            return s.ref_count;
        }

        fn release(self_opaque: *anyopaque) callconv(.winapi) ULONG {
            const obj = getObj(self_opaque);
            const s = obj.getState();
            s.ref_count -= 1;
            const rc = s.ref_count;
            if (rc == 0) {
                s.handler.onTerminate(&s.ctx);
                if (s.ctx.update_event) |evt| {
                    _ = evt.vtable.Release(@ptrCast(evt));
                }
                std.heap.c_allocator.destroy(s);
                std.heap.c_allocator.destroy(obj);
                _ = @atomicRmw(i32, &g_object_count, .Sub, 1, .monotonic);
            }
            return rc;
        }

        // ---- IDispatch ----
        fn getTypeInfoCount(_: *anyopaque, pctinfo: *c_uint) callconv(.winapi) HRESULT {
            pctinfo.* = 0;
            return S_OK;
        }

        fn getTypeInfo(_: *anyopaque, _: c_uint, _: u32, _: *?*anyopaque) callconv(.winapi) HRESULT {
            return E_NOTIMPL;
        }

        fn getIDsOfNames(_: *anyopaque, _: *const GUID, rgszNames: [*]const [*:0]const u16, cNames: c_uint, _: u32, rgDispId: [*]c_long) callconv(.winapi) HRESULT {
            if (cNames < 1) return S_OK;
            const name = rgszNames[0];
            const names = [_]struct { n: [*:0]const u16, id: c_long }{
                .{ .n = std.unicode.utf8ToUtf16LeStringLiteral("ServerStart"), .id = 10 },
                .{ .n = std.unicode.utf8ToUtf16LeStringLiteral("ConnectData"), .id = 11 },
                .{ .n = std.unicode.utf8ToUtf16LeStringLiteral("RefreshData"), .id = 12 },
                .{ .n = std.unicode.utf8ToUtf16LeStringLiteral("DisconnectData"), .id = 13 },
                .{ .n = std.unicode.utf8ToUtf16LeStringLiteral("Heartbeat"), .id = 14 },
                .{ .n = std.unicode.utf8ToUtf16LeStringLiteral("ServerTerminate"), .id = 15 },
            };
            for (names) |entry| {
                if (wcsEql(name, entry.n)) {
                    rgDispId[0] = entry.id;
                    return S_OK;
                }
            }
            rgDispId[0] = -1;
            return DISP_E_UNKNOWNNAME;
        }

        fn invoke(self_opaque: *anyopaque, dispid: c_long, _: *const GUID, _: u32, _: u16, pDispParams: *anyopaque, pVarResult: *anyopaque, _: *anyopaque, _: *anyopaque) callconv(.winapi) HRESULT {
            const params: *DISPPARAMS = @ptrCast(@alignCast(pDispParams));

            switch (dispid) {
                10 => { // ServerStart
                    if (params.cArgs < 1) return E_FAIL;
                    const args = params.rgvarg orelse return E_FAIL;
                    const cb_idx: usize = if (params.cArgs >= 2) 1 else 0;
                    const callback_var = &args[cb_idx];
                    if (callback_var.vt != VT_DISPATCH and callback_var.vt != VT_UNKNOWN) return E_FAIL;
                    const callback_ptr: *IRTDUpdateEvent = @ptrCast(@alignCast(callback_var.data.ptr orelse return E_FAIL));

                    var pfRes: LONG = 0;
                    const hr = serverStart(self_opaque, callback_ptr, &pfRes);
                    const result: *VARIANT = @ptrCast(@alignCast(pVarResult));
                    VariantInit(result);
                    result.vt = VT_I4;
                    result.data = .{ .lval = pfRes };
                    return hr;
                },
                11 => { // ConnectData
                    const args = params.rgvarg orelse return E_FAIL;
                    if (params.cArgs < 3) return E_FAIL;
                    const topic_id: LONG = args[2].data.lval;

                    const s = getObj(self_opaque).getState();
                    for (&s.ctx.topics) |*t| {
                        if (!t.active) {
                            t.active = true;
                            t.topic_id = topic_id;
                            t.dirty = true;
                            s.ctx.topic_count += 1;
                            break;
                        }
                    }

                    s.handler.onConnect(&s.ctx, topic_id, s.ctx.topic_count);

                    const result: *VARIANT = @ptrCast(@alignCast(pVarResult));
                    rtdValueToVariant(s.handler.onRefreshValue(&s.ctx, topic_id), result);
                    return S_OK;
                },
                12 => { // RefreshData
                    var topic_count: LONG = 0;
                    var psa: ?*SAFEARRAY = null;
                    const hr = refreshData(self_opaque, &topic_count, &psa);

                    if (params.cArgs >= 1) {
                        const args = params.rgvarg orelse return hr;
                        if (args[0].vt == (VT_BYREF | VT_I4)) {
                            const plong: *LONG = @ptrCast(@alignCast(args[0].data.ptr orelse return hr));
                            plong.* = topic_count;
                        }
                    }

                    const result: *VARIANT = @ptrCast(@alignCast(pVarResult));
                    VariantInit(result);
                    if (psa) |sa| {
                        result.vt = VT_ARRAY | VT_VARIANT;
                        result.data = .{ .ptr = @ptrCast(sa) };
                    }
                    return hr;
                },
                13 => { // DisconnectData
                    if (params.cArgs < 1) return E_FAIL;
                    const args = params.rgvarg orelse return E_FAIL;
                    return disconnectData(self_opaque, args[0].data.lval);
                },
                14 => { // Heartbeat
                    var pfRes: LONG = 0;
                    const hr = heartbeat(self_opaque, &pfRes);
                    const result: *VARIANT = @ptrCast(@alignCast(pVarResult));
                    VariantInit(result);
                    result.vt = VT_I4;
                    result.data = .{ .lval = pfRes };
                    return hr;
                },
                15 => { // ServerTerminate
                    return serverTerminate(self_opaque);
                },
                else => return DISP_E_MEMBERNOTFOUND,
            }
        }

        // ---- IRtdServer methods ----
        fn serverStart(self_opaque: *anyopaque, callback: *IRTDUpdateEvent, pfRes: *LONG) callconv(.winapi) HRESULT {
            debugLog("ServerStart", .{});
            const s = getObj(self_opaque).getState();

            _ = callback.vtable.AddRef(@ptrCast(callback));
            s.ctx.update_event = callback;

            s.handler.onStart(&s.ctx);

            pfRes.* = 1;
            return S_OK;
        }

        fn refreshData(self_opaque: *anyopaque, topic_count: *LONG, parray_out: *?*SAFEARRAY) callconv(.winapi) HRESULT {
            const s = getObj(self_opaque).getState();

            var dirty_count: ULONG = 0;
            for (&s.ctx.topics) |*t| {
                if (t.active and t.dirty) dirty_count += 1;
            }

            topic_count.* = @intCast(dirty_count);
            if (dirty_count == 0) {
                parray_out.* = null;
                return S_OK;
            }

            var bounds = [2]SAFEARRAYBOUND{
                .{ .cElements = 2, .lLbound = 0 },
                .{ .cElements = dirty_count, .lLbound = 0 },
            };
            const psa = SafeArrayCreate(VT_VARIANT, 2, &bounds) orelse return E_FAIL;

            var row: LONG = 0;
            for (&s.ctx.topics) |*t| {
                if (t.active and t.dirty) {
                    t.dirty = false;

                    var id_var: VARIANT = undefined;
                    VariantInit(&id_var);
                    id_var.vt = VT_I4;
                    id_var.data = .{ .lval = t.topic_id };
                    var idx0 = [2]LONG{ 0, row };
                    _ = SafeArrayPutElement(psa, &idx0, @ptrCast(&id_var));

                    var val_var: VARIANT = undefined;
                    rtdValueToVariant(s.handler.onRefreshValue(&s.ctx, t.topic_id), &val_var);
                    var idx1 = [2]LONG{ 1, row };
                    _ = SafeArrayPutElement(psa, &idx1, @ptrCast(&val_var));

                    row += 1;
                }
            }

            parray_out.* = psa;
            return S_OK;
        }

        fn disconnectData(self_opaque: *anyopaque, topic_id: LONG) callconv(.winapi) HRESULT {
            debugLog("DisconnectData: TopicID={d}", .{topic_id});
            const s = getObj(self_opaque).getState();

            for (&s.ctx.topics) |*t| {
                if (t.active and t.topic_id == topic_id) {
                    t.active = false;
                    s.ctx.topic_count -= 1;
                    break;
                }
            }
            s.handler.onDisconnect(&s.ctx, topic_id, s.ctx.topic_count);
            return S_OK;
        }

        fn heartbeat(_: *anyopaque, pfRes: *LONG) callconv(.winapi) HRESULT {
            pfRes.* = 1;
            return S_OK;
        }

        fn serverTerminate(self_opaque: *anyopaque) callconv(.winapi) HRESULT {
            debugLog("ServerTerminate", .{});
            const s = getObj(self_opaque).getState();

            s.handler.onTerminate(&s.ctx);

            if (s.ctx.update_event) |evt| {
                _ = evt.vtable.Release(@ptrCast(evt));
                s.ctx.update_event = null;
            }
            return S_OK;
        }

        fn connectDataStub(_: *anyopaque, _: LONG, _: **SAFEARRAY, _: *c_short, _: *VARIANT) callconv(.winapi) HRESULT {
            return E_NOTIMPL;
        }

        // ---- vtable instances ----
        const vtable_instance = IRtdServer_VTable{
            .QueryInterface = &queryInterface,
            .AddRef = &addRef,
            .Release = &release,
            .GetTypeInfoCount = &getTypeInfoCount,
            .GetTypeInfo = &getTypeInfo,
            .GetIDsOfNames = &getIDsOfNames,
            .Invoke = &invoke,
            .ServerStart = &serverStart,
            .ConnectData = &connectDataStub,
            .RefreshData = &refreshData,
            .DisconnectData = &disconnectData,
            .Heartbeat = &heartbeat,
            .ServerTerminate = &serverTerminate,
        };

        // ---- IClassFactory ----
        fn cfQueryInterface(self: *anyopaque, riid: *const GUID, ppv: *?*anyopaque) callconv(.winapi) HRESULT {
            if (guidEql(riid, &IID_IUnknown) or guidEql(riid, &IID_IClassFactory)) {
                ppv.* = self;
                return S_OK;
            }
            ppv.* = null;
            return E_NOINTERFACE;
        }
        fn cfAddRef(_: *anyopaque) callconv(.winapi) ULONG {
            return 1;
        }
        fn cfRelease(_: *anyopaque) callconv(.winapi) ULONG {
            return 1;
        }

        fn cfCreateInstance(_: *anyopaque, pUnkOuter: ?*anyopaque, riid: *const GUID, ppv: *?*anyopaque) callconv(.winapi) HRESULT {
            if (pUnkOuter != null) {
                ppv.* = null;
                return CLASS_E_NOAGGREGATION;
            }

            const obj = std.heap.c_allocator.create(ComObject) catch return E_FAIL;
            const s = std.heap.c_allocator.create(ServerState) catch {
                std.heap.c_allocator.destroy(obj);
                return E_FAIL;
            };
            s.* = .{};
            obj.* = .{ .state = @ptrCast(s) };
            _ = @atomicRmw(i32, &g_object_count, .Add, 1, .monotonic);

            return queryInterface(@ptrCast(obj), riid, ppv);
        }

        fn cfLockServer(_: *anyopaque, fLock: c_int) callconv(.winapi) HRESULT {
            if (fLock != 0) {
                _ = @atomicRmw(i32, &g_object_count, .Add, 1, .monotonic);
            } else {
                _ = @atomicRmw(i32, &g_object_count, .Sub, 1, .monotonic);
            }
            return S_OK;
        }

        const cf_vtable_instance = IClassFactory_VTable{
            .QueryInterface = &cfQueryInterface,
            .AddRef = &cfAddRef,
            .Release = &cfRelease,
            .CreateInstance = &cfCreateInstance,
            .LockServer = &cfLockServer,
        };

        // ---- DLL exports ----
        fn dllGetClassObject(rclsid: *const GUID, riid: *const GUID, ppv: *?*anyopaque) callconv(.winapi) HRESULT {
            if (!guidEql(rclsid, &config.clsid)) {
                ppv.* = null;
                return CLASS_E_CLASSNOTAVAILABLE;
            }
            return cfQueryInterface(@ptrCast(&g_class_factory), riid, ppv);
        }

        fn dllCanUnloadNow() callconv(.winapi) HRESULT {
            return if (@atomicLoad(i32, &g_object_count, .acquire) == 0) S_OK else S_FALSE;
        }

        /// Try to handle a DllGetClassObject call for this server's CLSID.
        /// Returns S_OK if matched, CLASS_E_CLASSNOTAVAILABLE if not ours.
        pub fn tryGetClassObject(rclsid: *const GUID, riid: *const GUID, ppv: *?*anyopaque) HRESULT {
            return dllGetClassObject(rclsid, riid, ppv);
        }

        /// Returns true if this server has active objects.
        pub fn hasActiveObjects() bool {
            return @atomicLoad(i32, &g_object_count, .acquire) != 0;
        }

        /// Call this from a comptime block to emit the DLL exports.
        pub fn exportDllFunctions() void {
            @export(&dllGetClassObject, .{ .name = "DllGetClassObject" });
            @export(&dllCanUnloadNow, .{ .name = "DllCanUnloadNow" });
        }
    };
}

// ============================================================================
// Some helpers
// ============================================================================

fn guidEql(a: *const GUID, b: *const GUID) bool {
    const a_bytes: *const [16]u8 = @ptrCast(a);
    const b_bytes: *const [16]u8 = @ptrCast(b);
    return std.mem.eql(u8, a_bytes, b_bytes);
}

fn wcsEql(a: [*:0]const u16, b: [*:0]const u16) bool {
    var i: usize = 0;
    while (true) : (i += 1) {
        if (a[i] != b[i]) return false;
        if (a[i] == 0) return true;
    }
}
