// RTD server auto-registration — writes COM registry entries to HKCU on xlAutoOpen.
// No admin required, idempotent, errors are logged but don't fail xlAutoOpen.

const std = @import("std");
const rtd = @import("rtd.zig");
const xl_helpers = @import("xl_helpers.zig");
const xl_imports = @import("xl_imports.zig");
const xl = xl_imports.xl;

// ============================================================================
// Windows Registry API (advapi32, already linked)
// ============================================================================

const HKEY = *anyopaque;
const LSTATUS = c_long;

const HKEY_CURRENT_USER: HKEY = @ptrFromInt(0x80000001);
const KEY_WRITE: u32 = 0x20006;
const REG_SZ: u32 = 1;
const REG_OPTION_NON_VOLATILE: u32 = 0;

extern "advapi32" fn RegCreateKeyExW(
    hKey: HKEY,
    lpSubKey: [*:0]const u16,
    Reserved: u32,
    lpClass: ?[*:0]const u16,
    dwOptions: u32,
    samDesired: u32,
    lpSecurityAttributes: ?*anyopaque,
    phkResult: *HKEY,
    lpdwDisposition: ?*u32,
) callconv(.winapi) LSTATUS;

extern "advapi32" fn RegSetValueExW(
    hKey: HKEY,
    lpValueName: ?[*:0]const u16,
    Reserved: u32,
    dwType: u32,
    lpData: [*]const u8,
    cbData: u32,
) callconv(.winapi) LSTATUS;

extern "advapi32" fn RegCloseKey(hKey: HKEY) callconv(.winapi) LSTATUS;

// ============================================================================
// Comptime GUID → string
// ============================================================================

fn hexChar(val: u4) u8 {
    const v: u8 = val;
    return if (v < 10) '0' + v else 'A' - 10 + v;
}

fn hexByte(b: u8) [2]u8 {
    return .{ hexChar(@truncate(b >> 4)), hexChar(@truncate(b & 0xF)) };
}

/// Formats a GUID as "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}" at comptime.
pub fn guidToString(comptime g: rtd.GUID) [38]u8 {
    comptime {
        var buf: [38]u8 = undefined;
        buf[0] = '{';

        // Data1: 4 bytes
        const d1 = std.mem.toBytes(std.mem.nativeToBig(u32, g.Data1));
        var pos: usize = 1;
        for (d1) |b| {
            const h = hexByte(b);
            buf[pos] = h[0];
            buf[pos + 1] = h[1];
            pos += 2;
        }
        buf[pos] = '-';
        pos += 1;

        // Data2: 2 bytes
        const d2 = std.mem.toBytes(std.mem.nativeToBig(u16, g.Data2));
        for (d2) |b| {
            const h = hexByte(b);
            buf[pos] = h[0];
            buf[pos + 1] = h[1];
            pos += 2;
        }
        buf[pos] = '-';
        pos += 1;

        // Data3: 2 bytes
        const d3 = std.mem.toBytes(std.mem.nativeToBig(u16, g.Data3));
        for (d3) |b| {
            const h = hexByte(b);
            buf[pos] = h[0];
            buf[pos + 1] = h[1];
            pos += 2;
        }
        buf[pos] = '-';
        pos += 1;

        // Data4[0..2]
        for (g.Data4[0..2]) |b| {
            const h = hexByte(b);
            buf[pos] = h[0];
            buf[pos + 1] = h[1];
            pos += 2;
        }
        buf[pos] = '-';
        pos += 1;

        // Data4[2..8]
        for (g.Data4[2..8]) |b| {
            const h = hexByte(b);
            buf[pos] = h[0];
            buf[pos + 1] = h[1];
            pos += 2;
        }

        buf[37] = '}';
        return buf;
    }
}

// ============================================================================
// XLL path extraction from XLOPER12
// ============================================================================

pub fn getXllPathSlice(xDLL: *const xl.XLOPER12) ?[]const u16 {
    if ((xDLL.xltype & 0xFFF) != xl.xltypeStr) return null;
    const str_ptr: [*]const u16 = @ptrCast(xDLL.val.str orelse return null);
    const len: usize = @intCast(str_ptr[0]);
    if (len == 0) return null;
    return str_ptr[1 .. len + 1];
}

// ============================================================================
// Registry helpers
// ============================================================================

fn setRegString(hkey: HKEY, name: ?[*:0]const u16, value: [*:0]const u16) void {
    // Calculate byte length including null terminator
    var len: usize = 0;
    while (value[len] != 0) : (len += 1) {}
    const cb: u32 = @intCast((len + 1) * 2);
    _ = RegSetValueExW(hkey, name, 0, REG_SZ, @ptrCast(value), cb);
}

fn setRegStringSlice(hkey: HKEY, name: ?[*:0]const u16, value: []const u16, buf: *[512]u16) void {
    const copy_len = @min(value.len, 511);
    @memcpy(buf[0..copy_len], value[0..copy_len]);
    buf[copy_len] = 0;
    const cb: u32 = @intCast((copy_len + 1) * 2);
    _ = RegSetValueExW(hkey, name, 0, REG_SZ, @ptrCast(buf), cb);
}

fn openKey(subkey: [*:0]const u16) ?HKEY {
    var hkey: HKEY = undefined;
    const status = RegCreateKeyExW(
        HKEY_CURRENT_USER,
        subkey,
        0,
        null,
        REG_OPTION_NON_VOLATILE,
        KEY_WRITE,
        null,
        &hkey,
        null,
    );
    return if (status == 0) hkey else null;
}

// ============================================================================
// Public registration function
// ============================================================================

pub fn registerRtdServer(
    comptime clsid: rtd.GUID,
    comptime prog_id: [:0]const u8,
    xll_path: []const u16,
) void {
    @setEvalBranchQuota(10000);
    const clsid_str = comptime guidToString(clsid);

    // Build all registry subkeys at comptime
    const prog_id_key = comptime std.unicode.utf8ToUtf16LeStringLiteral("Software\\Classes\\" ++ prog_id);
    const prog_id_clsid_key = comptime std.unicode.utf8ToUtf16LeStringLiteral("Software\\Classes\\" ++ prog_id ++ "\\CLSID");
    const clsid_key = comptime std.unicode.utf8ToUtf16LeStringLiteral("Software\\Classes\\CLSID\\" ++ clsid_str);
    const inproc_key = comptime std.unicode.utf8ToUtf16LeStringLiteral("Software\\Classes\\CLSID\\" ++ clsid_str ++ "\\InprocServer32");
    const progid_back_key = comptime std.unicode.utf8ToUtf16LeStringLiteral("Software\\Classes\\CLSID\\" ++ clsid_str ++ "\\ProgID");

    const wide_clsid = comptime std.unicode.utf8ToUtf16LeStringLiteral(&clsid_str);
    const wide_prog_id = comptime std.unicode.utf8ToUtf16LeStringLiteral(prog_id);
    const wide_apartment = comptime std.unicode.utf8ToUtf16LeStringLiteral("Apartment");
    const wide_threading = comptime std.unicode.utf8ToUtf16LeStringLiteral("ThreadingModel");

    var path_buf: [512]u16 = undefined;

    // 1. Software\Classes\{prog_id} -> Default = prog_id
    if (openKey(prog_id_key)) |hkey| {
        setRegString(hkey, null, wide_prog_id);
        _ = RegCloseKey(hkey);
    }

    // 2. Software\Classes\{prog_id}\CLSID -> Default = {CLSID}
    if (openKey(prog_id_clsid_key)) |hkey| {
        setRegString(hkey, null, wide_clsid);
        _ = RegCloseKey(hkey);
    }

    // 3. Software\Classes\CLSID\{CLSID} -> Default = prog_id
    if (openKey(clsid_key)) |hkey| {
        setRegString(hkey, null, wide_prog_id);
        _ = RegCloseKey(hkey);
    }

    // 4+5. Software\Classes\CLSID\{CLSID}\InprocServer32 -> Default = xll_path, ThreadingModel = Apartment
    if (openKey(inproc_key)) |hkey| {
        setRegStringSlice(hkey, null, xll_path, &path_buf);
        setRegString(hkey, wide_threading, wide_apartment);
        _ = RegCloseKey(hkey);
    }

    // 6. Software\Classes\CLSID\{CLSID}\ProgID -> Default = prog_id
    if (openKey(progid_back_key)) |hkey| {
        setRegString(hkey, null, wide_prog_id);
        _ = RegCloseKey(hkey);
    }

    xl_helpers.debugLogFmt("RTD server registered: {s} -> {s}", .{ prog_id, clsid_str });
}
