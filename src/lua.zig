// Lua 5.4 C API bindings and state management
const std = @import("std");
const c = @cImport({
    @cInclude("lua.h");
    @cInclude("lauxlib.h");
    @cInclude("lualib.h");
});

// Re-export types and constants
pub const lua_State = c.lua_State;
pub const LUA_OK = c.LUA_OK;
pub const LUA_TNIL = c.LUA_TNIL;
pub const LUA_TBOOLEAN = c.LUA_TBOOLEAN;
pub const LUA_TNUMBER = c.LUA_TNUMBER;
pub const LUA_TSTRING = c.LUA_TSTRING;
pub const LUA_TFUNCTION = c.LUA_TFUNCTION;
pub const LUA_TTABLE = c.LUA_TTABLE;
pub const LUA_REGISTRYINDEX = c.LUA_REGISTRYINDEX;

// Re-export functions
pub const lua_pushnumber = c.lua_pushnumber;
pub const lua_pushboolean = c.lua_pushboolean;
pub const lua_pushnil = c.lua_pushnil;
pub const lua_toboolean = c.lua_toboolean;
pub const lua_type = c.lua_type;
pub const lua_settop = c.lua_settop;
pub const lua_gettop = c.lua_gettop;

pub fn lua_pushlstring(L: *lua_State, s: [*]const u8, len: usize) void {
    _ = c.lua_pushlstring(L, s, len);
}

pub fn lua_pushstring(L: *lua_State, s: [*:0]const u8) void {
    _ = c.lua_pushstring(L, s);
}

pub fn lua_pcall(L: *lua_State, nargs: c_int, nresults: c_int, msgh: c_int) c_int {
    return c.lua_pcallk(L, nargs, nresults, msgh, 0, null);
}

pub fn lua_tonumber(L: *lua_State, idx: c_int) f64 {
    return c.lua_tonumberx(L, idx, null);
}

pub fn lua_tolstring(L: *lua_State, idx: c_int, len: *usize) ?[*]const u8 {
    return c.lua_tolstring(L, idx, len);
}

pub fn lua_pop(L: *lua_State, n: c_int) void {
    c.lua_settop(L, -(n) - 1);
}

pub fn lua_getglobal(L: *lua_State, name: [*:0]const u8) c_int {
    return c.lua_getglobal(L, name);
}

pub fn lua_getfield(L: *lua_State, idx: c_int, name: [*:0]const u8) c_int {
    return c.lua_getfield(L, idx, name);
}

pub fn lua_setglobal(L: *lua_State, name: [*:0]const u8) void {
    c.lua_setglobal(L, name);
}

// State management
var global_state: ?*lua_State = null;
var state_mutex: std.Thread.Mutex = .{};

/// Remove dangerous globals and module functions to sandbox user scripts.
/// Keeps safe functions like os.time, os.clock, os.date, os.difftime.
fn sandbox(L: *lua_State) void {
    logInfo("Lua: applying sandbox");

    // Remove globals that can load/execute arbitrary code or access the filesystem
    const removed_globals = [_][*:0]const u8{
        "dofile",
        "loadfile",
        "load",
        "require",
    };
    for (removed_globals) |name| {
        lua_pushnil(L);
        lua_setglobal(L, name);
    }

    // Remove the io library entirely
    lua_pushnil(L);
    lua_setglobal(L, "io");

    // Remove dangerous functions from os, keep the safe ones
    const removed_os_fns = [_][*:0]const u8{
        "execute",
        "remove",
        "rename",
        "tmpname",
        "getenv",
        "exit",
    };
    _ = lua_getglobal(L, "os");
    if (lua_type(L, -1) == LUA_TTABLE) {
        for (removed_os_fns) |name| {
            lua_pushnil(L);
            c.lua_setfield(L, -2, name);
        }
    }
    lua_pop(L, 1);
}

/// Initialize the global Lua state and load standard libraries
pub fn init() !void {
    state_mutex.lock();
    defer state_mutex.unlock();

    if (global_state != null) return;

    const L = c.luaL_newstate() orelse return error.LuaInitFailed;
    c.luaL_openlibs(L);
    sandbox(L);
    global_state = L;
}

/// Load and execute a Lua source string
pub fn loadScript(source: []const u8, chunk_name: [*:0]const u8) !void {
    state_mutex.lock();
    defer state_mutex.unlock();

    const L = global_state orelse return error.NotInitialized;

    // Load source
    if (c.luaL_loadbufferx(L, source.ptr, source.len, chunk_name, null) != LUA_OK) {
        logLuaError(L);
        lua_pop(L, 1);
        return error.LoadFailed;
    }

    // Execute the chunk to define globals
    if (lua_pcall(L, 0, 0, 0) != LUA_OK) {
        logLuaError(L);
        lua_pop(L, 1);
        return error.ExecFailed;
    }
}

fn logInfo(comptime msg: []const u8) void {
    if (@import("builtin").os.tag == .windows) {
        @import("xl_helpers.zig").debugLog(msg);
    } else {
        std.debug.print(msg ++ "\n", .{});
    }
}

/// Pull error string from top of Lua stack and log it
fn logLuaError(L: *lua_State) void {
    var len: usize = 0;
    const ptr = c.lua_tolstring(L, -1, &len);
    if (ptr) |p| {
        const msg = p[0..len];
        if (@import("builtin").os.tag == .windows) {
            @import("xl_helpers.zig").debugLogRuntime(msg);
        } else {
            std.debug.print("Lua error: {s}\n", .{msg});
        }
    }
}

/// Get the global Lua state
pub fn getState() ?*lua_State {
    return global_state;
}

/// Shut down the global Lua state
pub fn deinit() void {
    state_mutex.lock();
    defer state_mutex.unlock();

    if (global_state) |L| {
        c.lua_close(L);
        global_state = null;
    }
}
