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

// State pool management
//
// Each pool slot is an independent lua_State with its own globals/GC.
// Sync calls use slot 0 (the "main" state) under a mutex.
// Async workers acquire any free slot via CAS.

/// Pool size: configurable via build option `lua_states`, default 8.
const pool_size = blk: {
    const opt = @import("build_options").lua_states;
    break :blk if (opt > 0) opt else 8;
};

const StateSlot = struct {
    L: ?*lua_State = null,
    in_use: std.atomic.Value(bool) = std.atomic.Value(bool).init(false),
};

var state_pool: [pool_size]StateSlot = [_]StateSlot{.{}} ** pool_size;
var pool_initialized: bool = false;

/// Mutex for the main (sync) state — slot 0.
var main_state_mutex: std.Thread.Mutex = .{};

/// Cached script sources for loading into new states.
const ScriptEntry = struct {
    source: []const u8,
    name: [*:0]const u8,
};
var cached_scripts: std.ArrayListUnmanaged(ScriptEntry) = .empty;

/// Remove dangerous globals and module functions to sandbox user scripts.
/// Keeps safe functions like os.time, os.clock, os.date, os.difftime.
fn sandbox(L: *lua_State) void {
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

// ============================================================================
// Shared key-value store: xll.get(key) / xll.set(key, value)
//
// Mutex-protected store accessible from all pool states.
// Values are number, string, boolean, or nil.
// ============================================================================

const SharedValue = union(enum) {
    number: f64,
    string: []const u8, // heap-allocated copy
    boolean: bool,
};

var shared_store: std.StringHashMapUnmanaged(SharedValue) = .empty;
var shared_store_mutex: std.Thread.Mutex = .{};

fn sharedGet(key: []const u8) ?SharedValue {
    shared_store_mutex.lock();
    defer shared_store_mutex.unlock();
    return shared_store.get(key);
}

fn sharedSet(key: []const u8, value: ?SharedValue) void {
    shared_store_mutex.lock();
    defer shared_store_mutex.unlock();

    const alloc = std.heap.c_allocator;

    // Remove old value if present (free old string)
    if (shared_store.fetchRemove(key)) |kv| {
        if (kv.value == .string) alloc.free(kv.value.string);
        alloc.free(@constCast(kv.key));
    }

    // If setting nil, just remove
    const val = value orelse return;

    // Dupe key and string value
    const owned_key = alloc.dupe(u8, key) catch return;
    const owned_val: SharedValue = switch (val) {
        .string => |s| .{ .string = alloc.dupe(u8, s) catch {
            alloc.free(owned_key);
            return;
        } },
        else => val,
    };
    shared_store.put(alloc, owned_key, owned_val) catch {
        alloc.free(owned_key);
        if (owned_val == .string) alloc.free(owned_val.string);
    };
}

/// xll.get(key) — returns value or nil
fn luaXllGet(L: ?*c.lua_State) callconv(.c) c_int {
    const state = L orelse return 0;
    var len: usize = 0;
    const ptr = c.lua_tolstring(state, 1, &len) orelse {
        c.lua_pushnil(state);
        return 1;
    };
    const key = ptr[0..len];

    if (sharedGet(key)) |val| {
        switch (val) {
            .number => |n| c.lua_pushnumber(state, n),
            .boolean => |b| c.lua_pushboolean(state, if (b) 1 else 0),
            .string => |s| _ = c.lua_pushlstring(state, s.ptr, s.len),
        }
    } else {
        c.lua_pushnil(state);
    }
    return 1;
}

/// xll.set(key, value) — stores value (nil to delete)
fn luaXllSet(L: ?*c.lua_State) callconv(.c) c_int {
    const state = L orelse return 0;
    var key_len: usize = 0;
    const key_ptr = c.lua_tolstring(state, 1, &key_len) orelse return 0;
    const key = key_ptr[0..key_len];

    const val_type = c.lua_type(state, 2);
    const value: ?SharedValue = switch (val_type) {
        c.LUA_TNUMBER => .{ .number = c.lua_tonumberx(state, 2, null) },
        c.LUA_TBOOLEAN => .{ .boolean = c.lua_toboolean(state, 2) != 0 },
        c.LUA_TSTRING => blk: {
            var len: usize = 0;
            const p = c.lua_tolstring(state, 2, &len) orelse break :blk null;
            break :blk .{ .string = p[0..len] };
        },
        else => null, // nil or unsupported → delete
    };

    sharedSet(key, value);
    return 0;
}

/// Register the xll table (get/set) into a Lua state.
fn registerXllLib(L: *lua_State) void {
    // Create xll = {}
    c.lua_createtable(@ptrCast(L), 0, 2);

    // xll.get = luaXllGet
    c.lua_pushcclosure(@ptrCast(L), luaXllGet, 0);
    c.lua_setfield(@ptrCast(L), -2, "get");

    // xll.set = luaXllSet
    c.lua_pushcclosure(@ptrCast(L), luaXllSet, 0);
    c.lua_setfield(@ptrCast(L), -2, "set");

    c.lua_setglobal(@ptrCast(L), "xll");
}

/// Create and initialize a single Lua state with libs + sandbox + xll store + xllify builtins.
fn createState() !*lua_State {
    const L = c.luaL_newstate() orelse return error.LuaInitFailed;
    c.luaL_openlibs(L);
    sandbox(L);
    registerXllLib(L);
    @import("lua_builtins.zig").register(L);
    return L;
}

pub fn getPoolSize() usize {
    return pool_size;
}

/// Initialize the state pool. Creates `pool_size` independent Lua states.
pub fn init() !void {
    if (pool_initialized) return;

    for (&state_pool) |*slot| {
        slot.L = createState() catch return error.LuaInitFailed;
    }
    pool_initialized = true;
}

/// Load and execute a Lua source string on all pool states.
/// Also caches the script so future states could be re-initialized.
pub fn loadScript(source: []const u8, chunk_name: [*:0]const u8) !void {
    if (!pool_initialized) return error.NotInitialized;

    // Cache for reference
    cached_scripts.append(std.heap.c_allocator, .{ .source = source, .name = chunk_name }) catch {};

    for (&state_pool) |*slot| {
        const L = slot.L orelse continue;
        if (c.luaL_loadbufferx(L, source.ptr, source.len, chunk_name, null) != LUA_OK) {
            logLuaError(L);
            lua_pop(L, 1);
            return error.LoadFailed;
        }
        if (lua_pcall(L, 0, 0, 0) != LUA_OK) {
            logLuaError(L);
            lua_pop(L, 1);
            return error.ExecFailed;
        }
    }
}

/// Get the main (sync) state — slot 0. Caller must hold main_state_mutex.
pub fn getState() ?*lua_State {
    return state_pool[0].L;
}

/// Lock the main state for sync use.
pub fn lockMain() void {
    main_state_mutex.lock();
}

/// Unlock the main state after sync use.
pub fn unlockMain() void {
    main_state_mutex.unlock();
}

/// Acquire any free state from the pool for async use.
/// Spins briefly, returns null if no state available.
pub fn acquireState() ?*lua_State {
    // Try a few rounds of CAS across all slots
    var attempts: u32 = 0;
    while (attempts < 1000) : (attempts += 1) {
        for (&state_pool) |*slot| {
            if (slot.in_use.cmpxchgWeak(false, true, .acquire, .monotonic) == null) {
                return slot.L;
            }
        }
        // Brief spin — on Windows use Sleep(0), otherwise yield
        if (@import("builtin").os.tag == .windows) {
            const kernel32 = struct {
                extern "kernel32" fn Sleep(dwMilliseconds: u32) callconv(.winapi) void;
            };
            kernel32.Sleep(0);
        } else {
            std.Thread.yield() catch {};
        }
    }
    return null;
}

/// Release a state back to the pool.
pub fn releaseState(L: *lua_State) void {
    for (&state_pool) |*slot| {
        if (slot.L == L) {
            slot.in_use.store(false, .release);
            return;
        }
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

/// Shut down all pool states and free the shared store.
pub fn deinit() void {
    if (!pool_initialized) return;

    for (&state_pool) |*slot| {
        if (slot.L) |L| {
            c.lua_close(L);
            slot.L = null;
        }
        slot.in_use = std.atomic.Value(bool).init(false);
    }
    cached_scripts = .empty;

    // Free shared store
    const alloc = std.heap.c_allocator;
    var it = shared_store.iterator();
    while (it.next()) |entry| {
        if (entry.value_ptr.* == .string) alloc.free(entry.value_ptr.string);
        alloc.free(@constCast(entry.key_ptr.*));
    }
    shared_store.deinit(alloc);
    shared_store = .empty;

    pool_initialized = false;
}
