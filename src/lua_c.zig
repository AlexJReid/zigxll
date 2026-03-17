// Shared Lua 5.4 C API import — single @cImport to avoid duplicate opaque types.
// Both lua.zig and lua_builtins.zig import this file instead of doing their own @cImport,
// ensuring they share the same opaque type for lua_State.
pub const c = @cImport({
    @cInclude("lua.h");
    @cInclude("lauxlib.h");
    @cInclude("lualib.h");
});
