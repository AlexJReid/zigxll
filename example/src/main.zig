pub const function_modules = .{
    @import("my_functions.zig"),
    @import("async_functions.zig"),
    @import("my_macros.zig"),
    @import("lua_functions.zig"),
};

pub const rtd_servers = .{
    @import("timer_rtd.zig"),
};

pub const lua_scripts = .{
    .{ .name = "functions", .source = @embedFile("lua/functions.lua") },
};
