pub const function_modules = .{
    @import("my_functions.zig"),
    @import("async_functions.zig"),
    @import("my_macros.zig"),
};

pub const rtd_servers = .{
    @import("timer_rtd.zig"),
};
