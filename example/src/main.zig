pub const function_modules = .{
    @import("my_functions.zig"),
};

pub const rtd_servers = .{
    @import("timer_rtd.zig"),
};

// RTD server - force comptime evaluation so COM exports are emitted
comptime {
    _ = @import("timer_rtd.zig");
}
