// Test runner for all modules
// This file imports modules to run their tests

const test_options = @import("test_options");

comptime {
    _ = @import("xlvalue.zig");
    _ = @import("rtd_tests.zig");
    if (test_options.enable_lua) {
        _ = @import("lua_tests.zig");
    }
}
