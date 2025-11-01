const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.resolveTargetQuery(.{
        .cpu_arch = .x86_64,
        .os_tag = .windows,
        .abi = .msvc,
    });
    const optimize = b.standardOptimizeOption(.{ .preferred_optimize_mode = .ReleaseSmall });

    // Create a module for our user functions (don't add imports yet)
    const user_module = b.createModule(.{
        .root_source_file = b.path("src/main.zig"),
        .target = target,
        .optimize = optimize,
    });

    // Build the XLL using the framework's helper
    // The helper will wire up the xll import for the user module
    const xll_build = @import("xll");
    const xll = xll_build.buildXll(b, .{
        .name = "my_excel_functions",
        .user_module = user_module,
        .target = target,
        .optimize = optimize,
    });

    // Install the XLL (rename .dll to .xll)
    const install_xll = b.addInstallFile(xll.getEmittedBin(), "lib/my_excel_functions.xll");
    b.getInstallStep().dependOn(&install_xll.step);
}
