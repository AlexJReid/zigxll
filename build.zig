const std = @import("std");

// Framework build
pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{
        .default_target = .{
            .cpu_arch = .x86_64,
            .os_tag = .windows,
            .abi = .msvc,
        },
    });
    const optimize = b.standardOptimizeOption(.{ .preferred_optimize_mode = .ReleaseSmall });

    // Create module for xll framework (for external users)
    const xll_module = b.addModule("xll", .{
        .root_source_file = b.path("src/root.zig"),
        .target = target,
        .optimize = optimize,
    });
    xll_module.addIncludePath(b.path("excel/include"));

    // Build test XLL for framework development - simple direct build
    const xll = b.addLibrary(.{
        .name = "zigxll",
        .root_module = b.createModule(.{
            .root_source_file = b.path("src/framework_main.zig"),
            .target = target,
            .optimize = optimize,
        }),
        .linkage = .dynamic,
        .version = .{ .major = 1, .minor = 0, .patch = 0 },
    });

    xll.addIncludePath(b.path("excel/include"));
    xll.addLibraryPath(b.path("excel/lib"));

    xll.linkLibC();
    xll.linkSystemLibrary("user32");
    xll.linkSystemLibrary("xlcall32");
    xll.linkSystemLibrary("frmwrk32");

    const install_xll = b.addInstallFile(xll.getEmittedBin(), "lib/output.xll");
    b.getInstallStep().dependOn(&install_xll.step);
}

/// Helper function for users to create their own XLL
/// User provides a module that has a `function_modules` tuple
pub fn buildXll(
    b: *std.Build,
    options: struct {
        name: []const u8,
        user_module: *std.Build.Module,
        target: std.Build.ResolvedTarget,
        optimize: std.builtin.OptimizeMode,
    },
) *std.Build.Step.Compile {
    const target = options.target;
    const optimize = options.optimize;

    // Get xll dependency
    const xll_dep = b.dependency("xll", .{});
    const xll_framework = xll_dep.module("xll");

    // Give user module access to xll types (for ExcelFunction, etc.)
    options.user_module.addImport("xll", xll_framework);

    // Create the XLL using framework's entry point
    const xll = b.addLibrary(.{
        .name = options.name,
        .root_module = b.createModule(.{
            .root_source_file = xll_dep.path("src/xll_builder.zig"),
            .target = target,
            .optimize = optimize,
            .strip = true,
        }),
        .linkage = .dynamic,
        .version = .{ .major = 1, .minor = 0, .patch = 0 },
    });

    // Add framework module (xll_builder uses this as "xll_framework")
    xll.root_module.addImport("xll_framework", xll_framework);
    // Add LOCAL user module (contains function_modules tuple)
    xll.root_module.addImport("user_module", options.user_module);

    // Add Excel SDK from xll dependency
    const excel_include = xll_dep.path("excel/include");
    const excel_lib = xll_dep.path("excel/lib");

    xll.addIncludePath(excel_include);
    xll.addLibraryPath(excel_lib);
    xll.root_module.addIncludePath(excel_include);

    xll.linkLibC();
    xll.linkSystemLibrary("user32");
    xll.linkSystemLibrary("xlcall32");
    xll.linkSystemLibrary("frmwrk32");

    return xll;
}
