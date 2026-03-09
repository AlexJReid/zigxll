const std = @import("std");
const builtin = @import("builtin");

// Framework build
pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{
        .default_target = if (builtin.os.tag == .windows)
            .{ .cpu_arch = .x86_64, .os_tag = .windows }
        else
            .{ .cpu_arch = .x86_64, .os_tag = .windows, .abi = .msvc },
    });
    const optimize = b.standardOptimizeOption(.{ .preferred_optimize_mode = .ReleaseSmall });

    // Add include path for ZLS to find C headers during @cImport analysis
    b.addSearchPrefix("excel");

    // Create module for xll framework (for external users)
    const xll_module = b.addModule("xll", .{
        .root_source_file = b.path("src/root.zig"),
        .target = target,
        .optimize = optimize,
    });
    xll_module.addIncludePath(b.path("excel/include"));

    // Build test XLL for framework development - simple direct build
    const framework_build_options = b.addOptions();
    framework_build_options.addOption([]const u8, "xll_name", "zigxll (framework test build)");
    framework_build_options.addOption([]const u8, "framework_version", @import("build.zig.zon").version);

    const xll = b.addLibrary(.{
        .name = "zigxll",
        .root_module = b.createModule(.{
            .root_source_file = b.path("src/framework_main.zig"),
            .target = target,
            .optimize = optimize,
            .strip = true,
        }),
        .linkage = .dynamic,
    });

    xll.root_module.addImport("build_options", framework_build_options.createModule());

    xll.addIncludePath(b.path("excel/include"));
    xll.addLibraryPath(b.path("excel/lib"));

    if (builtin.os.tag == .windows) {
        addNativeMsvcPaths(b, xll);
    } else {
        addXwinPaths(b, xll);
    }

    xll.linkLibC();
    xll.linkSystemLibrary("user32");
    xll.linkSystemLibrary("xlcall32");
    xll.linkSystemLibrary("frmwrk32");

    const install_xll = b.addInstallFile(xll.getEmittedBin(), "lib/output.xll");
    b.getInstallStep().dependOn(&install_xll.step);

    // Add test step - uses native target so tests can run on Mac/Linux
    const native_target = b.resolveTargetQuery(.{});
    const tests = b.addTest(.{
        .root_module = b.createModule(.{
            .root_source_file = b.path("src/tests.zig"),
            .target = native_target,
            .optimize = optimize,
        }),
    });
    tests.addIncludePath(b.path("excel/include"));
    tests.linkLibC();

    // Only link Excel libraries on Windows
    if (native_target.result.os.tag == .windows) {
        tests.addLibraryPath(b.path("excel/lib"));
        tests.linkSystemLibrary("xlcall32");
        tests.linkSystemLibrary("frmwrk32");
    }

    const run_tests = b.addRunArtifact(tests);
    run_tests.has_side_effects = true;

    const test_step = b.step("test", "Run unit tests");
    test_step.dependOn(&run_tests.step);
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

    // Create build options to pass XLL name and version to framework
    const build_options = b.addOptions();
    build_options.addOption([]const u8, "xll_name", options.name);
    build_options.addOption([]const u8, "framework_version", @import("build.zig.zon").version);

    // Get xll dependency
    const xll_dep = b.dependency("xll", .{});
    const xll_framework = xll_dep.module("xll");

    // Add build_options to framework module so it can log XLL name
    xll_framework.addImport("build_options", build_options.createModule());

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
    });

    // Add framework module (xll_builder uses this as "xll_framework")
    xll.root_module.addImport("xll_framework", xll_framework);
    // Add LOCAL user module (contains function_modules tuple)
    xll.root_module.addImport("user_module", options.user_module);
    // Add build options so framework can log XLL name
    xll.root_module.addImport("build_options", build_options.createModule());

    // Add Excel SDK from xll dependency
    const excel_include = xll_dep.path("excel/include");
    const excel_lib = xll_dep.path("excel/lib");

    xll.addIncludePath(excel_include);
    xll.addLibraryPath(excel_lib);
    xll.root_module.addIncludePath(excel_include);

    // Add Windows SDK/CRT paths to both the XLL compile step and the user module,
    // so that any C code the user compiles (e.g. nats.c) can find vcrt/ucrt headers and libs.
    if (builtin.os.tag == .windows) {
        addNativeMsvcPaths(b, xll);
    } else {
        addXwinPaths(b, xll);
    }
    applyXwinToModule(b, options.user_module);

    xll.linkLibC();
    xll.linkSystemLibrary("user32");
    xll.linkSystemLibrary("xlcall32");
    xll.linkSystemLibrary("frmwrk32");

    // COM/RTD support
    xll.linkSystemLibrary("oleaut32");
    xll.linkSystemLibrary("advapi32");
    xll.linkSystemLibrary("ole32");

    return xll;
}

fn requireEnv(b: *std.Build, name: []const u8) []const u8 {
    return std.process.getEnvVarOwned(b.allocator, name) catch {
        std.log.err("Missing environment variable '{s}'. Run from a Visual Studio Developer Command Prompt (vcvarsall.bat).", .{name});
        @panic("MSVC environment not configured");
    };
}

/// On native Windows, use VCToolsInstallDir / WindowsSdkDir env vars to locate the MSVC CRT.
/// These are set by the Visual Studio Developer Command Prompt (vcvarsall.bat).
fn addNativeMsvcPaths(b: *std.Build, compile: *std.Build.Step.Compile) void {
    const vctools = requireEnv(b, "VCToolsInstallDir");
    const ucrt_sdk = requireEnv(b, "UniversalCRTSdkDir");
    const ucrt_ver = requireEnv(b, "UCRTVersion");
    const win_sdk = requireEnv(b, "WindowsSdkDir");
    const win_sdk_ver = requireEnv(b, "WindowsSDKVersion");

    const msvc_lib_dir = b.fmt("{s}lib\\x64", .{vctools});
    const ucrt_lib_dir = b.fmt("{s}Lib\\{s}\\ucrt\\x64", .{ ucrt_sdk, ucrt_ver });
    const ucrt_inc_dir = b.fmt("{s}Include\\{s}\\ucrt", .{ ucrt_sdk, ucrt_ver });
    const vctools_inc = b.fmt("{s}include", .{vctools});
    const kernel32_lib_dir = b.fmt("{s}Lib\\{s}\\um\\x64", .{ win_sdk, win_sdk_ver });
    const um_inc_dir = b.fmt("{s}Include\\{s}\\um", .{ win_sdk, win_sdk_ver });
    const shared_inc_dir = b.fmt("{s}Include\\{s}\\shared", .{ win_sdk, win_sdk_ver });

    compile.addLibraryPath(.{ .cwd_relative = msvc_lib_dir });
    compile.addLibraryPath(.{ .cwd_relative = ucrt_lib_dir });
    compile.addLibraryPath(.{ .cwd_relative = kernel32_lib_dir });

    compile.addSystemIncludePath(.{ .cwd_relative = vctools_inc });
    compile.addSystemIncludePath(.{ .cwd_relative = ucrt_inc_dir });
    compile.addSystemIncludePath(.{ .cwd_relative = um_inc_dir });
    compile.addSystemIncludePath(.{ .cwd_relative = shared_inc_dir });

    const libc_conf = b.fmt(
        "include_dir={s}\n" ++
        "sys_include_dir={s}\n" ++
        "crt_dir={s}\n" ++
        "msvc_lib_dir={s}\n" ++
        "kernel32_lib_dir={s}\n" ++
        "gcc_dir=\n"
    , .{ ucrt_inc_dir, vctools_inc, ucrt_lib_dir, msvc_lib_dir, kernel32_lib_dir });

    const libc_file = b.addWriteFiles().add("libc.conf", libc_conf);
    compile.setLibCFile(libc_file);
}

/// When cross-compiling from Mac/Linux, add xwin system include paths to a module
/// so C code compiled within it can find headers like <vcruntime.h>, <corecrt.h>, <windows.h>.
/// On native Windows, Zig finds these automatically.
fn applyXwinToModule(b: *std.Build, mod: *std.Build.Module) void {
    if (builtin.os.tag == .windows) return;

    const home = std.process.getEnvVarOwned(b.allocator, "HOME") catch return;
    const xwin_dir = std.fs.path.join(b.allocator, &.{ home, ".xwin" }) catch return;
    var dir = std.fs.openDirAbsolute(xwin_dir, .{}) catch return;
    dir.close();

    mod.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/crt/include", .{xwin_dir}) });
    mod.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/sdk/include/ucrt", .{xwin_dir}) });
    mod.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/sdk/include/um", .{xwin_dir}) });
    mod.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/sdk/include/shared", .{xwin_dir}) });
}

/// If ~/.xwin exists (installed via `brew install xwin && xwin --accept-license splat --output ~/.xwin`),
/// add its Windows SDK and CRT paths so we can cross-compile to Windows from Mac/Linux.
fn addXwinPaths(b: *std.Build, compile: *std.Build.Step.Compile) void {
    // xwin is only used for cross-compiling to Windows from Mac/Linux
    const home = std.process.getEnvVarOwned(b.allocator, "HOME") catch return;
    const xwin_dir = std.fs.path.join(b.allocator, &.{ home, ".xwin" }) catch return;
    var dir = std.fs.openDirAbsolute(xwin_dir, .{}) catch return;
    dir.close();

    compile.addLibraryPath(.{ .cwd_relative = b.fmt("{s}/sdk/lib/um/x86_64", .{xwin_dir}) });
    compile.addLibraryPath(.{ .cwd_relative = b.fmt("{s}/crt/lib/x86_64", .{xwin_dir}) });
    compile.addLibraryPath(.{ .cwd_relative = b.fmt("{s}/sdk/lib/ucrt/x86_64", .{xwin_dir}) });

    compile.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/crt/include", .{xwin_dir}) });
    compile.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/sdk/include/ucrt", .{xwin_dir}) });
    compile.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/sdk/include/um", .{xwin_dir}) });
    compile.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/sdk/include/shared", .{xwin_dir}) });

    // Write a libc configuration file so Zig knows where the MSVC CRT lives
    const libc_conf = b.fmt(
        \\include_dir={s}/sdk/include/ucrt
        \\sys_include_dir={s}/crt/include
        \\crt_dir={s}/crt/lib/x86_64
        \\msvc_lib_dir={s}/crt/lib/x86_64
        \\kernel32_lib_dir={s}/sdk/lib/um/x86_64
        \\gcc_dir=
        \\
    , .{ xwin_dir, xwin_dir, xwin_dir, xwin_dir, xwin_dir });

    const libc_file = b.addWriteFiles().add("libc.conf", libc_conf);
    compile.setLibCFile(libc_file);
}
