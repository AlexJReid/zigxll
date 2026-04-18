const std = @import("std");
const builtin = @import("builtin");

// Framework build
pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{
        .default_target = .{ .cpu_arch = .x86_64, .os_tag = .windows, .abi = .msvc },
    });
    const optimize = b.standardOptimizeOption(.{ .preferred_optimize_mode = .ReleaseSmall });
    const enable_lua = b.option(bool, "lua", "Enable Lua scripting support") orelse false;
    const lua_states = b.option(u32, "lua_states", "Number of Lua states in pool (default: CPU cores, clamped 2-16)") orelse 0;

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
    framework_build_options.addOption(u32, "lua_states", lua_states);

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

    xll.root_module.addIncludePath(b.path("excel/include"));
    xll.root_module.addLibraryPath(b.path("excel/lib"));

    if (enable_lua) {
        addLuaPaths(b, xll.root_module);
    }

    // MSVC CRT stubs — safe in XLL context (Excel already initialized the CRT)
    if (target.result.os.tag == .windows) {
        xll.root_module.addCSourceFiles(.{ .root = b.path("src"), .files = &.{"msvc_stubs.c"} });
    }

    if (builtin.os.tag == .windows) {
        addNativeMsvcPaths(b, xll.root_module);
    } else {
        addXwinPaths(b, xll, xll.root_module);
    }

    xll.root_module.link_libc = true;
    xll.root_module.linkSystemLibrary("user32", .{});
    xll.root_module.linkSystemLibrary("xlcall32", .{});
    xll.root_module.linkSystemLibrary("frmwrk32", .{});
    xll.root_module.linkSystemLibrary("vcruntime", .{});
    xll.root_module.linkSystemLibrary("ucrt", .{});

    const install_xll = b.addInstallFile(xll.getEmittedBin(), "lib/output.xll");
    b.getInstallStep().dependOn(&install_xll.step);

    // Add test step - uses native target so tests can run on Mac/Linux
    const native_target = b.resolveTargetQuery(.{});
    const test_options = b.addOptions();
    test_options.addOption(bool, "enable_lua", enable_lua);
    const tests = b.addTest(.{
        .root_module = b.createModule(.{
            .root_source_file = b.path("src/tests.zig"),
            .target = native_target,
            .optimize = optimize,
        }),
    });
    tests.root_module.addImport("test_options", test_options.createModule());
    tests.root_module.addIncludePath(b.path("excel/include"));
    if (enable_lua) {
        addLuaPaths(b, tests.root_module);
    }
    tests.root_module.link_libc = true;

    // Only link Excel libraries on Windows
    if (native_target.result.os.tag == .windows) {
        tests.root_module.addLibraryPath(b.path("excel/lib"));
        tests.root_module.linkSystemLibrary("xlcall32", .{});
        tests.root_module.linkSystemLibrary("frmwrk32", .{});
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
        enable_lua: bool = false,
        lua_states: u32 = 0,
        lua_json: ?std.Build.LazyPath = null,
        /// Lua script files to embed and generate Excel function declarations from.
        /// The framework runs lua_introspect.lua on these files and generates a Zig module
        /// with LuaFunction declarations and embedded script sources.
        lua_scripts: []const []const u8 = &.{},
        /// Directory to scan for .lua files (alternative to listing them individually).
        lua_scripts_dir: ?[]const u8 = null,
        lua_prefix: []const u8 = "FUNCS.",
        lua_category: []const u8 = "Lua Functions",
    },
) *std.Build.Step.Compile {
    const target = options.target;
    const optimize = options.optimize;

    // Create build options to pass XLL name and version to framework
    const build_options = b.addOptions();
    build_options.addOption([]const u8, "xll_name", options.name);
    build_options.addOption([]const u8, "framework_version", @import("build.zig.zon").version);
    build_options.addOption(u32, "lua_states", options.lua_states);

    // Allow overriding lua_prefix and lua_category from command line (-Dlua_prefix=MyLib.)
    const lua_prefix = b.option([]const u8, "lua_prefix", "Prefix for auto-generated Lua function names (default: \"FUNCS.\")") orelse options.lua_prefix;
    const lua_category = b.option([]const u8, "lua_category", "Category for Lua functions (default: \"Lua Functions\")") orelse options.lua_category;

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

    // Generate Lua function definitions from JSON, or provide empty stub
    {
        const wf = b.addWriteFiles();
        const gen_source = if (options.lua_json) |json_path| blk: {
            const lua_json_gen = @import("src/lua_json_gen.zig");
            const path3 = json_path.getPath3(b, null);
            const json_bytes = path3.root_dir.handle.readFileAlloc(
                b.allocator,
                path3.sub_path,
                1024 * 1024,
            ) catch @panic("Failed to read lua_json file");
            const generated_src = lua_json_gen.generate(b.allocator, json_bytes) catch
                @panic("Failed to generate Lua function definitions from JSON");
            break :blk wf.add("lua_json_functions.zig", generated_src);
        } else wf.add("lua_json_functions.zig", "// No JSON Lua functions configured\n");
        const lua_json_module = b.createModule(.{
            .root_source_file = gen_source,
            .target = target,
            .optimize = optimize,
        });
        lua_json_module.addImport("xll", xll_framework);
        xll.root_module.addImport("lua_json_module", lua_json_module);
    }

    // Generate Lua function declarations from annotated .lua scripts, or provide empty stub
    {
        // Collect scripts from both explicit list and directory scan
        var all_scripts: std.ArrayListUnmanaged([]const u8) = .empty;
        for (options.lua_scripts) |s| all_scripts.append(b.allocator, s) catch @panic("OOM");
        if (options.lua_scripts_dir) |dir_path| {
            var dir = std.fs.cwd().openDir(dir_path, .{ .iterate = true }) catch
                @panic("cannot open lua_scripts_dir");
            defer dir.close();
            var it = dir.iterate();
            while (it.next() catch @panic("lua_scripts_dir iterate failed")) |entry| {
                if (entry.kind == .file and std.mem.endsWith(u8, entry.name, ".lua")) {
                    const full = std.fs.path.join(b.allocator, &.{ dir_path, entry.name }) catch @panic("OOM");
                    all_scripts.append(b.allocator, full) catch @panic("OOM");
                }
            }
        }

        const gen_source: std.Build.LazyPath = if (all_scripts.items.len > 0) blk: {
            const lua_gen = b.addSystemCommand(&.{
                "lua",
                xll_dep.path("tools/lua_introspect.lua").getPath(b),
            });
            lua_gen.setCwd(b.path("."));
            lua_gen.addArgs(&.{ "--prefix", lua_prefix, "--category", lua_category, "--embed-root", "src" });
            for (all_scripts.items) |script| lua_gen.addArg(script);
            const lua_generated = lua_gen.captureStdOut();

            // Write generated file to user's source tree (for IDE support and @embedFile resolution)
            const update_src = b.addUpdateSourceFiles();
            update_src.addCopyFileToSource(lua_generated, "src/lua_generated.zig");
            xll.step.dependOn(&update_src.step);

            // Use the source-tree copy as module root so @embedFile resolves relative to src/
            break :blk b.path("src/lua_generated.zig");
        } else blk: {
            const wf = b.addWriteFiles();
            break :blk wf.add("lua_scripts_gen.zig", "// No Lua scripts configured\n");
        };
        const lua_scripts_mod = b.createModule(.{
            .root_source_file = gen_source,
            .target = target,
            .optimize = optimize,
        });
        lua_scripts_mod.addImport("xll", xll_framework);
        xll.root_module.addImport("lua_scripts_module", lua_scripts_mod);
    }

    // Add Excel SDK from xll dependency
    const excel_include = xll_dep.path("excel/include");
    const excel_lib = xll_dep.path("excel/lib");

    xll.root_module.addIncludePath(excel_include);
    xll.root_module.addLibraryPath(excel_lib);

    if (options.enable_lua) {
        addLuaFromDep(xll_dep, xll.root_module);
    }

    // MSVC CRT stubs — safe in XLL context (Excel already initialized the CRT)
    if (target.result.os.tag == .windows) {
        xll.root_module.addCSourceFiles(.{ .root = xll_dep.path("src"), .files = &.{"msvc_stubs.c"} });
    }

    // Add Windows SDK/CRT paths to both the XLL compile step and the user module,
    // so that any C code the user compiles (e.g. nats.c) can find vcrt/ucrt headers and libs.
    if (builtin.os.tag == .windows) {
        addNativeMsvcPaths(b, xll.root_module);
    } else {
        addXwinPaths(b, xll, xll.root_module);
        applyXwinToModule(b, options.user_module);
    }

    xll.root_module.link_libc = true;
    xll.root_module.linkSystemLibrary("user32", .{});
    xll.root_module.linkSystemLibrary("xlcall32", .{});
    xll.root_module.linkSystemLibrary("frmwrk32", .{});

    // COM/RTD support
    xll.root_module.linkSystemLibrary("oleaut32", .{});
    xll.root_module.linkSystemLibrary("advapi32", .{});
    xll.root_module.linkSystemLibrary("ole32", .{});
    xll.root_module.linkSystemLibrary("vcruntime", .{});
    xll.root_module.linkSystemLibrary("ucrt", .{});

    return xll;
}

/// Compile Lua 5.4 from source (from xll dependency)
fn addLuaFromDep(xll_dep: *std.Build.Dependency, mod: *std.Build.Module) void {
    const lua_src = xll_dep.path("deps/lua/src");
    mod.addIncludePath(lua_src);
    const xll_framework = xll_dep.module("xll");
    xll_framework.addIncludePath(lua_src);

    mod.addCSourceFiles(.{ .root = lua_src, .files = &lua_sources });
}

/// Compile Lua 5.4 from source (for framework's own build)
fn addLuaPaths(b: *std.Build, mod: *std.Build.Module) void {
    mod.addIncludePath(b.path("deps/lua/src"));
    mod.addCSourceFiles(.{ .root = b.path("deps/lua/src"), .files = &lua_sources });
}

const lua_sources = .{
    "lapi.c",     "lauxlib.c",  "lbaselib.c", "lcode.c",
    "lcorolib.c", "lctype.c",   "ldblib.c",   "ldebug.c",
    "ldo.c",      "ldump.c",    "lfunc.c",    "lgc.c",
    "linit.c",    "liolib.c",   "llex.c",     "lmathlib.c",
    "lmem.c",     "loadlib.c",  "lobject.c",  "lopcodes.c",
    "loslib.c",   "lparser.c",  "lstate.c",   "lstring.c",
    "lstrlib.c",  "ltable.c",   "ltablib.c",  "ltm.c",
    "lundump.c",  "lutf8lib.c", "lvm.c",      "lzio.c",
};

fn requireEnv(b: *std.Build, name: []const u8) []const u8 {
    return b.graph.environ_map.get(name) orelse {
        std.log.err("Missing environment variable '{s}'. Run from a Visual Studio Developer Command Prompt (vcvarsall.bat).", .{name});
        @panic("MSVC environment not configured");
    };
}

/// On native Windows, use VCToolsInstallDir / WindowsSdkDir env vars to locate the MSVC CRT.
/// These are set by the Visual Studio Developer Command Prompt (vcvarsall.bat).
fn addNativeMsvcPaths(b: *std.Build, mod: *std.Build.Module) void {
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

    mod.addLibraryPath(.{ .cwd_relative = msvc_lib_dir });
    mod.addLibraryPath(.{ .cwd_relative = ucrt_lib_dir });
    mod.addLibraryPath(.{ .cwd_relative = kernel32_lib_dir });

    mod.addSystemIncludePath(.{ .cwd_relative = vctools_inc });
    mod.addSystemIncludePath(.{ .cwd_relative = ucrt_inc_dir });
    mod.addSystemIncludePath(.{ .cwd_relative = um_inc_dir });
    mod.addSystemIncludePath(.{ .cwd_relative = shared_inc_dir });
}

/// When cross-compiling from Mac/Linux, add xwin system include paths to a module
/// so C code compiled within it can find headers like <vcruntime.h>, <corecrt.h>, <windows.h>.
/// On native Windows, Zig finds these automatically.
fn applyXwinToModule(b: *std.Build, mod: *std.Build.Module) void {
    if (builtin.os.tag == .windows) return;

    const home = b.graph.environ_map.get("HOME") orelse return;
    const xwin_dir = std.Io.Dir.path.join(b.allocator, &.{ home, ".xwin" }) catch return;
    // Check if xwin directory exists
    std.Io.Dir.accessAbsolute(b.graph.io, xwin_dir, .{}) catch return;

    mod.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/crt/include", .{xwin_dir}) });
    mod.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/sdk/include/ucrt", .{xwin_dir}) });
    mod.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/sdk/include/um", .{xwin_dir}) });
    mod.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/sdk/include/shared", .{xwin_dir}) });
}

/// If ~/.xwin exists (installed via `brew install xwin && xwin --accept-license splat --output ~/.xwin`),
/// add its Windows SDK and CRT paths so we can cross-compile to Windows from Mac/Linux.
fn addXwinPaths(b: *std.Build, compile: *std.Build.Step.Compile, mod: *std.Build.Module) void {
    // xwin is only used for cross-compiling to Windows from Mac/Linux
    const home = b.graph.environ_map.get("HOME") orelse return;
    const xwin_dir = std.Io.Dir.path.join(b.allocator, &.{ home, ".xwin" }) catch return;
    // Check if xwin directory exists
    std.Io.Dir.accessAbsolute(b.graph.io, xwin_dir, .{}) catch return;

    mod.addLibraryPath(.{ .cwd_relative = b.fmt("{s}/sdk/lib/um/x86_64", .{xwin_dir}) });
    mod.addLibraryPath(.{ .cwd_relative = b.fmt("{s}/crt/lib/x86_64", .{xwin_dir}) });
    mod.addLibraryPath(.{ .cwd_relative = b.fmt("{s}/sdk/lib/ucrt/x86_64", .{xwin_dir}) });

    mod.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/crt/include", .{xwin_dir}) });
    mod.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/sdk/include/ucrt", .{xwin_dir}) });
    mod.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/sdk/include/um", .{xwin_dir}) });
    mod.addSystemIncludePath(.{ .cwd_relative = b.fmt("{s}/sdk/include/shared", .{xwin_dir}) });

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
