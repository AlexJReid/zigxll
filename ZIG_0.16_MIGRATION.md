# Zig 0.16 Migration Guide

This document describes the changes required to migrate zigxll from Zig 0.14/0.15 to Zig 0.16.

## build.zig Changes

### Compile Step Methods Moved to root_module

Methods that were on `*std.Build.Step.Compile` have moved to `compile.root_module`:

```zig
// Old (0.14/0.15)
xll.addIncludePath(b.path("excel/include"));
xll.addLibraryPath(b.path("excel/lib"));
xll.addCSourceFiles(.{ .root = b.path("src"), .files = &.{"file.c"} });
xll.linkLibC();
xll.linkSystemLibrary("user32");

// New (0.16)
xll.root_module.addIncludePath(b.path("excel/include"));
xll.root_module.addLibraryPath(b.path("excel/lib"));
xll.root_module.addCSourceFiles(.{ .root = b.path("src"), .files = &.{"file.c"} });
xll.root_module.link_libc = true;
xll.root_module.linkSystemLibrary("user32", .{});
```

### Environment Variables

```zig
// Old
const home = std.process.getEnvVarOwned(b.allocator, "HOME") catch return;

// New
const home = b.graph.environ_map.get("HOME") orelse return;
```

### Filesystem Access

```zig
// Old
var dir = std.fs.openDirAbsolute(path, .{}) catch return;
dir.close();

// New
std.Io.Dir.accessAbsolute(b.graph.io, path, .{}) catch return;
```

### Path Joining

```zig
// Old
const path = std.fs.path.join(allocator, &.{ home, ".xwin" }) catch return;

// New
const path = std.Io.Dir.path.join(allocator, &.{ home, ".xwin" }) catch return;
```

### File Reading

```zig
// Old
const bytes = dir.handle.readFileAlloc(allocator, sub_path, max_size) catch ...;

// New (requires io parameter, sub_path before allocator, Io.Limit for max_size)
const bytes = dir.handle.readFileAlloc(b.graph.io, sub_path, allocator, std.Io.Limit.limited(max_size)) catch ...;
```

### Directory Operations

```zig
// Old
var dir = std.fs.cwd().openDir(path, .{ .iterate = true }) catch ...;
defer dir.close();
var it = dir.iterate();
while (it.next() catch ...) |entry| { ... }

// New (cwd() takes no args, openDir/close/next require io parameter)
var dir = std.Io.Dir.cwd().openDir(b.graph.io, path, .{ .iterate = true }) catch ...;
defer dir.close(b.graph.io);
var it = dir.iterate();
while (it.next(b.graph.io) catch ...) |entry| { ... }
```

### captureStdOut

```zig
// Old
const output = run.captureStdOut();

// New (requires options struct)
const output = run.captureStdOut(.{});
```

## Source File Changes

### Mutex API

`std.Thread.Mutex` has been replaced with `std.Io.Mutex`, which requires an `Io` instance for lock/unlock operations:

```zig
// Old
var mutex: std.Thread.Mutex = .{};
mutex.lock();
defer mutex.unlock();

// New
var mutex: std.Io.Mutex = std.Io.Mutex.init;
const io = std.Options.debug_io;
mutex.lock(io) catch {};
defer mutex.unlock(io);
```

### ArrayListUnmanaged Initialization

```zig
// Old
var list = std.ArrayListUnmanaged(u8){};

// New
var list: std.ArrayListUnmanaged(u8) = .empty;
```

### ArrayList API Changes

The `writer()` method has been removed from ArrayList. Use ArrayList methods directly:

```zig
// Old
var buf = std.ArrayListUnmanaged(u8){};
errdefer buf.deinit(allocator);
const writer = buf.writer(allocator);
try writer.writeAll("hello");
try writer.writeByte('|');
try writer.print("{d}", .{42});
return buf.toOwnedSlice(allocator);

// New (use ArrayList methods directly, allocator passed to each method)
var buf: std.ArrayList(u8) = .empty;
errdefer buf.deinit(allocator);
try buf.appendSlice(allocator, "hello");
try buf.append(allocator, '|');
try buf.print(allocator, "{d}", .{42});
return buf.toOwnedSlice(allocator);
```

Note: `std.ArrayList` in 0.16 is now the unmanaged version (formerly `ArrayListUnmanaged`).
The managed version is `std.array_list.Managed` but is deprecated.

### Thread.sleep Removed

`std.Thread.sleep` has been removed. For Windows targets, use the Win32 API directly
(this is what `std.Thread.sleep` called internally anyway):

```zig
// Old
std.Thread.sleep(2 * std.time.ns_per_s);

// New (Windows) - use kernel32.Sleep (milliseconds)
const kernel32 = struct {
    extern "kernel32" fn Sleep(dwMilliseconds: u32) callconv(.winapi) void;
};
kernel32.Sleep(2000);  // 2 seconds

// Or use the zigxll helper:
xll.async_infra.sleepMs(2000);
```

### Windows Types Removed

`std.os.windows.HRESULT` has been removed. Define locally:

```zig
// Old
pub const HRESULT = std.os.windows.HRESULT;

// New
pub const HRESULT = i32;
```

## Cross-Compilation (xwin) Workaround

Zig 0.16 may request `MSVCRTD.lib` (debug CRT) even in release builds when cross-compiling to Windows. The xwin SDK only includes release libraries.

**Workaround:** Create a symlink from the release library:

```bash
ln -sf ~/.xwin/crt/lib/x86_64/msvcrt.lib ~/.xwin/crt/lib/x86_64/MSVCRTD.lib
```

## Reference

- [Zig 0.16.0 Release Notes](https://ziglang.org/download/0.16.0/release-notes.html)
