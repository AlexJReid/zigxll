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

// New (requires io parameter)
const bytes = dir.handle.readFileAlloc(b.graph.io, allocator, sub_path, max_size) catch ...;
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

### ArrayList Writer

```zig
// Old
var list = std.ArrayListUnmanaged(u8){};
const w = list.writer(allocator);
// ... write using w ...
return list.toOwnedSlice(allocator);

// New
var writer = std.Io.Writer.Allocating.init(allocator);
errdefer writer.deinit();
const w = &writer.writer;
// ... write using w ...
return writer.toOwnedSlice();
```

## Cross-Compilation (xwin) Workaround

Zig 0.16 may request `MSVCRTD.lib` (debug CRT) even in release builds when cross-compiling to Windows. The xwin SDK only includes release libraries.

**Workaround:** Create a symlink from the release library:

```bash
ln -sf ~/.xwin/crt/lib/x86_64/msvcrt.lib ~/.xwin/crt/lib/x86_64/MSVCRTD.lib
```

## Reference

- [Zig 0.16.0 Release Notes](https://ziglang.org/download/0.16.0/release-notes.html)
