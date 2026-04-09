const std = @import("std");
const async_cache = @import("async_cache.zig");
const xl_imports = @import("xl_imports.zig");
const xl = xl_imports.xl;

const allocator = std.heap.c_allocator;

fn makeNumericXloper(val: f64) *xl.XLOPER12 {
    const ptr = allocator.create(xl.XLOPER12) catch unreachable;
    ptr.* = .{
        .xltype = xl.xltypeNum | xl.xlbitDLLFree,
        .val = .{ .num = val },
    };
    return ptr;
}

fn freeXloper(ptr: *xl.XLOPER12) void {
    allocator.destroy(ptr);
}

test "cache miss returns null" {
    var cache = async_cache.AsyncCache.init();
    defer cache.clear();

    try std.testing.expect(cache.get("nonexistent") == null);
}

test "put and get" {
    var cache = async_cache.AsyncCache.init();
    defer cache.clear();

    const xloper = makeNumericXloper(42.0);
    defer freeXloper(xloper);

    cache.put("key1", .{ .xloper = xloper, .completed = true });

    const result = cache.get("key1");
    try std.testing.expect(result != null);
    try std.testing.expect(result.?.completed);
    try std.testing.expectEqual(42.0, result.?.xloper.val.num);
}

test "put overwrites existing key" {
    var cache = async_cache.AsyncCache.init();
    defer cache.clear();

    const xloper1 = makeNumericXloper(1.0);
    defer freeXloper(xloper1);
    const xloper2 = makeNumericXloper(2.0);
    defer freeXloper(xloper2);

    cache.put("key", .{ .xloper = xloper1, .completed = false });
    cache.put("key", .{ .xloper = xloper2, .completed = true });

    const result = cache.get("key");
    try std.testing.expect(result != null);
    try std.testing.expect(result.?.completed);
    try std.testing.expectEqual(2.0, result.?.xloper.val.num);
}

test "contains" {
    var cache = async_cache.AsyncCache.init();
    defer cache.clear();

    const xloper = makeNumericXloper(0.0);
    defer freeXloper(xloper);

    try std.testing.expect(!cache.contains("key"));
    cache.put("key", .{ .xloper = xloper, .completed = false });
    try std.testing.expect(cache.contains("key"));
}

test "clear removes all entries" {
    var cache = async_cache.AsyncCache.init();

    const xloper1 = makeNumericXloper(1.0);
    defer freeXloper(xloper1);
    const xloper2 = makeNumericXloper(2.0);
    defer freeXloper(xloper2);

    cache.put("a", .{ .xloper = xloper1, .completed = true });
    cache.put("b", .{ .xloper = xloper2, .completed = true });

    cache.clear();

    try std.testing.expect(cache.get("a") == null);
    try std.testing.expect(cache.get("b") == null);
}

test "in-progress then completed" {
    var cache = async_cache.AsyncCache.init();
    defer cache.clear();

    const pending = makeNumericXloper(0.0);
    defer freeXloper(pending);
    const done = makeNumericXloper(99.0);
    defer freeXloper(done);

    // Mark in-progress
    cache.put("calc|5", .{ .xloper = pending, .completed = false });
    const r1 = cache.get("calc|5");
    try std.testing.expect(!r1.?.completed);

    // Complete
    cache.put("calc|5", .{ .xloper = done, .completed = true });
    const r2 = cache.get("calc|5");
    try std.testing.expect(r2.?.completed);
    try std.testing.expectEqual(99.0, r2.?.xloper.val.num);
}

test "concurrent reads and writes" {
    var cache = async_cache.AsyncCache.init();
    defer cache.clear();

    const xloper = makeNumericXloper(42.0);
    defer freeXloper(xloper);
    cache.put("shared", .{ .xloper = xloper, .completed = true });

    const Reader = struct {
        fn run(c: *async_cache.AsyncCache) void {
            for (0..1000) |_| {
                if (c.get("shared")) |r| {
                    std.debug.assert(r.xloper.val.num == 42.0);
                }
            }
        }
    };

    var threads: [4]std.Thread = undefined;
    for (&threads) |*t| {
        t.* = std.Thread.spawn(.{}, Reader.run, .{&cache}) catch unreachable;
    }
    for (&threads) |t| t.join();
}

test "global cache singleton" {
    const c1 = async_cache.getGlobalCache();
    const c2 = async_cache.getGlobalCache();
    try std.testing.expectEqual(c1, c2);
}
