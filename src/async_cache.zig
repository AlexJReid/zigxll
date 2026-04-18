// Thread-safe cache for async Excel function results.
//
// Keyed by topic string (function name + serialized args).
// Values are stored with a `completed` flag — once true, the UDF
// returns the value directly instead of calling xlfRtd, which
// causes Excel to drop the RTD subscription automatically.

const std = @import("std");
const xl_imports = @import("xl_imports.zig");
const xl = xl_imports.xl;

const allocator = std.heap.c_allocator;

pub const CachedResult = struct {
    /// The XLOPER12 value to return to Excel.  Heap-allocated with xlbitDLLFree.
    xloper: *xl.XLOPER12,
    /// True once the async work has finished (success or error).
    completed: bool,
};

/// Thread-safe string-keyed cache of XLOPER12 results.
/// The cache owns copies of all key strings.
pub const AsyncCache = struct {
    map: std.StringHashMap(CachedResult),
    mutex: std.Io.Mutex,

    const io = std.Options.debug_io;

    pub fn init() AsyncCache {
        return .{
            .map = std.StringHashMap(CachedResult).init(allocator),
            .mutex = std.Io.Mutex.init,
        };
    }

    fn lock(self: *AsyncCache) void {
        self.mutex.lock(io) catch {};
    }

    fn unlock(self: *AsyncCache) void {
        self.mutex.unlock(io);
    }

    /// Look up a topic key.  Returns null on miss.
    pub fn get(self: *AsyncCache, key: []const u8) ?CachedResult {
        self.lock();
        defer self.unlock();
        return self.map.get(key);
    }

    /// Store a result for a topic key.
    /// If the key is new, the cache dupes it and owns the copy.
    /// If the key already exists, just updates the value.
    pub fn put(self: *AsyncCache, key: []const u8, result: CachedResult) void {
        self.lock();
        defer self.unlock();

        // If key already exists, just update the value
        if (self.map.getEntry(key)) |entry| {
            entry.value_ptr.* = result;
            return;
        }

        // New key — dupe it so the cache owns the string
        const owned_key = allocator.dupe(u8, key) catch return;
        self.map.put(owned_key, result) catch {
            allocator.free(owned_key);
        };
    }

    /// Remove all cached results, forcing async functions to re-execute.
    pub fn clear(self: *AsyncCache) void {
        self.lock();
        defer self.unlock();
        var it = self.map.iterator();
        while (it.next()) |entry| {
            allocator.free(@constCast(entry.key_ptr.*));
        }
        self.map.clearAndFree();
    }

    /// Check whether a key exists (used to avoid double-spawning).
    pub fn contains(self: *AsyncCache, key: []const u8) bool {
        self.lock();
        defer self.unlock();
        return self.map.contains(key);
    }
};

// Global singleton — all async functions share one cache.
var global_cache: ?*AsyncCache = null;
var cache_init_mutex: std.Io.Mutex = std.Io.Mutex.init;

pub fn getGlobalCache() *AsyncCache {
    const io = std.Options.debug_io;
    cache_init_mutex.lock(io) catch {};
    defer cache_init_mutex.unlock(io);
    if (global_cache) |c| return c;
    const c = allocator.create(AsyncCache) catch unreachable;
    c.* = AsyncCache.init();
    global_cache = c;
    return c;
}
