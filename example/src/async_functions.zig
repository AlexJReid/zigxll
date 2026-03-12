// Async Excel function examples
//
// These demonstrate both fire-and-forget async and async with
// intermediate values via yield.  In Excel:
//
//   =ZigXLL.SLOW_DOUBLE(42)     — shows #N/A for ~3s, then 84.0
//   =ZigXLL.MONTE_CARLO(100,10000) — shows progress updates, then final estimate

const xll = @import("xll");
const ExcelFunction = xll.ExcelFunction;
const ParamMeta = xll.ParamMeta;
const AsyncContext = xll.AsyncContext;
const sleepMs = xll.async_infra.sleepMs;

// ============================================================================
// Fire-and-forget async — identical signature to a sync function
// ============================================================================

pub const slow_double = ExcelFunction(.{
    .name = "ZigXLL.SLOW_DOUBLE",
    .description = "Double a number (async — simulates slow computation)",
    .category = "Zig Async",
    .is_async = true,
    .params = &[_]ParamMeta{
        .{ .name = "x", .description = "Number to double" },
    },
    .func = slowDoubleImpl,
});

fn slowDoubleImpl(x: f64) !f64 {
    // Simulate expensive work — cell shows #N/A during this sleep
    sleepMs(3000);
    return x * 2.0;
}

// ============================================================================
// Async with intermediate values via yield
// ============================================================================

pub const monte_carlo = ExcelFunction(.{
    .name = "ZigXLL.MONTE_CARLO",
    .description = "Estimate pi via Monte Carlo (async with progress)",
    .category = "Zig Async",
    .is_async = true,
    .params = &[_]ParamMeta{
        .{ .name = "batches", .description = "Number of batches (more = slower, more accurate)" },
        .{ .name = "samples_per_batch", .description = "Samples per batch" },
    },
    .func = monteCarloImpl,
});

fn monteCarloImpl(batches_f: f64, samples_f: f64, ctx: *AsyncContext) !f64 {
    const batch_count: u32 = @intFromFloat(@max(1, batches_f));
    const samples_per_batch: u32 = @intFromFloat(@max(100, samples_f));

    ctx.yield(.{ .string = "Starting..." });

    var total_inside: u64 = 0;
    var total_samples: u64 = 0;

    // Simple xorshift RNG (no OS deps, works on thread pool)
    var rng_state: u64 = 0x12345678_9ABCDEF0;
    const xorshift = struct {
        fn next(state: *u64) f64 {
            state.* ^= state.* << 13;
            state.* ^= state.* >> 7;
            state.* ^= state.* << 17;
            return @as(f64, @floatFromInt(state.* & 0x7FFFFFFF)) / @as(f64, 0x7FFFFFFF);
        }
    };

    for (0..batch_count) |batch| {
        var inside: u64 = 0;
        for (0..samples_per_batch) |_| {
            const x = xorshift.next(&rng_state);
            const y = xorshift.next(&rng_state);
            if (x * x + y * y <= 1.0) inside += 1;
        }

        total_inside += inside;
        total_samples += samples_per_batch;

        const estimate = 4.0 * @as(f64, @floatFromInt(total_inside)) / @as(f64, @floatFromInt(total_samples));

        // Yield progress — cell updates with current estimate
        ctx.yield(.{ .double = estimate });

        // Pace the updates so Excel can keep up
        if (batch + 1 < batch_count) {
            sleepMs(500);
        }
    }

    return 4.0 * @as(f64, @floatFromInt(total_inside)) / @as(f64, @floatFromInt(total_samples));
}
