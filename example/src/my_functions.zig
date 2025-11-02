// My custom Excel functions
const std = @import("std");
const xll = @import("xll");
const ExcelFunction = xll.ExcelFunction;
const ParamMeta = xll.ParamMeta;

const allocator = std.heap.c_allocator;

// Example custom function
pub const double = ExcelFunction(.{
    .name = "double",
    .description = "Double a number",
    .category = "Zig Functions",
    .params = &[_]ParamMeta{
        .{ .name = "x", .description = "Number to double" },
    },
    .func = doubleFunc,
});

fn doubleFunc(x: f64) !f64 {
    return x * 2;
}

pub const reverse = ExcelFunction(.{
    .name = "reverse",
    .description = "Reverse a string",
    .category = "Zig Functions",
    .params = &[_]ParamMeta{
        .{ .name = "text", .description = "Text to reverse" },
    },
    .func = reverseFunc,
});

fn reverseFunc(text: []const u8) ![]const u8 {
    var result = try allocator.alloc(u8, text.len);
    for (text, 0..) |c, i| {
        result[text.len - 1 - i] = c;
    }

    return result;
}

// Claude wrote these so be careful.
// Black-Scholes Call Option
pub const bs_call = ExcelFunction(.{
    .name = "ZigXLLExample.BS_CALL",
    .description = "Black-Scholes call option price",
    .category = "Zig Functions",
    .params = &[_]ParamMeta{
        .{ .name = "S", .description = "Current stock price" },
        .{ .name = "K", .description = "Strike price" },
        .{ .name = "T", .description = "Time to maturity (years)" },
        .{ .name = "r", .description = "Risk-free rate" },
        .{ .name = "sigma", .description = "Volatility" },
    },
    .func = blackScholesCall,
});

fn blackScholesCall(S: f64, K: f64, T: f64, r: f64, sigma: f64) !f64 {
    const d1 = (std.math.log(f64, std.math.e, S / K) + (r + 0.5 * sigma * sigma) * T) / (sigma * @sqrt(T));
    const d2 = d1 - sigma * @sqrt(T);

    const N_d1 = normalCDF(d1);
    const N_d2 = normalCDF(d2);

    return S * N_d1 - K * @exp(-r * T) * N_d2;
}

// Black-Scholes Put Option
pub const bs_put = ExcelFunction(.{
    .name = "ZigXLLExample.BS_PUT",
    .description = "Black-Scholes put option price",
    .category = "Zig Functions",
    .params = &[_]ParamMeta{
        .{ .name = "S", .description = "Current stock price" },
        .{ .name = "K", .description = "Strike price" },
        .{ .name = "T", .description = "Time to maturity (years)" },
        .{ .name = "r", .description = "Risk-free rate" },
        .{ .name = "sigma", .description = "Volatility" },
    },
    .func = blackScholesPut,
});

fn blackScholesPut(S: f64, K: f64, T: f64, r: f64, sigma: f64) !f64 {
    const d1 = (std.math.log(f64, std.math.e, S / K) + (r + 0.5 * sigma * sigma) * T) / (sigma * @sqrt(T));
    const d2 = d1 - sigma * @sqrt(T);

    const N_d1 = normalCDF(-d1);
    const N_d2 = normalCDF(-d2);

    return K * @exp(-r * T) * N_d2 - S * N_d1;
}

// Normal cumulative distribution function (Zelen & Severo approximation)
fn normalCDF(x: f64) f64 {
    const b0: f64 = 0.2316419;
    const b1: f64 = 0.319381530;
    const b2: f64 = -0.356563782;
    const b3: f64 = 1.781477937;
    const b4: f64 = -1.821255978;
    const b5: f64 = 1.330274429;

    if (x >= 0.0) {
        const t = 1.0 / (1.0 + b0 * x);
        const poly = b1 * t + b2 * t * t + b3 * std.math.pow(f64, t, 3) + b4 * std.math.pow(f64, t, 4) + b5 * std.math.pow(f64, t, 5);
        return 1.0 - poly * @exp(-x * x / 2.0) / @sqrt(2.0 * std.math.pi);
    } else {
        return 1.0 - normalCDF(-x);
    }
}

// Tests
test "normalCDF" {
    const expectApproxEqRel = std.testing.expectApproxEqRel;

    // Test standard normal distribution values
    try expectApproxEqRel(normalCDF(0.0), 0.5, 0.0001);
    try expectApproxEqRel(normalCDF(1.0), 0.8413, 0.001);
    try expectApproxEqRel(normalCDF(2.0), 0.9772, 0.001);
    try expectApproxEqRel(normalCDF(-1.0), 0.1587, 0.001);
}

test "Black-Scholes call option" {
    const expectApproxEqRel = std.testing.expectApproxEqRel;

    // Test case: S=100, K=100, T=1, r=0.05, sigma=0.2
    // Expected value calculated from known BS formula
    const call_price = try blackScholesCall(100.0, 100.0, 1.0, 0.05, 0.2);
    try expectApproxEqRel(call_price, 10.4506, 0.01);

    // Test in-the-money call: S > K
    const itm_call = try blackScholesCall(110.0, 100.0, 1.0, 0.05, 0.2);
    try std.testing.expect(itm_call > call_price);

    // Test out-of-the-money call: S < K
    const otm_call = try blackScholesCall(90.0, 100.0, 1.0, 0.05, 0.2);
    try std.testing.expect(otm_call < call_price);
}

test "Black-Scholes put option" {
    const expectApproxEqRel = std.testing.expectApproxEqRel;

    // Test case: S=100, K=100, T=1, r=0.05, sigma=0.2
    const put_price = try blackScholesPut(100.0, 100.0, 1.0, 0.05, 0.2);
    try expectApproxEqRel(put_price, 5.5735, 0.01);

    // Test in-the-money put: S < K
    const itm_put = try blackScholesPut(90.0, 100.0, 1.0, 0.05, 0.2);
    try std.testing.expect(itm_put > put_price);

    // Test out-of-the-money put: S > K
    const otm_put = try blackScholesPut(110.0, 100.0, 1.0, 0.05, 0.2);
    try std.testing.expect(otm_put < put_price);
}

test "Put-Call parity" {
    const expectApproxEqRel = std.testing.expectApproxEqRel;

    // Put-Call parity: C - P = S - K*e^(-rT)
    const S = 100.0;
    const K = 100.0;
    const T = 1.0;
    const r = 0.05;
    const sigma = 0.2;

    const call = try blackScholesCall(S, K, T, r, sigma);
    const put = try blackScholesPut(S, K, T, r, sigma);

    const left_side = call - put;
    const right_side = S - K * @exp(-r * T);

    try expectApproxEqRel(left_side, right_side, 0.0001);
}
