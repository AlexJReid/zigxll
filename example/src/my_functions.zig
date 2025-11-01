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

// Black-Scholes Call Option Price
pub const bs_call = ExcelFunction(.{
    .name = "MyFunctions.BSCall",
    .description = "Black-Scholes European Call Option Price",
    .category = "Finance",
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
    if (T <= 0) return error.InvalidMaturity;
    if (sigma <= 0) return error.InvalidVolatility;
    if (S <= 0) return error.InvalidStockPrice;
    if (K <= 0) return error.InvalidStrikePrice;

    const d1 = (std.math.log(f64, std.math.e, S / K) + (r + sigma * sigma / 2.0) * T) / (sigma * std.math.sqrt(T));
    const d2 = d1 - sigma * std.math.sqrt(T);

    const call_price = S * cumulativeNormal(d1) - K * @exp(-r * T) * cumulativeNormal(d2);
    return call_price;
}

// Black-Scholes Put Option Price
pub const bs_put = ExcelFunction(.{
    .name = "MyFunctions.BSPut",
    .description = "Black-Scholes European Put Option Price",
    .category = "Finance",
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
    if (T <= 0) return error.InvalidMaturity;
    if (sigma <= 0) return error.InvalidVolatility;
    if (S <= 0) return error.InvalidStockPrice;
    if (K <= 0) return error.InvalidStrikePrice;

    const d1 = (std.math.log(f64, std.math.e, S / K) + (r + sigma * sigma / 2.0) * T) / (sigma * std.math.sqrt(T));
    const d2 = d1 - sigma * std.math.sqrt(T);

    const put_price = K * @exp(-r * T) * cumulativeNormal(-d2) - S * cumulativeNormal(-d1);
    return put_price;
}

// Cumulative Normal Distribution (approximation)
fn cumulativeNormal(x: f64) f64 {
    const a1: f64 = 0.319381530;
    const a2: f64 = -0.356563782;
    const a3: f64 = 1.781477937;
    const a4: f64 = -1.821255978;
    const a5: f64 = 1.330274429;
    const gamma: f64 = 0.2316419;

    const k = 1.0 / (1.0 + gamma * @abs(x));
    const k2 = k * k;
    const k3 = k2 * k;
    const k4 = k3 * k;
    const k5 = k4 * k;

    const sqrt_2pi = std.math.sqrt(2.0 * std.math.pi);
    const pdf = @exp(-0.5 * x * x) / sqrt_2pi;

    const cdf_approx = 1.0 - pdf * (a1 * k + a2 * k2 + a3 * k3 + a4 * k4 + a5 * k5);

    if (x >= 0.0) {
        return cdf_approx;
    } else {
        return 1.0 - cdf_approx;
    }
}
