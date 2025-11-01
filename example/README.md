# Example ZigXLL User Project

This is an example of how to use ZigXLL as a library.

## Structure

```
src/
├── main.zig              ← Entry point (re-exports zigxll framework)
├── user_functions.zig    ← Register your function modules here
└── my_functions.zig      ← Your custom functions
```

## Building

```bash
zig build
```

Your XLL will be in `zig-out/lib/my_excel_functions.xll`

## Example Functions Included

This example includes several custom Excel functions. **Note that these are for demonstration purposes.**

- **double(x)** - Doubles a number
- **reverse(text)** - Reverses a string

- **MyFunctions.BSCall(S, K, T, r, sigma)** - Black-Scholes call option price
  - S: Current stock price
  - K: Strike price
  - T: Time to maturity (years)
  - r: Risk-free rate
  - sigma: Volatility

- **MyFunctions.BSPut(S, K, T, r, sigma)** - Black-Scholes put option price
  - S: Current stock price
  - K: Strike price
  - T: Time to maturity (years)
  - r: Risk-free rate
  - sigma: Volatility

Example usage in Excel:
```
=MyFunctions.BSCall(100, 105, 1, 0.05, 0.2)
=MyFunctions.BSPut(100, 105, 1, 0.05, 0.2)
```
