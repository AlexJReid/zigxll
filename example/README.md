# Example ZigXLL Project

Example project showing how to use ZigXLL to create custom Excel functions.

## Quick Start

```bash
zig build
```

Output: `zig-out/lib/my_excel_functions.xll`

## Example Functions

- `double(x)` - Doubles a number
- `reverse(text)` - Reverses a string
- `ZigXLLExample.BS_CALL(S, K, T, r, sigma)` - Black-Scholes call option price
- `ZigXLLExample.BS_PUT(S, K, T, r, sigma)` - Black-Scholes put option price
