# Example ZigXLL Project

Example project showing how to use ZigXLL to create custom Excel functions.

## Quick Start

```bash
zig build
```

Output: `zig-out/lib/my_excel_functions.xll`

If you download an artifact from this repo you will need to unzip the file and unblock it. More info: https://support.microsoft.com/en-gb/topic/excel-is-blocking-untrusted-xll-add-ins-by-default-1e3752e2-1177-4444-a807-7b700266a6fb

## Example Functions

- `double(x)` - Doubles a number
- `reverse(text)` - Reverses a string
- `ZigXLLExample.BS_CALL(S, K, T, r, sigma)` - Black-Scholes call option price
- `ZigXLLExample.BS_PUT(S, K, T, r, sigma)` - Black-Scholes put option price
