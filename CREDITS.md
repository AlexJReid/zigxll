# Credits

## Microsoft Excel SDK

This project uses the Microsoft Excel 2013 XLL SDK headers and libraries:

- **Excel 2013 XLL SDK**
- Download: https://www.microsoft.com/en-gb/download/details.aspx?id=35567
- Files used:
  - `excel/include/xlcall.h` - Excel C API definitions
  - `excel/include/FRAMEWRK.H` - Framework helper functions
  - `excel/lib/xlcall32.lib` - Excel C API library
  - `excel/lib/frmwrk32.lib` - Framework helper library

Copyright (c) Microsoft Corporation. All rights reserved.

## Development

- Built with [Zig](https://ziglang.org/) programming language
- Claude (Anthropic) helped with some of the trickier comptime bits
