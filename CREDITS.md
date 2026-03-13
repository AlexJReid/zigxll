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

## Lua 5.4

This project embeds [Lua 5.4](https://www.lua.org/) for scripting support.

Copyright (c) 1994-2024 Lua.org, PUC-Rio.

Licensed under the MIT License. See https://www.lua.org/license.html

## Development

- Built with [Zig](https://ziglang.org/) programming language
- Claude (Anthropic) helped with some of the trickier comptime bits
