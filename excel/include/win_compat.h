/*
 * Windows type compatibility header for xlcall.h
 * Provides minimal type definitions needed by xlcall.h and FRAMEWRK.H
 * Used on all platforms — avoids pulling in full windows.h which
 * can fail with Zig's @cImport on native MSVC installs.
 */

#ifndef WIN_COMPAT_H
#define WIN_COMPAT_H

#include <stddef.h>
#include <stdint.h>

/* Basic integer types */
typedef int32_t INT32;
typedef uint32_t UINT32;
typedef uint32_t DWORD;
typedef uint16_t WORD;
typedef uint8_t BYTE;
/* BOOL is defined by xlcall.h as INT32 */
typedef long LONG;

/* Pointer-sized types */
typedef uintptr_t DWORD_PTR;
typedef uintptr_t ULONG_PTR;

/* Wide character types */
typedef uint16_t WCHAR;
typedef WCHAR* LPWSTR;
typedef const WCHAR* LPCWSTR;

/* String types */
typedef char* LPSTR;
typedef const char* LPCSTR;

/* Void and handle types */
typedef void VOID;
typedef void* PVOID;
typedef void* LPVOID;
typedef void* HANDLE;
typedef void* HWND;

/* Windows structures */
typedef struct tagPOINT {
    LONG x;
    LONG y;
} POINT;

/* Calling convention macros — no-op for Zig's @cImport.
 * Zig handles calling conventions via its own ABI, not C macros.
 * Always define these regardless of platform. */
#ifndef CALLBACK
#define CALLBACK
#endif
#ifndef WINAPI
#define WINAPI
#endif
#undef pascal
#define pascal
#undef cdecl
#define cdecl
#undef _cdecl
#define _cdecl
#undef far
#define far

/* TRUE/FALSE if not defined */
#ifndef TRUE
#define TRUE 1
#endif
#ifndef FALSE
#define FALSE 0
#endif

#endif /* WIN_COMPAT_H */
