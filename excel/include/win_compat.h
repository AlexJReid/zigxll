/*
 * Windows type compatibility header for xlcall.h
 * Provides minimal type definitions needed by xlcall.h and FRAMEWRK.H
 * Used on all platforms — avoids pulling in full windows.h which
 * can fail with Zig's @cImport on native MSVC installs.
 */

#ifndef WIN_COMPAT_H
#define WIN_COMPAT_H

/* Avoid <stdint.h>/<stddef.h> — on native MSVC they pull in vcruntime.h
 * which uses __int64 and other constructs that Zig's @cImport (clang) cannot parse. */

/* size_t for FRAMEWRK.H */
#ifdef _WIN64
typedef unsigned long long size_t;
#else
typedef unsigned long size_t;
#endif

/* Basic integer types */
typedef int INT32;
typedef unsigned int UINT32;
typedef unsigned int DWORD;
typedef unsigned short WORD;
typedef unsigned char BYTE;
/* BOOL is defined by xlcall.h as INT32 */
typedef long LONG;

/* Pointer-sized types — always targeting x86_64 */
#ifdef _WIN64
typedef unsigned long long DWORD_PTR;
typedef unsigned long long ULONG_PTR;
#else
typedef unsigned long DWORD_PTR;
typedef unsigned long ULONG_PTR;
#endif

/* Wide character types */
typedef unsigned short WCHAR;
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
