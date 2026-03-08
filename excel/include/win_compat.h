/*
 * Windows type compatibility header for non-Windows platforms
 * Provides minimal type definitions needed by xlcall.h
 */

#ifndef WIN_COMPAT_H
#define WIN_COMPAT_H

#ifndef _WIN32

#include <stddef.h>
#include <stdint.h>

/* Basic integer types */
typedef int32_t INT32;
typedef uint32_t UINT32;
typedef uint32_t DWORD;
typedef uint16_t WORD;
typedef uint8_t BYTE;
typedef int BOOL;
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

/* Calling convention macros (no-op on non-Windows) */
#define CALLBACK
#define WINAPI
#define _cdecl
#define cdecl
#define pascal
#define far

/* TRUE/FALSE if not defined */
#ifndef TRUE
#define TRUE 1
#endif
#ifndef FALSE
#define FALSE 0
#endif

#endif /* _WIN32 */

#endif /* WIN_COMPAT_H */
