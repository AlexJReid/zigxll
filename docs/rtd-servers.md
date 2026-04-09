---
layout: default
title: RTD Servers
---

# RTD Servers

ZigXLL includes a reusable RTD (Real-Time Data) server framework, letting you push live data into Excel from Zig with no C++ or ATL required.

## Overview

An RTD server is a COM object that Excel polls for updated values. The framework handles all COM boilerplate (IUnknown, IDispatch, IClassFactory, vtables, DLL exports) and auto-registration. You just implement a handler with your data logic.

## Quick start

Define a handler struct and wire it up with `RtdServer`:

```zig
const rtd = @import("rtd.zig");

const MyHandler = struct {
    pub fn onStart(_: *MyHandler, ctx: *rtd.RtdContext) void {
        // Called when Excel starts the RTD server.
        // Spawn threads, open connections, etc.
    }

    pub fn onConnect(_: *MyHandler, ctx: *rtd.RtdContext, topic_id: i32, topic_count: usize) void {
        // A cell subscribed to a topic. Fires immediately.
        // Topic strings available via ctx.topics.get(topic_id).?.strings
    }

    pub fn onConnectBatch(_: *MyHandler, ctx: *rtd.RtdContext, topic_ids: []const i32) void {
        // Optional: called once per RefreshData with all new topic_ids.
        // Ideal for bulk broker subscriptions.
    }

    pub fn onDisconnect(_: *MyHandler, ctx: *rtd.RtdContext, topic_id: i32, topic_count: usize) void {
        // A cell unsubscribed from a topic.
    }

    pub fn onRefreshValue(_: *MyHandler, ctx: *rtd.RtdContext, topic_id: i32) rtd.RtdValue {
        // Return the current value for a topic.
        return .{ .double = 3.14 };
    }

    pub fn onTerminate(_: *MyHandler, ctx: *rtd.RtdContext) void {
        // Called when Excel shuts down the RTD server.
        // Clean up threads, connections, etc.
    }
};

const MyRtd = rtd.RtdServer(MyHandler, .{
    .clsid = rtd.guid("A1B2C3D4-E5F6-7890-1234-567890ABCDEF"),
    .prog_id = "myapp.rtd",
});

// In your main/root module, force the DLL exports:
comptime { MyRtd.exportDllFunctions(); }
```

In Excel, use `=RTD("myapp.rtd", , "some_topic")` to subscribe.

## Handler callbacks

| Callback | When it's called |
|---|---|
| `onStart` | Excel starts the RTD session (ServerStart). Set up your data source here. |
| `onConnect` | A cell subscribes to a topic. Fires immediately per-topic. `topic_count` is the new total. |
| `onConnectBatch` | *(optional)* Called once in RefreshData with all topic_ids that connected since the last refresh. Ideal for bulk subscription to external brokers. |
| `onDisconnect` | A cell unsubscribes. `topic_count` is the new total. |
| `onRefreshValue` | Excel wants the current value for a topic. Return an `RtdValue`. |
| `onTerminate` | Excel is shutting down the RTD session. Tear down resources. |

Both `onConnect` and `onConnectBatch` fire — they are not mutually exclusive. Use `onConnect` for per-topic bookkeeping (e.g. mapping topic_id to a key) and `onConnectBatch` for bulk external operations (e.g. subscribing to a batch of NATS subjects in one call).

## RtdContext

Your handler receives an `RtdContext` pointer with:

- **`update_event`** — Excel's callback interface. Usually you don't touch this directly.
- **`topics`** — `AutoHashMap(i32, TopicEntry)` of topic_id → entry. Each `TopicEntry` has a `dirty` flag and a `strings` field (see below).
- **`pending_connects`** — List of topic_ids that connected since the last RefreshData. Managed by the framework — you read it in `onConnectBatch`.
- **`topic_count`** — Number of currently active topics.
- **`user_data`** — An `?*anyopaque` pointer for your own state. Cast your allocated state into this in `onStart` and retrieve it in other callbacks.
- **`notifyExcel()`** — Call this to tell Excel that new data is available. Excel will then call `onRefreshValue` for dirty topics.
- **`markAllDirty()`** — Marks all active topics as dirty, so the next `RefreshData` cycle includes them all.

## Topic strings

When Excel calls `ConnectData`, it passes an array of topic strings (the arguments after the ProgID in `=RTD("prog.id", , "topic1", "topic2")`). These are extracted, converted to UTF-8, and stored in the `TopicEntry`:

```zig
pub fn onConnect(self: *MyHandler, ctx: *rtd.RtdContext, topic_id: i32, _: usize) void {
    if (ctx.topics.get(topic_id)) |entry| {
        // entry.strings is []const []const u8
        // e.g. for =RTD("myprog.rtd", , "NYSE", "AAPL") → .{ "NYSE", "AAPL" }
        for (entry.strings) |s| {
            std.log.info("topic string: {s}", .{s});
        }
    }
}
```

The strings are also available in `onConnectBatch` — look up each topic_id in `ctx.topics`:

```zig
pub fn onConnectBatch(self: *MyHandler, ctx: *rtd.RtdContext, topic_ids: []const i32) void {
    // Collect all subjects and subscribe in one broker call
    for (topic_ids) |tid| {
        if (ctx.topics.get(tid)) |entry| {
            const subject = entry.strings[0]; // e.g. "prices.AAPL"
            // ...
        }
    }
    broker.subscribeBatch(subjects);
}
```

## Auto-registration

The `rtd_registry.zig` module handles COM registration automatically. When the XLL loads (`xlAutoOpen`), it writes the necessary registry entries under `HKEY_CURRENT_USER` so no admin privileges are needed. Registration is idempotent.

The registry entries map your `prog_id` to your `clsid` and point `InprocServer32` at the XLL path, with `ThreadingModel = Apartment`.

## Pushing updates

A typical pattern is to start a background thread in `onStart` that fetches data and notifies Excel:

```zig
pub fn onStart(ctx: *rtd.RtdContext) void {
    // Store context for the background thread
    // Start a thread that periodically:
    //   1. Updates your data
    //   2. Calls ctx.markAllDirty()
    //   3. Calls ctx.notifyExcel()
}
```

Excel will then call `onRefreshValue` for each dirty topic during its next refresh cycle.

## RtdValue

`onRefreshValue` returns an `RtdValue` tagged union supporting all Excel-compatible VARIANT types:

```zig
// Integer
return .{ .int = 42 };

// Double
return .{ .double = 3.14 };

// Boolean
return .{ .boolean = true };

// UTF-16 string (use comptime helper for literals)
return RtdValue.fromUtf8("hello");

// Empty cell
return .empty;
```

The framework converts these to COM VARIANTs automatically (VT_I4, VT_R8, VT_BOOL, VT_BSTR, VT_EMPTY). Strings are allocated as BSTRs via `SysAllocStringLen` — Excel owns and frees them.

## GUID helper

Use `rtd.guid()` to parse a GUID string at comptime:

```zig
const my_clsid = rtd.guid("A1B2C3D4-E5F6-7890-1234-567890ABCDEF");
// Also accepts braces:
const my_clsid2 = rtd.guid("{A1B2C3D4-E5F6-7890-1234-567890ABCDEF}");
```

## Using RTD from a UDF

You can wrap an RTD subscription in a regular Excel function using `rtd_call.subscribe()`. This lets users call `=MYPRICE("AAPL")` instead of `=RTD("myprog.rtd", , "AAPL")`:

```zig
const xll = @import("xll");
const xl = xll.xl;
const rtd_call = xll.rtd_call;
const ExcelFunction = xll.ExcelFunction;
const ParamMeta = xll.ParamMeta;

fn myPriceImpl(symbol: []const u8) !*xl.XLOPER12 {
    return rtd_call.subscribe("myprog.rtd", &.{symbol});
}

pub const myPrice = ExcelFunction(.{
    .name = "MYPRICE",
    .description = "Live price for a symbol",
    .thread_safe = false,
    .func = myPriceImpl,
    .params = &[_]ParamMeta{
        .{ .name = "symbol", .description = "Ticker symbol", .type = []const u8 },
    },
});
```

**Important:** RTD wrapper functions must set `.thread_safe = false`. The underlying `xlfRtd` call is not thread-safe and must run on Excel's main thread. (Async functions handle this automatically — `.is_async = true` forces `thread_safe = false`.)

Excel handles the RTD subscription automatically — the cell updates whenever the RTD server pushes a new value. Multiple topics are supported:

```zig
// Subscribe with multiple topic strings
return rtd_call.subscribe("myprog.rtd", &.{ exchange, symbol });
```

## Using RTD from Lua

Lua functions can subscribe to an RTD server using `xllify.rtd_subscribe`. This is the Lua equivalent of calling `rtd_call.subscribe()` from Zig:

```lua
--- Subscribe to a live NATS subject
-- @param subject string NATS subject
-- @thread_safe false
function nats_sub(subject)
    return xllify.rtd_subscribe("zigxll.connectors.nats", subject)
end
```

`xllify.rtd_subscribe(prog_id, topic1, ...)` accepts a prog ID string and up to 28 topic strings. The cell updates live whenever the RTD server pushes a new value.

**Thread safety is mandatory.** Functions calling `xllify.rtd_subscribe` must be `@thread_safe false` — `xlfRtd` must run on Excel's main thread. See [Lua Functions — RTD subscriptions](lua-functions.md#rtd-subscriptions-from-lua) for full details.

## Limitations

- No hard limit on topic count (uses dynamically-sized HashMap).
- Values support `i32`, `f64`, `bool`, UTF-16 strings, and empty (via `RtdValue` tagged union).
- Single RTD server class per XLL (one CLSID/ProgID pair).
