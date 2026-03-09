const xll = @import("xll");
const xl = xll.xl;
const XLValue = xll.XLValue;
const ExcelMacro = xll.ExcelMacro;

const allocator = @import("std").heap.c_allocator;

pub const hello_alert = ExcelMacro(.{
    .name = "ZigXLL.HelloAlert",
    .description = "Show a hello world alert",
    .category = "Zig Macros",
    .func = helloAlert,
});

fn helloAlert() !void {
    var msg = try XLValue.fromUtf8String(allocator, "Hello from ZigXLL!");
    defer msg.deinit();
    _ = xl.Excel12f(xl.xlcAlert, null, 1, &msg.m_val);
}
