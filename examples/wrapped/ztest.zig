pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add worksheets.
    const workbook = xwz.initWorkBook(null, xlsx_path, null) catch |err| {
        std.debug.print("initWorkBook error: {} - {s}\n", .{ err, xwz.strError(err) });
        return err;
    };
    // const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);
    try worksheet.writeString(.{ .row = 0, .col = 0 }, "qwe", .default);
    try worksheet.writeString(.{ .row = 0, .col = 5 }, "asd", .default);
    try worksheet.writeString(.{ .row = 7, .col = 7 }, "zxc", .default);
    // try worksheet.writeString(0, 5, "asd");
    // try worksheet.writeString(7, 7, "zxc");
    try worksheet.setHPageBreaks(&.{ 2, 6, 10 });
    _ = try workbook.addWorkSheet("qwe");
    const qwe = try workbook.getWorkSheetByName("qwe");
    _ = qwe;

    try workbook.defineName("Exchange_rate", "=0.96");
    try workbook.defineName("Sales", "=Sheet1!$G$1:$H$10");
    try workbook.defineName("qwe!Sales", "=qwe!$G$1:$G$10");
}

const std = @import("std");
var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
