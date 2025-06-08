pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(alloc, xlsx_path, null);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Write some data for the chart
    try worksheet.writeNumber(.{}, 10, .default);
    try worksheet.writeNumber(.{ .row = 1 }, 40, .default);
    try worksheet.writeNumber(.{ .row = 2 }, 50, .default);
    try worksheet.writeNumber(.{ .row = 3 }, 20, .default);
    try worksheet.writeNumber(.{ .row = 4 }, 10, .default);
    try worksheet.writeNumber(.{ .row = 5 }, 50, .default);

    // Create a chart object
    const chart = try workbook.addChart(.line);

    // Configure the chart
    _ = try chart.addSeries(null, "Sheet1!$A$1:$A$6");

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .col = 2 }, chart, null);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
