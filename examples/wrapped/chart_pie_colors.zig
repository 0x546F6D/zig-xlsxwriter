pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(alloc, xlsx_path, null);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Write data for the chart
    try worksheet.writeString(.{}, "Pass", .default);
    try worksheet.writeString(.{ .row = 1 }, "Fail", .default);
    try worksheet.writeNumber(.{ .col = 1 }, 90, .default);
    try worksheet.writeNumber(.{ .row = 1, .col = 1 }, 10, .default);

    // Create a pie chart
    const chart = try workbook.addChart(.pie);

    // Add the data series
    const series = try chart.addSeries("=Sheet1!$A$1:$A$2", "=Sheet1!$B$1:$B$2");

    // Create fills for chart segments
    const red_fill: xwz.ChartFill = .{ .color = .red };
    const green_fill: xwz.ChartFill = .{ .color = .green };

    // Create points with fills
    const red_point: xwz.ChartPoint = .{ .fill = &red_fill };
    const green_point: xwz.ChartPoint = .{ .fill = &green_fill };

    // Create array of points
    const points = &.{ green_point, red_point };

    // Set the points on the series
    try series.setPoints(points);

    // Insert chart into worksheet
    try worksheet.insertChart(.{ .row = 1, .col = 3 }, chart, null);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
