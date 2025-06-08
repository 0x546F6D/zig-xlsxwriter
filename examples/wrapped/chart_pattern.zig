pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Add a bold format to use to highlight the header cells.
    const bold = try workbook.addFormat();
    bold.setBold();

    // Write some data for the chart.
    try worksheet.writeString(.{}, "Shingle", bold);
    try worksheet.writeNumber(.{ .row = 1 }, 105, .default);
    try worksheet.writeNumber(.{ .row = 2 }, 150, .default);
    try worksheet.writeNumber(.{ .row = 3 }, 130, .default);
    try worksheet.writeNumber(.{ .row = 4 }, 90, .default);
    try worksheet.writeString(.{ .col = 1 }, "Brick", bold);
    try worksheet.writeNumber(.{ .row = 1, .col = 1 }, 50, .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 1 }, 120, .default);
    try worksheet.writeNumber(.{ .row = 3, .col = 1 }, 100, .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 1 }, 110, .default);

    // Create a chart object.
    const chart = try workbook.addChart(.column);

    // Configure the chart.
    const series1 = try chart.addSeries(null, "Sheet1!$A$2:$A$5");
    const series2 = try chart.addSeries(null, "Sheet1!$B$2:$B$5");

    series1.setName("=Sheet1!$A$1");
    series2.setName("=Sheet1!$B$1");

    chart.titleSetName("Cladding types");
    chart.x_axis.setName("Region");
    chart.y_axis.setName("Number of houses");

    // Configure and add the chart series patterns.
    const pattern1: xwz.ChartPattern = .{
        .type = .shingle,
        .fg_color = @enumFromInt(0x804000),
        .bg_color = @enumFromInt(0xC68C53),
    };

    const pattern2: xwz.ChartPattern = .{
        .type = .horizontal_brick,
        .fg_color = @enumFromInt(0xB30000),
        .bg_color = @enumFromInt(0xFF6666),
    };

    series1.setPattern(pattern1);
    series2.setPattern(pattern2);

    // Configure and set the chart series borders.
    const line1: xwz.ChartLine = .{
        .color = @enumFromInt(0x804000),
    };

    const line2: xwz.ChartLine = .{
        .color = @enumFromInt(0xb30000),
    };

    series1.setLine(line1);
    series2.setLine(line2);

    // Widen the gap between the series/categories.
    chart.setSeriesGap(70);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 1, .col = 3 }, chart, null);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
