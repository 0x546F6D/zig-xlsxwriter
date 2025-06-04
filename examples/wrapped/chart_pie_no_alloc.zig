pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Add a bold format to use to highlight the header cells
    const bold = try workbook.addFormat();
    bold.setBold();

    // Write some data for the chart.
    try writeWorksheetData(worksheet, bold);

    // Chart 1: Create a simple pie chart.
    var chart = try workbook.addChart(.pie);

    // Add the first series to the chart.
    var series = try chart.addSeries("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

    // Set the name for the series instead of the default "Series 1".
    series.setName("Pie sales data");

    // Add a chart title.
    chart.titleSetName("Popular Pie Types");

    // Set an Excel chart style.
    chart.setStyle(10);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 1, .col = 3 }, chart, null);

    // Chart 2: Create a pie chart with user defined segment colors.
    chart = try workbook.addChart(.pie);

    // Add the first series to the chart.
    series = try chart.addSeries("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

    // Set the name for the series instead of the default "Series 1".
    series.setName("Pie sales data");

    // Add a chart title.
    chart.titleSetName("Pie Chart with user defined colors");

    // Add fills for use in the chart.
    const fill1 = xwz.ChartFill{ .color = @enumFromInt(0x5ABA10) };
    const fill2 = xwz.ChartFill{ .color = @enumFromInt(0xFE110E) };
    const fill3 = xwz.ChartFill{ .color = @enumFromInt(0xCA5C05) };

    // Add some points with fills.
    const point1: PointNoAlloc = .{ .fill = @constCast(&fill1.toC()) };
    const point2: PointNoAlloc = .{ .fill = @constCast(&fill2.toC()) };
    const point3: PointNoAlloc = .{ .fill = @constCast(&fill3.toC()) };

    // Create an array of the point objects.
    const points_array: PointNoAllocArray = &.{
        &point1,
        &point2,
        &point3,
    };

    // Add/override the points/segments of the chart.
    try series.setPointsNoAlloc(points_array);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 17, .col = 3 }, chart, null);

    // Chart 3: Create a pie chart with rotation of the segments.
    chart = try workbook.addChart(.pie);

    // Add the first series to the chart.
    series = try chart.addSeries("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

    // Set the name for the series instead of the default "Series 1".
    series.setName("Pie sales data");

    // Add a chart title.
    chart.titleSetName("Pie Chart with segment rotation");

    // Change the angle/rotation of the first segment.
    chart.setRotation(90);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 33, .col = 3 }, chart, null);
}

// Write some data to the worksheet.
fn writeWorksheetData(worksheet: WorkSheet, bold: Format) !void {
    try worksheet.writeString(.{}, "Category", bold);
    try worksheet.writeString(.{ .row = 1 }, "Apple", .default);
    try worksheet.writeString(.{ .row = 2 }, "Cherry", .default);
    try worksheet.writeString(.{ .row = 3 }, "Pecan", .default);

    try worksheet.writeString(.{ .col = 1 }, "Values", bold);
    try worksheet.writeNumber(.{ .row = 1, .col = 1 }, 60, .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 1 }, 30, .default);
    try worksheet.writeNumber(.{ .row = 3, .col = 1 }, 10, .default);
}

const std = @import("std");
var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
const WorkSheet = xwz.WorkSheet;
const Format = xwz.Format;
const PointNoAlloc = xwz.ChartSeriesPointNoAlloc;
const PointNoAllocArray = xwz.ChartSeriesPointNoAllocArray;
