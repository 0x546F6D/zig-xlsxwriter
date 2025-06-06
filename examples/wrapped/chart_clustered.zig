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

    // Chart: Create a column chart
    var chart = try workbook.addChart(.column);

    // Configure the series. Note, that the categories are 2D ranges (from
    // column A to column B). This creates the clusters. The series are shown
    // as formula strings for clarity but you can also use variables with the
    // chart_series_set_categories() and chart_series_set_values()
    // functions.
    _ = try chart.addSeries("=Sheet1!$A$2:$B$6", "=Sheet1!$C$2:$C$6");
    _ = try chart.addSeries("=Sheet1!$A$2:$B$6", "=Sheet1!$D$2:$D$6");
    _ = try chart.addSeries("=Sheet1!$A$2:$B$6", "=Sheet1!$E$2:$E$6");

    // Set an Excel chart style
    chart.setStyle(37);

    // Turn off the legend
    chart.legendSetPosition(.none);

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .row = 2, .col = 6 }, chart, null);
}

// Write some data to the worksheet
fn writeWorksheetData(worksheet: WorkSheet, bold: Format) !void {
    try worksheet.writeString(.{}, "Types", bold);
    try worksheet.writeString(.{ .row = 1 }, "Type 1", .default);
    try worksheet.writeString(.{ .row = 4 }, "Type 2", .default);

    try worksheet.writeString(.{ .col = 1 }, "Sub Type", bold);
    try worksheet.writeString(.{ .row = 1, .col = 1 }, "Sub Type A", .default);
    try worksheet.writeString(.{ .row = 2, .col = 1 }, "Sub Type B", .default);
    try worksheet.writeString(.{ .row = 3, .col = 1 }, "Sub Type C", .default);
    try worksheet.writeString(.{ .row = 4, .col = 1 }, "Sub Type D", .default);
    try worksheet.writeString(.{ .row = 5, .col = 1 }, "Sub Type E", .default);

    try worksheet.writeString(.{ .col = 2 }, "Value 1", bold);
    try worksheet.writeNumber(.{ .row = 1, .col = 2 }, 5000, .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 2 }, 2000, .default);
    try worksheet.writeNumber(.{ .row = 3, .col = 2 }, 250, .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 2 }, 6000, .default);
    try worksheet.writeNumber(.{ .row = 5, .col = 2 }, 500, .default);

    try worksheet.writeString(.{ .col = 3 }, "Value 2", bold);
    try worksheet.writeNumber(.{ .row = 1, .col = 3 }, 8000, .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 3 }, 3000, .default);
    try worksheet.writeNumber(.{ .row = 3, .col = 3 }, 1000, .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 3 }, 6000, .default);
    try worksheet.writeNumber(.{ .row = 5, .col = 3 }, 300, .default);

    try worksheet.writeString(.{ .col = 4 }, "Value 3", bold);
    try worksheet.writeNumber(.{ .row = 1, .col = 4 }, 6000, .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 4 }, 4000, .default);
    try worksheet.writeNumber(.{ .row = 3, .col = 4 }, 2000, .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 4 }, 6500, .default);
    try worksheet.writeNumber(.{ .row = 5, .col = 4 }, 200, .default);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
const WorkSheet = xwz.WorkSheet;
const Format = xwz.Format;
