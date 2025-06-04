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

    // Write some data for the chart
    try writeWorksheetData(worksheet, bold);

    // Chart 1: Create a bar chart
    var chart = try workbook.addChart(.bar);

    // Add the first series to the chart
    var series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    // Set the name for the series instead of the default "Series 1"
    series.setName("=Sheet1!$B$1");

    // Add a second series but leave the categories and values undefined
    series = try chart.addSeries(null, null);

    // Configure the series using a syntax that is easier to define programmatically
    series.setCategories("Sheet1", .{ .first_row = 1, .last_row = 6 }); // "=Sheet1!$A$2:$A$7"
    series.setValues("Sheet1", .{ .first_row = 1, .first_col = 2, .last_row = 6, .last_col = 2 }); // "=Sheet1!$C$2:$C$7"
    series.setNameRange("Sheet1", .{ .col = 2 }); // "=Sheet1!$C$1"

    // Add a chart title and some axis labels
    chart.titleSetName("Results of sample analysis");
    chart.x_axis.setName("Test number");
    chart.y_axis.setName("Sample length (mm)");

    // Set an Excel chart style
    chart.setStyle(11);

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .row = 1, .col = 4 }, chart, null);

    // Chart 2: Create a stacked bar chart
    chart = try workbook.addChart(.bar_stacked);

    // Add the first series to the chart
    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    // Set the name for the series instead of the default "Series 1"
    series.setName("=Sheet1!$B$1");

    // Add the second series to the chart
    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    // Set the name for the series instead of the default "Series 2"
    series.setName("=Sheet1!$C$1");

    // Add a chart title and some axis labels
    chart.titleSetName("Results of sample analysis");
    chart.x_axis.setName("Test number");
    chart.y_axis.setName("Sample length (mm)");

    // Set an Excel chart style
    chart.setStyle(12);

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .row = 17, .col = 4 }, chart, null);

    // Chart 3: Create a percent stacked bar chart
    chart = try workbook.addChart(.bar_stacked_percent);

    // Add the first series to the chart
    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    // Set the name for the series instead of the default "Series 1"
    series.setName("=Sheet1!$B$1");

    // Add the second series to the chart
    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    // Set the name for the series instead of the default "Series 2"
    series.setName("=Sheet1!$C$1");

    // Add a chart title and some axis labels
    chart.titleSetName("Results of sample analysis");
    chart.x_axis.setName("Test number");
    chart.y_axis.setName("Sample length (mm)");
    // Set an Excel chart style
    chart.setStyle(13);

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .row = 33, .col = 4 }, chart, null);
}

// Write some data to the worksheet
fn writeWorksheetData(worksheet: WorkSheet, bold: Format) !void {
    // Three columns of data
    const data = [_][3]u8{
        [_]u8{ 2, 10, 30 },
        [_]u8{ 3, 40, 60 },
        [_]u8{ 4, 50, 70 },
        [_]u8{ 5, 20, 50 },
        [_]u8{ 6, 10, 40 },
        [_]u8{ 7, 50, 30 },
    };

    try worksheet.writeString(.{}, "Number", bold);
    try worksheet.writeString(.{ .col = 1 }, "Batch 1", bold);
    try worksheet.writeString(.{ .col = 2 }, "Batch 2", bold);

    for (data, 0..) |row, row_idx| {
        for (row, 0..) |value, col_idx| {
            try worksheet.writeNumber(
                .{ .row = @intCast(row_idx + 1), .col = @intCast(col_idx) },
                @floatFromInt(value),
                .default,
            );
        }
    }
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
const WorkSheet = xwz.WorkSheet;
const Format = xwz.Format;
