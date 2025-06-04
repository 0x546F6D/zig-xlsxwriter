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

    // Chart 1: Create an area chart
    const chart1 = try workbook.addChart(.area);

    // Add the first series to the chart
    const series1 = try chart1.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    // Set the name for the series instead of the default "Series 1"
    series1.setName("=Sheet1!$B$1");

    // Add a second series but leave the categories and values undefined
    const series2 = try chart1.addSeries(null, null);
    // Configure the series using a syntax that is easier to define programmatically
    series2.setCategories("Sheet1", .{ .first_row = 1, .last_row = 6 }); // "=Sheet1!$A$2:$A$7"
    series2.setValues("Sheet1", .{ .first_row = 1, .first_col = 2, .last_row = 6, .last_col = 2 }); // "=Sheet1!$C$2:$C$7"
    series2.setNameRange("Sheet1", .{ .col = 2 }); // "=Sheet1!$C$1"

    // Add a chart title and some axis labels
    chart1.titleSetName("Results of sample analysis");
    chart1.x_axis.setName("Test number");
    chart1.y_axis.setName("Sample length (mm)");

    // Set an Excel chart style
    chart1.setStyle(11);

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .row = 1, .col = 4 }, chart1, null);

    // Chart 2: Create a stacked area chart
    const chart2 = try workbook.addChart(.area_stacked);

    // Add the first series to the chart
    const series3 = try chart2.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    // Set the name for the series instead of the default "Series 1"
    series3.setName("=Sheet1!$B$1");

    // Add the second series to the chart
    const series4 = try chart2.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    // Set the name for the series instead of the default "Series 2"
    series4.setName("=Sheet1!$C$1");

    // Add a chart title and some axis labels
    chart2.titleSetName("Results of sample analysis");
    chart2.x_axis.setName("Test number");
    chart2.y_axis.setName("Sample length (mm)");

    // Set an Excel chart style
    chart2.setStyle(12);

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .row = 17, .col = 4 }, chart2, null);

    // Chart 3: Create a percent stacked area chart
    const chart3 = try workbook.addChart(.area_stacked_percent);

    // Add the first series to the chart
    const series5 = try chart3.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    // Set the name for the series instead of the default "Series 1"
    series5.setName("=Sheet1!$B$1");

    // Add the second series to the chart
    const series6 = try chart3.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    // Set the name for the series instead of the default "Series 2"
    series6.setName("=Sheet1!$C$1");

    // Add a chart title and some axis labels
    chart3.titleSetName("Results of sample analysis");
    chart3.x_axis.setName("Test number");
    chart3.y_axis.setName("Sample length (mm)");

    // Set an Excel chart style
    chart3.setStyle(13);

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .row = 33, .col = 4 }, chart3, null);
}

// Write some data to the worksheet
fn writeWorksheetData(worksheet: WorkSheet, bold: Format) !void {
    const data = [_][3]u8{
        // Three columns of data
        [_]u8{ 2, 40, 30 },
        [_]u8{ 3, 40, 25 },
        [_]u8{ 4, 50, 30 },
        [_]u8{ 5, 30, 10 },
        [_]u8{ 6, 25, 5 },
        [_]u8{ 7, 50, 10 },
    };

    try worksheet.writeString(.{}, "Number", bold);
    try worksheet.writeString(.{ .col = 1 }, "Batch 1", bold);
    try worksheet.writeString(.{ .col = 2 }, "Batch 2", bold);

    for (data, 0..) |row_data, row_idx| {
        for (row_data, 0..) |cell_value, col_idx| {
            try worksheet.writeNumber(
                .{ .row = @intCast(row_idx + 1), .col = @intCast(col_idx) },
                @floatFromInt(cell_value),
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
