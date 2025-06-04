pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    const asset_path = try h.getAssetPath(alloc, "logo.png");
    defer alloc.free(asset_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);
    const chartsheet = try workbook.addChartSheet(null);

    // Add a bold format for headers
    const bold = try workbook.addFormat();
    bold.setBold();

    // Write the data for the chart
    try writeWorksheetData(worksheet, bold);

    // Create a bar chart
    const chart = try workbook.addChart(.bar);

    // Add the first series to the chart
    var series = try chart.addSeries(
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    // Set the name for the first series
    series.setName("=Sheet1!$B$1");

    // Add a second series and configure it programmatically
    series = try chart.addSeries(null, null);

    series.setCategories(
        "Sheet1",
        .{ .first_row = 1, .last_row = 6 },
    );
    series.setValues(
        "Sheet1",
        .{ .first_row = 1, .first_col = 2, .last_row = 6, .last_col = 2 },
    );
    series.setNameRange("Sheet1", .{ .col = 2 });

    // Add chart title and axis labels
    chart.titleSetName("Results of sample analysis");
    chart.x_axis.setName("Test number");
    chart.y_axis.setName("Sample length (mm)");

    // Set chart style
    chart.setStyle(11);

    // Add the chart to the chartsheet
    try chartsheet.setChart(chart);

    // Make the chartsheet active
    chartsheet.activate();
}

fn writeWorksheetData(worksheet: WorkSheet, bold: Format) !void {
    const data = [_][3]u8{
        .{ 2, 10, 30 },
        .{ 3, 40, 60 },
        .{ 4, 50, 70 },
        .{ 5, 20, 50 },
        .{ 6, 10, 40 },
        .{ 7, 50, 30 },
    };

    try worksheet.writeString(.{}, "Number", bold);
    try worksheet.writeString(.{ .col = 1 }, "Batch 1", bold);
    try worksheet.writeString(.{ .col = 2 }, "Batch 2", bold);

    for (data, 0..) |row, row_idx| {
        for (row, 0..) |value, col_idx| {
            try worksheet.writeNumber(
                .{
                    .row = @intCast(row_idx + 1),
                    .col = @intCast(col_idx),
                },
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
