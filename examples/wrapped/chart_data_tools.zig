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

    // Chart 1. Example with High Low Lines
    var chart = try workbook.addChart(.line);

    // Add a chart title
    chart.titleSetName("Chart with High-Low Lines");

    // Add the first series to the chart
    _ = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add high-low lines to the chart
    chart.setHighLowLines(null);

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .row = 1, .col = 4 }, chart, null);

    // Chart 2. Example with Drop Lines
    chart = try workbook.addChart(.line);

    // Add a chart title
    chart.titleSetName("Chart with Drop Lines");

    // Add the first series to the chart
    _ = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add drop lines to the chart
    chart.setDropLines(null);

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .row = 17, .col = 4 }, chart, null);

    // Chart 3. Example with Up-Down bars
    chart = try workbook.addChart(.line);

    // Add a chart title
    chart.titleSetName("Chart with Up-Down bars");

    // Add the first series to the chart
    _ = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add Up-Down bars to the chart
    chart.setUpDownBars();

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .row = 33, .col = 4 }, chart, null);

    // Chart 4. Example with Up-Down bars with formatting
    chart = try workbook.addChart(.line);

    // Add a chart title
    chart.titleSetName("Chart with Up-Down bars");

    // Add the first series to the chart
    _ = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add Up-Down bars to the chart, with formatting
    const line = xwz.ChartLine{ .color = .black };
    const up_fill = xwz.ChartFill{ .color = @enumFromInt(0x00B050) };
    const down_fill = xwz.ChartFill{ .color = .red };

    chart.setUpDownBarsFormat(line, up_fill, line, down_fill);

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .row = 49, .col = 4 }, chart, null);

    // Chart 5. Example with Markers and data labels
    chart = try workbook.addChart(.line);

    // Add a chart title
    chart.titleSetName("Chart with Data Labels and Markers");

    // Add the first series to the chart
    var series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add series markers
    series.setMarkerType(.circle);

    // Add series data labels
    series.setLabels();

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .row = 65, .col = 4 }, chart, null);

    // Chart 6. Example with Error Bars
    chart = try workbook.addChart(.line);

    // Add a chart title
    chart.titleSetName("Chart with Error Bars");

    // Add the first series to the chart
    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add error bars to show Standard Error
    series.y_error_bars.setType(.std_error, 0);

    // Add series data labels
    series.setLabels();

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .row = 81, .col = 4 }, chart, null);

    // Chart 7. Example with a trendline
    chart = try workbook.addChart(.line);

    // Add a chart title
    chart.titleSetName("Chart with a Trendline");

    // Add the first series to the chart
    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add a polynomial trendline
    const poly_line = xwz.ChartLine{
        .color = .gray,
        .dash_type = .long_dash,
    };

    series.setTrendline(.poly, 3);
    series.setTrendlineLine(poly_line);

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .row = 97, .col = 4 }, chart, null);
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
