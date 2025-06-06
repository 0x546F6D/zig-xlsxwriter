pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(alloc, xlsx_path, null);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Add a bold format to use to highlight the header cells
    const bold = try workbook.addFormat();
    bold.setBold();

    // Write some data for the chart.
    try writeWorksheetData(worksheet, bold);

    // Chart 1: Create a simple doughnut chart
    var chart = try workbook.addChart(.doughnut);

    // Add the first series to the chart.
    var series = try chart.addSeries("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

    // Set the name for the series instead of the default "Series 1".
    series.setName("Doughnut sales data");

    // Add a chart title.
    chart.titleSetName("Popular Doughnut Types");

    // Set an Excel chart style.
    chart.setStyle(10);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 1, .col = 3 }, chart, null);

    // Chart 2: Create a doughnut chart with user defined segment colors
    chart = try workbook.addChart(.doughnut);

    // Add the first series to the chart.
    series = try chart.addSeries("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

    // Set the name for the series instead of the default "Series 1".
    series.setName("Doughnut sales data");

    // Add a chart title.
    chart.titleSetName("Doughnut Chart with user defined colors");

    // Add some points with fills.
    const point1: Point = .{ .fill = &.{ .color = @enumFromInt(0xFA58D0) } };
    const point2: Point = .{ .fill = &.{ .color = @enumFromInt(0x61210B) } };
    const point3: Point = .{ .fill = &.{ .color = @enumFromInt(0xF5F6CE) } };

    // Create an array of the point objects.
    const points_array = &.{
        point1,
        point2,
        point3,
    };

    // Add/override the points/segments of the chart.
    try series.setPoints(points_array);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 17, .col = 3 }, chart, null);

    // chart = xlsxwriter.workbook_add_chart(workbook, xlsxwriter.LXW_CHART_DOUGHNUT);
    //
    // // Add the first series to the chart
    // series = xlsxwriter.chart_add_series(chart, "=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");
    //
    // // Set the name for the series instead of the default "Series 1"
    // _ = xlsxwriter.chart_series_set_name(series, "Doughnut sales data");
    //
    // // Add a chart title
    // _ = xlsxwriter.chart_title_set_name(chart, "Doughnut Chart with user defined colors");
    //
    // // Add fills for use in the chart
    // var fill1 = xlsxwriter.lxw_chart_fill{ .color = 0xFA58D0 };
    // var fill2 = xlsxwriter.lxw_chart_fill{ .color = 0x61210B };
    // var fill3 = xlsxwriter.lxw_chart_fill{ .color = 0xF5F6CE };
    //
    // // Add some points with the above fills
    // var point1 = xlsxwriter.lxw_chart_point{ .fill = &fill1 };
    // var point2 = xlsxwriter.lxw_chart_point{ .fill = &fill2 };
    // var point3 = xlsxwriter.lxw_chart_point{ .fill = &fill3 };
    //
    // // Create an array of the point objects
    // const points_array = [_]*xlsxwriter.lxw_chart_point{
    //     &point1,
    //     &point2,
    //     &point3,
    // };
    //
    // // Create a null-terminated array of pointers as required by the C API
    // var points_ptrs: [4][*c]xlsxwriter.lxw_chart_point = undefined;
    // for (points_array, 0..) |point, i| {
    //     points_ptrs[i] = point;
    // }
    // points_ptrs[3] = null;
    //
    // // Add/override the points/segments of the chart
    // _ = xlsxwriter.chart_series_set_points(series, @ptrCast(&points_ptrs));
    //
    // // Insert the chart into the worksheet
    // _ = xlsxwriter.worksheet_insert_chart(worksheet, 17, 3, chart);

    // Chart 3: Create a Doughnut chart with rotation of the segments
    chart = try workbook.addChart(.doughnut);

    // Add the first series to the chart.
    series = try chart.addSeries("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

    // Set the name for the series instead of the default "Series 1".
    series.setName("Doughnut sales data");

    // Add a chart title.
    chart.titleSetName("Doughnut Chart with segment rotation");

    // Change the angle/rotation of the first segment.
    chart.setRotation(90);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 33, .col = 3 }, chart, null);

    // Chart 4: Create a Doughnut chart with user defined hole size and other options
    chart = try workbook.addChart(.doughnut);

    // Add the first series to the chart.
    series = try chart.addSeries("=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

    // Set the name for the series instead of the default "Series 1".
    series.setName("Doughnut sales data");

    // Add a chart title.
    chart.titleSetName("Doughnut Chart with options applied.");

    // Add/override the points/segments defined in Chart 2
    try series.setPoints(points_array);

    // Set an Excel chart style
    chart.setStyle(26);

    // Change the angle/rotation of the first segment.
    chart.setRotation(28);

    // Change the hole size
    chart.setHoleSize(33);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 49, .col = 3 }, chart, null);
}

// Write some data to the worksheet
fn writeWorksheetData(worksheet: WorkSheet, bold: Format) !void {
    try worksheet.writeString(.{}, "Category", bold);
    try worksheet.writeString(.{ .row = 1 }, "Glazed", .default);
    try worksheet.writeString(.{ .row = 2 }, "Chacolate", .default);
    try worksheet.writeString(.{ .row = 3 }, "Cream", .default);

    try worksheet.writeString(.{ .col = 1 }, "Values", bold);
    try worksheet.writeNumber(.{ .row = 1, .col = 1 }, 50, .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 1 }, 35, .default);
    try worksheet.writeNumber(.{ .row = 3, .col = 1 }, 15, .default);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
const WorkSheet = xwz.WorkSheet;
const Format = xwz.Format;
const Point = xwz.ChartPoint;
