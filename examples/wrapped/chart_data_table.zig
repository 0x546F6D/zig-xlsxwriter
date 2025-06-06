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

    // Chart 1: Column chart with data table
    var chart = try workbook.addChart(.column);
    var series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    series.setName("=Sheet1!$B$1");

    series = try chart.addSeries(null, null);
    series.setCategories("Sheet1", .{ .first_row = 1, .last_row = 6 });
    series.setValues("Sheet1", .{ .first_row = 1, .first_col = 2, .last_row = 6, .last_col = 2 });
    series.setNameRange("Sheet1", .{ .col = 2 });

    chart.titleSetName("Chart with Data Table");
    chart.x_axis.setName("Test number");
    chart.y_axis.setName("Sample length (mm)");

    chart.setTable();
    try worksheet.insertChart(.{ .row = 1, .col = 4 }, chart, null);

    // Chart 2: Column chart with data table and legend keys
    chart = try workbook.addChart(.column);
    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    series.setName("=Sheet1!$B$1");

    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    series.setName("=Sheet1!$C$1");

    chart.titleSetName("Chart with legend keys");
    chart.x_axis.setName("Test number");
    chart.y_axis.setName("Sample length (mm)");

    chart.setTable();
    chart.setTableGrid(true, true, true, true);
    chart.legendSetPosition(.none);

    try worksheet.insertChart(.{ .row = 17, .col = 4 }, chart, null);
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

    for (data, 0..) |row, i| {
        for (row, 0..) |val, j| {
            try worksheet.writeNumber(
                .{ .row = @intCast(i + 1), .col = @intCast(j) },
                @floatFromInt(val),
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
