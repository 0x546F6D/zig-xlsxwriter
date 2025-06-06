pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Write some data for the chart
    try worksheet.writeNumber(.{}, 10, .default);
    try worksheet.writeNumber(.{ .row = 1 }, 40, .default);
    try worksheet.writeNumber(.{ .row = 2 }, 50, .default);
    try worksheet.writeNumber(.{ .row = 3 }, 20, .default);
    try worksheet.writeNumber(.{ .row = 4 }, 10, .default);
    try worksheet.writeNumber(.{ .row = 5 }, 50, .default);

    // Create a chart object
    var chart = try workbook.addChart(.line);

    // Configure the chart
    _ = try chart.addSeries(null, "Sheet1!$A$1:$A$6");

    // Create some fonts to use in the chart
    const font1: xwz.ChartFont = .{ .name = "Calibri", .color = .blue };
    const font2: xwz.ChartFont = .{ .name = "Courrier", .color = @enumFromInt(0x92D050) };
    const font3: xwz.ChartFont = .{ .name = "Arial", .color = @enumFromInt(0x00B0F0) };
    const font4: xwz.ChartFont = .{ .name = "Century", .color = .red };
    const font5: xwz.ChartFont = .{ .rotation = -30 };
    const font6: xwz.ChartFont = .{
        .color = @enumFromInt(0x7030A0),
        .bold = true,
        .italic = true,
        .underline = true,
    };

    // Write the chart title with a font
    chart.titleSetName("Test Results");
    chart.titleSetNameFont(font1);

    // Write the Y axis with a font
    chart.y_axis.setName("Units");
    chart.y_axis.setNameFont(font2);
    chart.y_axis.setNumFont(font3);

    // Write the X axis with a font
    chart.x_axis.setName("Month");
    chart.x_axis.setNameFont(font4);
    chart.x_axis.setNumFont(font5);

    // Display the chart legend at the bottom of the chart
    chart.legendSetPosition(.bottom);
    chart.legendSetFont(font6);

    // Insert the chart into the worksheet
    try worksheet.insertChart(.{ .col = 2 }, chart, null);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
