pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xlsxwriter.initWorkBook(xlsx_path.ptr);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Write some data for the chart. */
    try write_worksheet_data(worksheet);

    // Create a chart object. */
    const chart = try workbook.addChart(ChartType.column);

    // Configure the chart. In simplest case we just add some value data
    // series. The null categories will default to 1 to 5 like in Excel.
    _ = try chart.addSeries(null, "Sheet1!$A$1:$A$5");
    _ = try chart.addSeries(null, "Sheet1!$B$1:$B$5");
    _ = try chart.addSeries(null, "Sheet1!$C$1:$C$5");

    var font: ChartFont = .{
        .name = @constCast("Chart example"),
        .bold = xlsxwriter.explicit_false,
        .color = @intFromEnum(xlsxwriter.DefinedColor.blue),
        .italic = xlsxwriter.explicit_false,
        .size = 16,
        .rotation = 0,
        .underline = 0,
        .charset = 0,
        .pitch_family = 0,
        .baseline = 0,
    };

    chart.titleSetName("Year End Results", &font);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(xlsxwriter.nameToRow("B7"), xlsxwriter.nameToCol("B7"), chart);
}

fn write_worksheet_data(worksheet: xlsxwriter.WorkSheet) !void {
    const data: [5][3]f64 = .{
        [3]f64{ 1, 2, 3 },
        [3]f64{ 2, 4, 6 },
        [3]f64{ 3, 6, 9 },
        [3]f64{ 4, 8, 12 },
        [3]f64{ 5, 10, 15 },
    };

    for (data, 0..) |set, row| {
        for (set, 0..) |cell, col| {
            try worksheet.writeNumber(@intCast(row), @intCast(col), cell, .none);
        }
    }
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");
const ChartType = xlsxwriter.ChartType;
const ChartFont = xlsxwriter.ChartFont;
