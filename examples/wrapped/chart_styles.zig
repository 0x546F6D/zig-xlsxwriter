pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};

    // Define chart types and names
    const chart_types = [_]xwz.ChartType{
        .column,
        .area,
        .line,
        .pie,
    };
    const chart_names = [_][:0]const u8{ "Column", "Area", "Line", "Pie" };

    // Create a worksheet for each chart type
    for (chart_types, chart_names) |chart_type, chart_name| {
        const worksheet = try workbook.addWorkSheet(chart_name);
        worksheet.setZoom(30);

        // Create 48 charts, each with a different style
        var style_num: u8 = 1;
        var row_num: usize = 0;
        while (row_num < 90) : (row_num += 15) {
            var col_num: usize = 0;
            while (col_num < 64) : (col_num += 8) {
                const chart = try workbook.addChart(chart_type);

                // Create chart title with style number
                var title_buf: [32]u8 = undefined;
                const title = std.fmt.bufPrintZ(&title_buf, "Style {d}", .{style_num}) catch unreachable;

                _ = try chart.addSeries(null, "=Data!$A$1:$A$6");
                chart.titleSetName(title);
                chart.setStyle(style_num);

                try worksheet.insertChart(.{ .row = @intCast(row_num), .col = @intCast(col_num) }, chart, null);

                style_num += 1;
            }
        }
    }

    // Create a worksheet with data for the charts
    const data_worksheet = try workbook.addWorkSheet("Data");
    try data_worksheet.writeNumber(.{}, 10, .default);
    try data_worksheet.writeNumber(.{ .row = 1 }, 40, .default);
    try data_worksheet.writeNumber(.{ .row = 2 }, 50, .default);
    try data_worksheet.writeNumber(.{ .row = 3 }, 20, .default);
    try data_worksheet.writeNumber(.{ .row = 4 }, 10, .default);
    try data_worksheet.writeNumber(.{ .row = 5 }, 50, .default);
}

const std = @import("std");
var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
