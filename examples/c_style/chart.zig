const lxw = @import("lxw");

pub fn main() void {

    // Create a new workbook and add a worksheet.
    const workbook: ?*lxw.lxw_workbook = lxw.workbook_new("zig-chart.xlsx");
    const worksheet: ?*lxw.lxw_worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Write some data for the chart.
    write_worksheet_data(worksheet);

    // Create a chart object.
    const chart: ?*lxw.lxw_chart = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_COLUMN);

    // Configure the chart. In simplest case we just add some value data
    // series. The null categories will default to 1 to 5 like in Excel.

    _ = lxw.chart_add_series(chart, null, "Sheet1!$A$1:$A$5");
    _ = lxw.chart_add_series(chart, null, "Sheet1!$B$1:$B$5");
    _ = lxw.chart_add_series(chart, null, "Sheet1!$C$1:$C$5");

    var font: lxw.lxw_chart_font = .{
        .name = @constCast("Chart example"),
        .bold = lxw.LXW_EXPLICIT_FALSE,
        .color = lxw.LXW_COLOR_BLUE,
        .italic = lxw.LXW_EXPLICIT_FALSE,
        .size = 16,
        .rotation = 0,
        .underline = 0,
        .charset = 0,
        .pitch_family = 0,
        .baseline = 0,
    };

    _ = lxw.chart_title_set_name(chart, "Year End Results");
    _ = lxw.chart_title_set_name_font(chart, &font);

    // Insert the chart into the worksheet.
    // CELL("B7") expands to row, col
    const row_b7 = lxw.lxw_name_to_row("B7");
    const col_b7 = lxw.lxw_name_to_col("B7");
    _ = lxw.worksheet_insert_chart(worksheet, row_b7, col_b7, chart);

    _ = lxw.workbook_close(workbook);
}

fn write_worksheet_data(worksheet: ?*lxw.lxw_worksheet) void {
    const data: [5][3]f64 = .{
        [3]f64{ 1, 2, 3 },
        [3]f64{ 2, 4, 6 },
        [3]f64{ 3, 6, 9 },
        [3]f64{ 4, 8, 12 },
        [3]f64{ 5, 10, 15 },
    };

    for (data, 0..) |row_data, row| {
        for (row_data, 0..) |value, col| {
            _ = lxw.worksheet_write_number(worksheet, @intCast(row), @intCast(col), value, null);
        }
    }
}
