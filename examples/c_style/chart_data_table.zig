//
// An example of creating Excel column charts with data tables using the
// libxlsxwriter library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

fn writeWorksheetData(worksheet: ?*lxw.lxw_worksheet, bold: ?*lxw.lxw_format) void {
    const data = [_][3]u8{
        .{ 2, 10, 30 },
        .{ 3, 40, 60 },
        .{ 4, 50, 70 },
        .{ 5, 20, 50 },
        .{ 6, 10, 40 },
        .{ 7, 50, 30 },
    };

    _ = lxw.worksheet_write_string(worksheet, 0, 0, "Number", bold);
    _ = lxw.worksheet_write_string(worksheet, 0, 1, "Batch 1", bold);
    _ = lxw.worksheet_write_string(worksheet, 0, 2, "Batch 2", bold);

    for (data, 0..) |row, i| {
        for (row, 0..) |val, j| {
            _ = lxw.worksheet_write_number(worksheet, @as(u32, @intCast(i + 1)), @as(u16, @intCast(j)), @as(f64, @floatFromInt(val)), null);
        }
    }
}

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-chart_data_table.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Add bold format
    const bold = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_bold(bold);

    // Write data
    writeWorksheetData(worksheet, bold);

    // Chart 1: Column chart with data table
    const chart1 = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_COLUMN);
    var series = lxw.chart_add_series(chart1, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = lxw.chart_series_set_name(series, "=Sheet1!$B$1");

    series = lxw.chart_add_series(chart1, null, null);
    _ = lxw.chart_series_set_categories(series, "Sheet1", 1, 0, 6, 0);
    _ = lxw.chart_series_set_values(series, "Sheet1", 1, 2, 6, 2);
    _ = lxw.chart_series_set_name_range(series, "Sheet1", 0, 2);

    _ = lxw.chart_title_set_name(chart1, "Chart with Data Table");
    _ = lxw.chart_axis_set_name(chart1.*.x_axis, "Test number");
    _ = lxw.chart_axis_set_name(chart1.*.y_axis, "Sample length (mm)");

    _ = lxw.chart_set_table(chart1);
    _ = lxw.worksheet_insert_chart(worksheet, 1, 4, chart1);

    // Chart 2: Column chart with data table and legend keys
    const chart2 = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_COLUMN);

    series = lxw.chart_add_series(chart2, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = lxw.chart_series_set_name(series, "=Sheet1!$B$1");

    series = lxw.chart_add_series(chart2, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    _ = lxw.chart_series_set_name(series, "=Sheet1!$C$1");

    _ = lxw.chart_title_set_name(chart2, "Data Table with legend keys");
    _ = lxw.chart_axis_set_name(chart2.*.x_axis, "Test number");
    _ = lxw.chart_axis_set_name(chart2.*.y_axis, "Sample length (mm)");

    _ = lxw.chart_set_table(chart2);
    _ = lxw.chart_set_table_grid(chart2, 1, 1, 1, 1);
    _ = lxw.chart_legend_set_position(chart2, lxw.LXW_CHART_LEGEND_NONE);

    _ = lxw.worksheet_insert_chart(worksheet, 17, 4, chart2);

    _ = lxw.workbook_close(workbook);
}
