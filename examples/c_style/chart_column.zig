//
// An example of creating Excel column charts using the libxlsxwriter library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");
const Cell = lxw.Cell;

fn writeWorksheetData(worksheet: *lxw.lxw_worksheet, bold: *lxw.lxw_format) void {
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

    for (data, 0..) |row, row_num| {
        for (row, 0..) |value, col_num| {
            _ = lxw.worksheet_write_number(worksheet, @intCast(row_num + 1), @intCast(col_num), @floatFromInt(value), null);
        }
    }
}

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-chart_column.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);
    var series: *lxw.lxw_chart_series = undefined;

    // Add a bold format to use to highlight the header cells.
    const bold = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_bold(bold);

    // Write some data for the chart.
    writeWorksheetData(worksheet, bold);

    // Chart 1: Create a column chart.
    var chart = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_COLUMN);

    // Add the first series to the chart.
    series = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Set the name for the series instead of the default "Series 1".
    _ = lxw.chart_series_set_name(series, "=Sheet1!$B$1");

    // Add a second series but leave the categories and values undefined.
    series = lxw.chart_add_series(chart, null, null);

    // Configure the series using a syntax that is easier to define programmatically.
    _ = lxw.chart_series_set_categories(series, "Sheet1", 1, 0, 6, 0); // "=Sheet1!$A$2:$A$7"
    _ = lxw.chart_series_set_values(series, "Sheet1", 1, 2, 6, 2); // "=Sheet1!$C$2:$C$7"
    _ = lxw.chart_series_set_name_range(series, "Sheet1", 0, 2); // "=Sheet1!$C$1"

    // Add a chart title and some axis labels.
    _ = lxw.chart_title_set_name(chart, "Results of sample analysis");
    _ = lxw.chart_axis_set_name(chart.*.x_axis, "Test number");
    _ = lxw.chart_axis_set_name(chart.*.y_axis, "Sample length (mm)");

    // Set an Excel chart style.
    _ = lxw.chart_set_style(chart, 11);

    // Insert the chart into the worksheet.
    _ = lxw.worksheet_insert_chart(worksheet, 1, 4, chart);

    // Chart 2: Create a stacked column chart.
    chart = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_COLUMN_STACKED);

    // Add the first series to the chart.
    series = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = lxw.chart_series_set_name(series, "=Sheet1!$B$1");

    // Add the second series to the chart.
    series = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    _ = lxw.chart_series_set_name(series, "=Sheet1!$C$1");

    // Add a chart title and some axis labels.
    _ = lxw.chart_title_set_name(chart, "Results of sample analysis");
    _ = lxw.chart_axis_set_name(chart.*.x_axis, "Test number");
    _ = lxw.chart_axis_set_name(chart.*.y_axis, "Sample length (mm)");

    // Set an Excel chart style.
    _ = lxw.chart_set_style(chart, 12);

    // Insert the chart into the worksheet.
    _ = lxw.worksheet_insert_chart(worksheet, 17, 4, chart);

    // Chart 3: Create a percent stacked column chart.
    chart = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_COLUMN_STACKED_PERCENT);

    // Add the first series to the chart.
    series = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = lxw.chart_series_set_name(series, "=Sheet1!$B$1");

    // Add the second series to the chart.
    series = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");
    _ = lxw.chart_series_set_name(series, "=Sheet1!$C$1");

    // Add a chart title and some axis labels.
    _ = lxw.chart_title_set_name(chart, "Results of sample analysis");
    _ = lxw.chart_axis_set_name(chart.*.x_axis, "Test number");
    _ = lxw.chart_axis_set_name(chart.*.y_axis, "Sample length (mm)");

    // Set an Excel chart style.
    _ = lxw.chart_set_style(chart, 13);

    // Insert the chart into the worksheet.
    _ = lxw.worksheet_insert_chart(worksheet, 33, 4, chart);

    _ = lxw.workbook_close(workbook);
}
