//
// An example of creating an Excel chartsheet using the libxlsxwriter library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

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

    for (data, 0..) |row, row_idx| {
        for (row, 0..) |value, col_idx| {
            _ = lxw.worksheet_write_number(
                worksheet,
                @as(u32, @intCast(row_idx + 1)),
                @as(u16, @intCast(col_idx)),
                @as(f64, @floatFromInt(value)),
                null,
            );
        }
    }
}

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-chartsheet.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);
    const chartsheet = lxw.workbook_add_chartsheet(workbook, null);

    // Add a bold format for headers
    const bold = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_bold(bold);

    // Write the data for the chart
    writeWorksheetData(worksheet, bold);

    // Create a bar chart
    const chart = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_BAR);

    // Add the first series to the chart
    var series = lxw.chart_add_series(
        chart,
        "=Sheet1!$A$2:$A$7",
        "=Sheet1!$B$2:$B$7",
    );

    // Set the name for the first series
    _ = lxw.chart_series_set_name(series, "=Sheet1!$B$1");

    // Add a second series and configure it programmatically
    series = lxw.chart_add_series(chart, null, null);

    _ = lxw.chart_series_set_categories(series, "Sheet1", 1, 0, 6, 0);
    _ = lxw.chart_series_set_values(series, "Sheet1", 1, 2, 6, 2);
    _ = lxw.chart_series_set_name_range(series, "Sheet1", 0, 2);

    // Add chart title and axis labels
    _ = lxw.chart_title_set_name(chart, "Results of sample analysis");
    _ = lxw.chart_axis_set_name(chart.*.x_axis, "Test number");
    _ = lxw.chart_axis_set_name(chart.*.y_axis, "Sample length (mm)");

    // Set chart style
    _ = lxw.chart_set_style(chart, 11);

    // Add the chart to the chartsheet
    _ = lxw.chartsheet_set_chart(chartsheet, chart);

    // Make the chartsheet active
    _ = lxw.chartsheet_activate(chartsheet);

    _ = lxw.workbook_close(workbook);
}
