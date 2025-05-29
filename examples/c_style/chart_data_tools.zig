//
// A demo of an various Excel chart data tools that are available via a
// libxlsxwriter chart.
//
// These include Drop Lines and High-Low Lines.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

// Write some data to the worksheet
fn writeWorksheetData(worksheet: *lxw.lxw_worksheet, bold: *lxw.lxw_format) void {
    const data = [_][3]u8{
        // Three columns of data
        [_]u8{ 2, 10, 30 },
        [_]u8{ 3, 40, 60 },
        [_]u8{ 4, 50, 70 },
        [_]u8{ 5, 20, 50 },
        [_]u8{ 6, 10, 40 },
        [_]u8{ 7, 50, 30 },
    };

    _ = lxw.worksheet_write_string(worksheet, 0, 0, "Number", bold);
    _ = lxw.worksheet_write_string(worksheet, 0, 1, "Batch 1", bold);
    _ = lxw.worksheet_write_string(worksheet, 0, 2, "Batch 2", bold);

    for (data, 0..) |row, row_idx| {
        for (row, 0..) |value, col_idx| {
            _ = lxw.worksheet_write_number(
                worksheet,
                @intCast(row_idx + 1),
                @intCast(col_idx),
                @floatFromInt(value),
                null,
            );
        }
    }
}

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-chart_data_tools.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Add a bold format to use to highlight the header cells
    const bold = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_bold(bold);

    // Write some data for the chart
    writeWorksheetData(worksheet, bold);

    // Chart 1. Example with High Low Lines
    var chart = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_LINE);

    // Add a chart title
    _ = lxw.chart_title_set_name(chart, "Chart with High-Low Lines");

    // Add the first series to the chart
    _ = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add high-low lines to the chart
    _ = lxw.chart_set_high_low_lines(chart, null);

    // Insert the chart into the worksheet
    _ = lxw.worksheet_insert_chart(worksheet, 1, 4, chart);

    // Chart 2. Example with Drop Lines
    chart = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_LINE);

    // Add a chart title
    _ = lxw.chart_title_set_name(chart, "Chart with Drop Lines");

    // Add the first series to the chart
    _ = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add drop lines to the chart
    _ = lxw.chart_set_drop_lines(chart, null);

    // Insert the chart into the worksheet
    _ = lxw.worksheet_insert_chart(worksheet, 17, 4, chart);

    // Chart 3. Example with Up-Down bars
    chart = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_LINE);

    // Add a chart title
    _ = lxw.chart_title_set_name(chart, "Chart with Up-Down bars");

    // Add the first series to the chart
    _ = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add Up-Down bars to the chart
    _ = lxw.chart_set_up_down_bars(chart);

    // Insert the chart into the worksheet
    _ = lxw.worksheet_insert_chart(worksheet, 33, 4, chart);

    // Chart 4. Example with Up-Down bars with formatting
    chart = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_LINE);

    // Add a chart title
    _ = lxw.chart_title_set_name(chart, "Chart with Up-Down bars");

    // Add the first series to the chart
    _ = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add Up-Down bars to the chart, with formatting
    var line = lxw.lxw_chart_line{ .color = lxw.LXW_COLOR_BLACK };
    var up_fill = lxw.lxw_chart_fill{ .color = 0x00B050 };
    var down_fill = lxw.lxw_chart_fill{ .color = lxw.LXW_COLOR_RED };

    _ = lxw.chart_set_up_down_bars_format(chart, &line, &up_fill, &line, &down_fill);

    // Insert the chart into the worksheet
    _ = lxw.worksheet_insert_chart(worksheet, 49, 4, chart);

    // Chart 5. Example with Markers and data labels
    chart = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_LINE);

    // Add a chart title
    _ = lxw.chart_title_set_name(chart, "Chart with Data Labels and Markers");

    // Add the first series to the chart
    const series = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add series markers
    _ = lxw.chart_series_set_marker_type(series, lxw.LXW_CHART_MARKER_CIRCLE);

    // Add series data labels
    _ = lxw.chart_series_set_labels(series);

    // Insert the chart into the worksheet
    _ = lxw.worksheet_insert_chart(worksheet, 65, 4, chart);

    // Chart 6. Example with Error Bars
    chart = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_LINE);

    // Add a chart title
    _ = lxw.chart_title_set_name(chart, "Chart with Error Bars");

    // Add the first series to the chart
    const series2 = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add error bars to show Standard Error
    _ = lxw.chart_series_set_error_bars(
        series2.*.y_error_bars,
        lxw.LXW_CHART_ERROR_BAR_TYPE_STD_ERROR,
        0,
    );

    // Add series data labels
    _ = lxw.chart_series_set_labels(series2);

    // Insert the chart into the worksheet
    _ = lxw.worksheet_insert_chart(worksheet, 81, 4, chart);

    // Chart 7. Example with a trendline
    chart = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_LINE);

    // Add a chart title
    _ = lxw.chart_title_set_name(chart, "Chart with a Trendline");

    // Add the first series to the chart
    const series3 = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");
    _ = lxw.chart_add_series(chart, "=Sheet1!$A$2:$A$7", "=Sheet1!$C$2:$C$7");

    // Add a polynomial trendline
    var poly_line = lxw.lxw_chart_line{
        .color = lxw.LXW_COLOR_GRAY,
        .dash_type = lxw.LXW_CHART_LINE_DASH_LONG_DASH,
    };

    _ = lxw.chart_series_set_trendline(series3, lxw.LXW_CHART_TRENDLINE_TYPE_POLY, 3);
    _ = lxw.chart_series_set_trendline_line(series3, &poly_line);

    // Insert the chart into the worksheet
    _ = lxw.worksheet_insert_chart(worksheet, 97, 4, chart);

    _ = lxw.workbook_close(workbook);
}
