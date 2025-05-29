//
// An example of creating an Excel pie chart with user defined colors using
// the libxlsxwriter library.
//
// In general formatting is applied to an entire series in a chart. However,
// it is occasionally required to format individual points in a series. In
// particular this is required for Pie/Doughnut charts where each segment is
// represented by a point.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-chart_pie_colors.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Write data for the chart
    _ = lxw.worksheet_write_string(worksheet, 0, 0, "Pass", null);
    _ = lxw.worksheet_write_string(worksheet, 1, 0, "Fail", null);
    _ = lxw.worksheet_write_number(worksheet, 0, 1, 90, null);
    _ = lxw.worksheet_write_number(worksheet, 1, 1, 10, null);

    // Create a pie chart
    const chart = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_PIE);

    // Add the data series
    const series = lxw.chart_add_series(
        chart,
        "=Sheet1!$A$1:$A$2",
        "=Sheet1!$B$1:$B$2",
    );

    // Create fills for chart segments
    var red_fill = lxw.lxw_chart_fill{
        .color = lxw.LXW_COLOR_RED,
    };
    var green_fill = lxw.lxw_chart_fill{
        .color = lxw.LXW_COLOR_GREEN,
    };

    // Create points with fills
    var red_point = lxw.lxw_chart_point{
        .fill = &red_fill,
    };
    var green_point = lxw.lxw_chart_point{
        .fill = &green_fill,
    };

    // Create array of points (null terminated)
    var points = [_]?*lxw.lxw_chart_point{
        &green_point,
        &red_point,
        null,
    };

    // Set the points on the series
    _ = lxw.chart_series_set_points(series, &points);

    // Insert chart into worksheet
    _ = lxw.worksheet_insert_chart(worksheet, 1, 3, chart);

    _ = lxw.workbook_close(workbook);
}
