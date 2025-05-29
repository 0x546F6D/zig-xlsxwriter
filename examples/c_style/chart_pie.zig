//
// An example of creating an Excel pie chart using the libxlsxwriter library.
//
// The demo also shows how to set segment colors. It is possible to define
// chart colors for most types of libxlsxwriter charts via the series
// formatting functions. However, Pie/Doughnut charts are a special case since
// each segment is represented as a point so it is necessary to assign
// formatting to each point in the series.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

// Write some data to the worksheet.
fn write_worksheet_data(worksheet: *lxw.lxw_worksheet, bold: *lxw.lxw_format) void {
    _ = lxw.worksheet_write_string(worksheet, 0, 0, "Category", bold);
    _ = lxw.worksheet_write_string(worksheet, 1, 0, "Apple", null);
    _ = lxw.worksheet_write_string(worksheet, 2, 0, "Cherry", null);
    _ = lxw.worksheet_write_string(worksheet, 3, 0, "Pecan", null);

    _ = lxw.worksheet_write_string(worksheet, 0, 1, "Values", bold);
    _ = lxw.worksheet_write_number(worksheet, 1, 1, 60, null);
    _ = lxw.worksheet_write_number(worksheet, 2, 1, 30, null);
    _ = lxw.worksheet_write_number(worksheet, 3, 1, 10, null);
}

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-chart_pie.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Add a bold format to use to highlight the header cells.
    const bold = lxw.workbook_add_format(workbook);
    lxw.format_set_bold(bold);

    // Write some data for the chart.
    write_worksheet_data(worksheet, bold);

    // Chart 1: Create a simple pie chart.
    const chart1 = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_PIE);

    // Add the first series to the chart.
    const series1 = lxw.chart_add_series(chart1, "=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

    // Set the name for the series instead of the default "Series 1".
    lxw.chart_series_set_name(series1, "Pie sales data");

    // Add a chart title.
    lxw.chart_title_set_name(chart1, "Popular Pie Types");

    // Set an Excel chart style.
    lxw.chart_set_style(chart1, 10);

    // Insert the chart into the worksheet.
    _ = lxw.worksheet_insert_chart(worksheet, 1, 3, chart1);

    // Chart 2: Create a pie chart with user defined segment colors.
    const chart2 = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_PIE);

    // Add the first series to the chart.
    const series2 = lxw.chart_add_series(chart2, "=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

    // Set the name for the series instead of the default "Series 1".
    lxw.chart_series_set_name(series2, "Pie sales data");

    // Add a chart title.
    lxw.chart_title_set_name(chart2, "Pie Chart with user defined colors");

    // Add fills for use in the chart.
    var fill1 = lxw.lxw_chart_fill{ .color = 0x5ABA10 };
    var fill2 = lxw.lxw_chart_fill{ .color = 0xFE110E };
    var fill3 = lxw.lxw_chart_fill{ .color = 0xCA5C05 };

    // Add some points with the above fills.
    var point1 = lxw.lxw_chart_point{ .fill = &fill1 };
    var point2 = lxw.lxw_chart_point{ .fill = &fill2 };
    var point3 = lxw.lxw_chart_point{ .fill = &fill3 };

    // Create an array of the point objects.
    const points_array = [_]*lxw.lxw_chart_point{
        &point1,
        &point2,
        &point3,
    };

    // Create a null-terminated array of pointers as required by the C API
    var points_ptrs: [4][*c]lxw.lxw_chart_point = undefined;
    for (points_array, 0..) |point, i| {
        points_ptrs[i] = point;
    }
    points_ptrs[3] = null;

    // Add/override the points/segments of the chart.
    _ = lxw.chart_series_set_points(series2, @ptrCast(&points_ptrs));

    // Insert the chart into the worksheet.
    _ = lxw.worksheet_insert_chart(worksheet, 17, 3, chart2);

    // Chart 3: Create a pie chart with rotation of the segments.
    const chart3 = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_PIE);

    // Add the first series to the chart.
    const series3 = lxw.chart_add_series(chart3, "=Sheet1!$A$2:$A$4", "=Sheet1!$B$2:$B$4");

    // Set the name for the series instead of the default "Series 1".
    lxw.chart_series_set_name(series3, "Pie sales data");

    // Add a chart title.
    lxw.chart_title_set_name(chart3, "Pie Chart with segment rotation");

    // Change the angle/rotation of the first segment.
    lxw.chart_set_rotation(chart3, 90);

    // Insert the chart into the worksheet.
    _ = lxw.worksheet_insert_chart(worksheet, 33, 3, chart3);

    _ = lxw.workbook_close(workbook);
}
