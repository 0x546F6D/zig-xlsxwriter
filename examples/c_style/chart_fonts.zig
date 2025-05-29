//
// An example of a simple Excel chart with user defined fonts using the
// libxlsxwriter library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-chart_fonts.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Write some data for the chart
    _ = lxw.worksheet_write_number(worksheet, 0, 0, 10, null);
    _ = lxw.worksheet_write_number(worksheet, 1, 0, 40, null);
    _ = lxw.worksheet_write_number(worksheet, 2, 0, 50, null);
    _ = lxw.worksheet_write_number(worksheet, 3, 0, 20, null);
    _ = lxw.worksheet_write_number(worksheet, 4, 0, 10, null);
    _ = lxw.worksheet_write_number(worksheet, 5, 0, 50, null);

    // Create a chart object
    const chart = lxw.workbook_add_chart(workbook, lxw.LXW_CHART_LINE);

    // Configure the chart
    _ = lxw.chart_add_series(chart, null, "Sheet1!$A$1:$A$6");

    // Create some fonts to use in the chart
    var font1 = lxw.lxw_chart_font{
        .name = "Calibri",
        .color = lxw.LXW_COLOR_BLUE,
        .bold = lxw.LXW_EXPLICIT_FALSE,
        .italic = lxw.LXW_EXPLICIT_FALSE,
        .underline = 0,
        .size = 0,
        .rotation = 0,
        .baseline = 0,
        .pitch_family = 0,
        .charset = 0,
    };

    var font2 = lxw.lxw_chart_font{
        .name = "Courier",
        .color = 0x92D050,
        .bold = lxw.LXW_EXPLICIT_FALSE,
        .italic = lxw.LXW_EXPLICIT_FALSE,
        .underline = 0,
        .size = 0,
        .rotation = 0,
        .baseline = 0,
        .pitch_family = 0,
        .charset = 0,
    };

    var font3 = lxw.lxw_chart_font{
        .name = "Arial",
        .color = 0x00B0F0,
        .bold = lxw.LXW_EXPLICIT_FALSE,
        .italic = lxw.LXW_EXPLICIT_FALSE,
        .underline = 0,
        .size = 0,
        .rotation = 0,
        .baseline = 0,
        .pitch_family = 0,
        .charset = 0,
    };

    var font4 = lxw.lxw_chart_font{
        .name = "Century",
        .color = lxw.LXW_COLOR_RED,
        .bold = lxw.LXW_EXPLICIT_FALSE,
        .italic = lxw.LXW_EXPLICIT_FALSE,
        .underline = 0,
        .size = 0,
        .rotation = 0,
        .baseline = 0,
        .pitch_family = 0,
        .charset = 0,
    };

    var font5 = lxw.lxw_chart_font{
        .name = null,
        .color = 0,
        .bold = lxw.LXW_EXPLICIT_FALSE,
        .italic = lxw.LXW_EXPLICIT_FALSE,
        .underline = 0,
        .size = 0,
        .rotation = -30,
        .baseline = 0,
        .pitch_family = 0,
        .charset = 0,
    };

    var font6 = lxw.lxw_chart_font{
        .name = null,
        .color = 0x7030A0,
        .bold = lxw.LXW_TRUE,
        .italic = lxw.LXW_TRUE,
        .underline = lxw.LXW_TRUE,
        .size = 0,
        .rotation = 0,
        .baseline = 0,
        .pitch_family = 0,
        .charset = 0,
    };

    // Write the chart title with a font
    _ = lxw.chart_title_set_name(chart, "Test Results");
    _ = lxw.chart_title_set_name_font(chart, &font1);

    // Write the Y axis with a font
    _ = lxw.chart_axis_set_name(chart.*.y_axis, "Units");
    _ = lxw.chart_axis_set_name_font(chart.*.y_axis, &font2);
    _ = lxw.chart_axis_set_num_font(chart.*.y_axis, &font3);

    // Write the X axis with a font
    _ = lxw.chart_axis_set_name(chart.*.x_axis, "Month");
    _ = lxw.chart_axis_set_name_font(chart.*.x_axis, &font4);
    _ = lxw.chart_axis_set_num_font(chart.*.x_axis, &font5);

    // Display the chart legend at the bottom of the chart
    _ = lxw.chart_legend_set_position(chart, lxw.LXW_CHART_LEGEND_BOTTOM);
    _ = lxw.chart_legend_set_font(chart, &font6);

    // Insert the chart into the worksheet
    _ = lxw.worksheet_insert_chart(worksheet, 0, 2, chart);

    _ = lxw.workbook_close(workbook);
}
