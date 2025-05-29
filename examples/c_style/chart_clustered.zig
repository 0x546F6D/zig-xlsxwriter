//
// An example of a clustered category chart using the libxlsxwriter library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

// Write some data to the worksheet
fn writeWorksheetData(worksheet: *lxw.lxw_worksheet, bold: *lxw.lxw_format) void {
    _ = lxw.worksheet_write_string(
        worksheet,
        0,
        0,
        "Types",
        bold,
    );
    _ = lxw.worksheet_write_string(
        worksheet,
        1,
        0,
        "Type 1",
        null,
    );
    _ = lxw.worksheet_write_string(
        worksheet,
        4,
        0,
        "Type 2",
        null,
    );

    _ = lxw.worksheet_write_string(
        worksheet,
        0,
        1,
        "Sub Type",
        bold,
    );
    _ = lxw.worksheet_write_string(
        worksheet,
        1,
        1,
        "Sub Type A",
        null,
    );
    _ = lxw.worksheet_write_string(
        worksheet,
        2,
        1,
        "Sub Type B",
        null,
    );
    _ = lxw.worksheet_write_string(
        worksheet,
        3,
        1,
        "Sub Type C",
        null,
    );
    _ = lxw.worksheet_write_string(
        worksheet,
        4,
        1,
        "Sub Type D",
        null,
    );
    _ = lxw.worksheet_write_string(
        worksheet,
        5,
        1,
        "Sub Type E",
        null,
    );

    _ = lxw.worksheet_write_string(
        worksheet,
        0,
        2,
        "Value 1",
        bold,
    );
    _ = lxw.worksheet_write_number(
        worksheet,
        1,
        2,
        5000,
        null,
    );
    _ = lxw.worksheet_write_number(
        worksheet,
        2,
        2,
        2000,
        null,
    );
    _ = lxw.worksheet_write_number(
        worksheet,
        3,
        2,
        250,
        null,
    );
    _ = lxw.worksheet_write_number(
        worksheet,
        4,
        2,
        6000,
        null,
    );
    _ = lxw.worksheet_write_number(
        worksheet,
        5,
        2,
        500,
        null,
    );

    _ = lxw.worksheet_write_string(
        worksheet,
        0,
        3,
        "Value 2",
        bold,
    );
    _ = lxw.worksheet_write_number(
        worksheet,
        1,
        3,
        8000,
        null,
    );
    _ = lxw.worksheet_write_number(
        worksheet,
        2,
        3,
        3000,
        null,
    );
    _ = lxw.worksheet_write_number(
        worksheet,
        3,
        3,
        1000,
        null,
    );
    _ = lxw.worksheet_write_number(
        worksheet,
        4,
        3,
        6000,
        null,
    );
    _ = lxw.worksheet_write_number(
        worksheet,
        5,
        3,
        300,
        null,
    );

    _ = lxw.worksheet_write_string(
        worksheet,
        0,
        4,
        "Value 3",
        bold,
    );
    _ = lxw.worksheet_write_number(
        worksheet,
        1,
        4,
        6000,
        null,
    );
    _ = lxw.worksheet_write_number(
        worksheet,
        2,
        4,
        4000,
        null,
    );
    _ = lxw.worksheet_write_number(
        worksheet,
        3,
        4,
        2000,
        null,
    );
    _ = lxw.worksheet_write_number(
        worksheet,
        4,
        4,
        6500,
        null,
    );
    _ = lxw.worksheet_write_number(
        worksheet,
        5,
        4,
        200,
        null,
    );
}

pub fn main() !void {
    const workbook = lxw.workbook_new(
        "zig-chart_clustered.xlsx",
    );
    const worksheet = lxw.workbook_add_worksheet(
        workbook,
        null,
    );
    const chart = lxw.workbook_add_chart(
        workbook,
        lxw.LXW_CHART_COLUMN,
    );

    // Add a bold format to use to highlight the header cells
    const bold = lxw.workbook_add_format(
        workbook,
    );
    _ = lxw.format_set_bold(
        bold,
    );

    // Write some data for the chart
    writeWorksheetData(worksheet, bold);

    // Configure the series. Note, that the categories are 2D ranges (from
    // column A to column B). This creates the clusters. The series are shown
    // as formula strings for clarity but you can also use variables with the
    // chart_series_set_categories() and chart_series_set_values()
    // functions.
    _ = lxw.chart_add_series(
        chart,
        "=Sheet1!$A$2:$B$6",
        "=Sheet1!$C$2:$C$6",
    );

    _ = lxw.chart_add_series(
        chart,
        "=Sheet1!$A$2:$B$6",
        "=Sheet1!$D$2:$D$6",
    );

    _ = lxw.chart_add_series(
        chart,
        "=Sheet1!$A$2:$B$6",
        "=Sheet1!$E$2:$E$6",
    );

    // Set an Excel chart style
    _ = lxw.chart_set_style(
        chart,
        37,
    );

    // Turn off the legend
    _ = lxw.chart_legend_set_position(
        chart,
        lxw.LXW_CHART_LEGEND_NONE,
    );

    // Insert the chart into the worksheet
    _ = lxw.worksheet_insert_chart(
        worksheet,
        2, // row
        6, // col (G)
        chart,
    );

    _ = lxw.workbook_close(workbook);
}
