//
// Example of writing dates and times in Excel using an lxw_datetime struct
// and date formatting.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    // A datetime to display.
    var datetime = lxw.lxw_datetime{
        .year = 2013,
        .month = 2,
        .day = 28,
        .hour = 12,
        .min = 0,
        .sec = 0.0,
    };

    const workbook = lxw.workbook_new("zig-dates_and_times02.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Add a format with date formatting.
    const format = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_num_format(format, "mmm d yyyy hh:mm AM/PM");

    // Widen the first column to make the text clearer.
    _ = lxw.worksheet_set_column(worksheet, 0, 0, 20, null);

    // Write the datetime without formatting.
    _ = lxw.worksheet_write_datetime(worksheet, 0, 0, &datetime, null); // 41333.5

    // Write the datetime with formatting.
    _ = lxw.worksheet_write_datetime(worksheet, 1, 0, &datetime, format); // Feb 28 2013 12:00 PM

    _ = lxw.workbook_close(workbook);
}
