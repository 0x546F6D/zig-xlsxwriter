//
// Example of writing dates and times in Excel using a Unix datetime and date
// formatting.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-dates_and_times03.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Add a format with date formatting.
    const format = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_num_format(format, "mmm d yyyy hh:mm AM/PM");

    // Widen the first column to make the text clearer.
    _ = lxw.worksheet_set_column(worksheet, 0, 0, 20, null);

    // Write some Unix datetimes with formatting.

    // 1970-01-01. The Unix epoch.
    _ = lxw.worksheet_write_unixtime(worksheet, 0, 0, 0, format);

    // 2000-01-01.
    _ = lxw.worksheet_write_unixtime(worksheet, 1, 0, 1577836800, format);

    // 1900-01-01.
    _ = lxw.worksheet_write_unixtime(worksheet, 2, 0, -2208988800, format);

    _ = lxw.workbook_close(workbook);
}
