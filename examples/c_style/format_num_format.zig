//
// Example of writing some data with numeric formatting to a simple Excel file
// using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-format_num_format.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Widen the first column to make the text clearer.
    _ = lxw.worksheet_set_column(worksheet, 0, 0, 30, null);

    // Add some formats.
    const format01 = lxw.workbook_add_format(workbook);
    const format02 = lxw.workbook_add_format(workbook);
    const format03 = lxw.workbook_add_format(workbook);
    const format04 = lxw.workbook_add_format(workbook);
    const format05 = lxw.workbook_add_format(workbook);
    const format06 = lxw.workbook_add_format(workbook);
    const format07 = lxw.workbook_add_format(workbook);
    const format08 = lxw.workbook_add_format(workbook);
    const format09 = lxw.workbook_add_format(workbook);
    const format10 = lxw.workbook_add_format(workbook);
    const format11 = lxw.workbook_add_format(workbook);

    // Set some example number formats.
    _ = lxw.format_set_num_format(format01, "0.000");
    _ = lxw.format_set_num_format(format02, "#,##0");
    _ = lxw.format_set_num_format(format03, "#,##0.00");
    _ = lxw.format_set_num_format(format04, "0.00");
    _ = lxw.format_set_num_format(format05, "mm/dd/yy");
    _ = lxw.format_set_num_format(format06, "mmm d yyyy");
    _ = lxw.format_set_num_format(format07, "d mmmm yyyy");
    _ = lxw.format_set_num_format(format08, "dd/mm/yyyy hh:mm AM/PM");
    _ = lxw.format_set_num_format(format09, "0 \"dollar and\" .00 \"cents\"");

    // Write data using the formats.
    _ = lxw.worksheet_write_number(worksheet, 0, 0, 3.1415926, null); // 3.1415926
    _ = lxw.worksheet_write_number(worksheet, 1, 0, 3.1415926, format01); // 3.142
    _ = lxw.worksheet_write_number(worksheet, 2, 0, 1234.56, format02); // 1,235
    _ = lxw.worksheet_write_number(worksheet, 3, 0, 1234.56, format03); // 1,234.56
    _ = lxw.worksheet_write_number(worksheet, 4, 0, 49.99, format04); // 49.99
    _ = lxw.worksheet_write_number(worksheet, 5, 0, 36892.521, format05); // 01/01/01
    _ = lxw.worksheet_write_number(worksheet, 6, 0, 36892.521, format06); // Jan 1 2001
    _ = lxw.worksheet_write_number(worksheet, 7, 0, 36892.521, format07); // 1 January 2001
    _ = lxw.worksheet_write_number(worksheet, 8, 0, 36892.521, format08); // 01/01/2001 12:30 AM
    _ = lxw.worksheet_write_number(worksheet, 9, 0, 1.87, format09); // 1 dollar and .87 cents

    // Show limited conditional number formats.
    _ = lxw.format_set_num_format(format10, "[Green]General;[Red]-General;General");
    _ = lxw.worksheet_write_number(worksheet, 10, 0, 123, format10); // > 0 Green
    _ = lxw.worksheet_write_number(worksheet, 11, 0, -45, format10); // < 0 Red
    _ = lxw.worksheet_write_number(worksheet, 12, 0, 0, format10); // = 0 Default color

    // Format a Zip code.
    _ = lxw.format_set_num_format(format11, "00000");
    _ = lxw.worksheet_write_number(worksheet, 13, 0, 1209, format11);

    _ = lxw.workbook_close(workbook);
}
