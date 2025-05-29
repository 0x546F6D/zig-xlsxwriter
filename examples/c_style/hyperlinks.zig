//
// Example of writing urls/hyperlinks with the libxlsxwriter library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    // Create a new workbook
    const workbook = lxw.workbook_new("zig-hyperlinks.xlsx");

    // Add a worksheet
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Get the default url format (used in the overwriting examples below)
    const url_format = lxw.workbook_get_default_url_format(workbook);

    // Create a user defined link format
    const red_format = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_underline(red_format, lxw.LXW_UNDERLINE_SINGLE);
    _ = lxw.format_set_font_color(red_format, lxw.LXW_COLOR_RED);

    // Widen the first column to make the text clearer
    _ = lxw.worksheet_set_column(worksheet, 0, 0, 30, null);

    // Write a hyperlink. A default blue underline will be used if the format is NULL
    _ = lxw.worksheet_write_url(worksheet, 0, 0, "http://libxlsxwriter.github.io", null);

    // Write a hyperlink but overwrite the displayed string. Note, we need to
    // specify the format for the string to match the default hyperlink
    _ = lxw.worksheet_write_url(worksheet, 2, 0, "http://libxlsxwriter.github.io", null);
    _ = lxw.worksheet_write_string(worksheet, 2, 0, "Read the documentation.", url_format);

    // Write a hyperlink with a different format
    _ = lxw.worksheet_write_url(worksheet, 4, 0, "http://libxlsxwriter.github.io", red_format);

    // Write a mail hyperlink
    _ = lxw.worksheet_write_url(worksheet, 6, 0, "mailto:jmcnamara@cpan.org", null);

    // Write a mail hyperlink and overwrite the displayed string. We again
    // specify the format for the string to match the default hyperlink
    _ = lxw.worksheet_write_url(worksheet, 8, 0, "mailto:jmcnamara@cpan.org", null);
    _ = lxw.worksheet_write_string(worksheet, 8, 0, "Drop me a line.", url_format);

    // Close the workbook, save the file and free any memory
    _ = lxw.workbook_close(workbook);
}
