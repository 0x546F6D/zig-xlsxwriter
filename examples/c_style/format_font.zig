// Example of writing some data with font formatting to a simple Excel
// file using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//

const lxw = @import("lxw");

pub fn main() !void {
    // Create a new workbook.
    const workbook = lxw.workbook_new("zig-format_font.xlsx");

    // Add a worksheet.
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Widen the first column to make the text clearer.
    _ = lxw.worksheet_set_column(worksheet, 0, 0, 20, null);

    // Add some formats.
    const format1 = lxw.workbook_add_format(workbook);
    const format2 = lxw.workbook_add_format(workbook);
    const format3 = lxw.workbook_add_format(workbook);

    // Set the bold property for format 1.
    _ = lxw.format_set_bold(format1);

    // Set the italic property for format 2.
    _ = lxw.format_set_italic(format2);

    // Set the bold and italic properties for format 3.
    _ = lxw.format_set_bold(format3);
    _ = lxw.format_set_italic(format3);

    // Write some formatted strings.
    _ = lxw.worksheet_write_string(worksheet, 0, 0, "This is bold", format1);
    _ = lxw.worksheet_write_string(worksheet, 1, 0, "This is italic", format2);
    _ = lxw.worksheet_write_string(worksheet, 2, 0, "Bold and italic", format3);

    // Close the workbook, save the file and free any memory.
    _ = lxw.workbook_close(workbook);
}
