//
// Anatomy of a simple libxlsxwriter program.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    // Create a new workbook
    const workbook = lxw.workbook_new("zig-anatomy.xlsx");

    // Add a worksheet with a user defined sheet name
    const worksheet1 = lxw.workbook_add_worksheet(workbook, "Demo");

    // Add a worksheet with Excel's default sheet name: Sheet2
    const worksheet2 = lxw.workbook_add_worksheet(workbook, null);

    // Add some cell formats
    const myformat1 = lxw.workbook_add_format(workbook);
    const myformat2 = lxw.workbook_add_format(workbook);

    // Set the bold property for the first format
    _ = lxw.format_set_bold(myformat1);

    // Set a number format for the second format
    _ = lxw.format_set_num_format(myformat2, "$#,##0.00");

    // Widen the first column to make the text clearer
    _ = lxw.worksheet_set_column(worksheet1, 0, 0, 20, null);

    // Write some unformatted data
    _ = lxw.worksheet_write_string(worksheet1, 0, 0, "Peach", null);
    _ = lxw.worksheet_write_string(worksheet1, 1, 0, "Plum", null);

    // Write formatted data
    _ = lxw.worksheet_write_string(worksheet1, 2, 0, "Pear", myformat1);

    // Formats can be reused
    _ = lxw.worksheet_write_string(worksheet1, 3, 0, "Persimmon", myformat1);

    // Write some numbers
    _ = lxw.worksheet_write_number(worksheet1, 5, 0, 123, null);
    _ = lxw.worksheet_write_number(worksheet1, 6, 0, 4567.555, myformat2);

    // Write to the second worksheet
    _ = lxw.worksheet_write_string(worksheet2, 0, 0, "Some text", myformat1);

    // Close the workbook, save the file and free any memory
    const error_code = lxw.workbook_close(workbook);

    // Check if there was any error creating the xlsx file
    if (error_code != lxw.LXW_NO_ERROR) {
        std.debug.print("Error in workbook_close().\nError {d} = {s}\n", .{ error_code, lxw.lxw_strerror(error_code) });
        return error.WorkbookCloseError;
    }
}
