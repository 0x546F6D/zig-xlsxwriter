// Example of writing some data to a simple Excel file using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//

const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-hello.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    _ = lxw.worksheet_write_string(worksheet, 0, 0, "Hello", null);
    _ = lxw.worksheet_write_number(worksheet, 1, 0, 123, null);

    _ = lxw.workbook_close(workbook);
}
