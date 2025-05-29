// A simple Unicode UTF-8 example using libxlsxwriter.
//
// Note: The source file must be UTF-8 encoded.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//

const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-utf8.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    _ = lxw.worksheet_write_string(worksheet, 2, 1, "Это фраза на русском!", null);

    _ = lxw.workbook_close(workbook);
}
