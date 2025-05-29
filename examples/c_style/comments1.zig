//
// An example of writing cell comments to a worksheet using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-comments1.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    _ = lxw.worksheet_write_string(worksheet, 0, 0, "Hello", null);

    _ = lxw.worksheet_write_comment(worksheet, 0, 0, "This is a comment");

    _ = lxw.workbook_close(workbook);
}
