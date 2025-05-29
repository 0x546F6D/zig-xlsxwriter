//
// An example of merging cells containing a rich string using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-merge_rich_string.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Configure a format for the merged range.
    const merge_format = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_align(merge_format, lxw.LXW_ALIGN_CENTER);
    _ = lxw.format_set_align(
        merge_format,
        lxw.LXW_ALIGN_VERTICAL_CENTER,
    );
    _ = lxw.format_set_border(merge_format, lxw.LXW_BORDER_THIN);

    // Configure formats for the rich string.
    const red = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_font_color(red, lxw.LXW_COLOR_RED);

    const blue = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_font_color(blue, lxw.LXW_COLOR_BLUE);

    // Create the fragments for the rich string.
    var fragment1 = lxw.lxw_rich_string_tuple{
        .format = null,
        .string = "This is ",
    };
    var fragment2 = lxw.lxw_rich_string_tuple{
        .format = red,
        .string = "red",
    };
    var fragment3 = lxw.lxw_rich_string_tuple{
        .format = null,
        .string = " and this is ",
    };
    var fragment4 = lxw.lxw_rich_string_tuple{
        .format = blue,
        .string = "blue",
    };

    var rich_string = [_:null]?*lxw.lxw_rich_string_tuple{
        &fragment1,
        &fragment2,
        &fragment3,
        &fragment4,
        null,
    };

    // Write an empty string to the merged range.
    _ = lxw.worksheet_merge_range(
        worksheet,
        1,
        1,
        4,
        3,
        "",
        merge_format,
    );

    // We then overwrite the first merged cell with a rich string. Note that
    // we must also pass the cell format used in the merged cells format at
    // the end.
    _ = lxw.worksheet_write_rich_string(
        worksheet,
        1,
        1,
        &rich_string,
        merge_format,
    );

    _ = lxw.workbook_close(workbook);
}
