const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook =
        lxw.workbook_new(
            "zig-rich_strings.xlsx",
        );
    const worksheet =
        lxw.workbook_add_worksheet(
            workbook,
            null,
        );

    // Set up some formats to use.
    const bold =
        lxw.workbook_add_format(workbook);
    _ = lxw.format_set_bold(bold);

    const italic =
        lxw.workbook_add_format(workbook);
    _ = lxw.format_set_italic(italic);

    const red =
        lxw.workbook_add_format(workbook);
    _ = lxw.format_set_font_color(
        red,
        lxw.LXW_COLOR_RED,
    );

    const blue =
        lxw.workbook_add_format(workbook);
    _ = lxw.format_set_font_color(
        blue,
        lxw.LXW_COLOR_BLUE,
    );

    const center =
        lxw.workbook_add_format(workbook);
    _ = lxw.format_set_align(
        center,
        lxw.LXW_ALIGN_CENTER,
    );

    const superscript =
        lxw.workbook_add_format(workbook);
    _ = lxw.format_set_font_script(
        superscript,
        lxw.LXW_FONT_SUPERSCRIPT,
    );

    // Make the first column wider for clarity.
    _ = lxw.worksheet_set_column(
        worksheet,
        0,
        0,
        30,
        null,
    );

    // Example 1: Bold and italic text
    // Write individual cells with appropriate formatting
    var fragment11 =
        lxw.lxw_rich_string_tuple{
            .format = null,
            .string = "This is ",
        };
    var fragment12 =
        lxw.lxw_rich_string_tuple{
            .format = bold,
            .string = "bold",
        };
    var fragment13 =
        lxw.lxw_rich_string_tuple{
            .format = null,
            .string = " and this is ",
        };
    var fragment14 =
        lxw.lxw_rich_string_tuple{
            .format = italic,
            .string = "italic",
        };
    var rich_string1 =
        [_:null]?*lxw.lxw_rich_string_tuple{
            &fragment11,
            &fragment12,
            &fragment13,
            &fragment14,
            null,
        };
    _ = lxw.worksheet_write_rich_string(
        worksheet,
        0,
        0,
        &rich_string1,
        null,
    );

    // Example 2: Red and blue text

    var fragment21 =
        lxw.lxw_rich_string_tuple{
            .format = null,
            .string = "This is ",
        };
    var fragment22 =
        lxw.lxw_rich_string_tuple{
            .format = red,
            .string = "red",
        };
    var fragment23 =
        lxw.lxw_rich_string_tuple{
            .format = null,
            .string = " and this is ",
        };
    var fragment24 =
        lxw.lxw_rich_string_tuple{
            .format = blue,
            .string = "blue",
        };
    var rich_string2 =
        [_:null]?*lxw.lxw_rich_string_tuple{
            &fragment21,
            &fragment22,
            &fragment23,
            &fragment24,
            null,
        };
    _ = lxw.worksheet_write_rich_string(
        worksheet,
        2,
        0,
        &rich_string2,
        null,
    );

    // Example 3. A rich string plus cell formatting.
    var fragment31 = lxw.lxw_rich_string_tuple{
        .format = null,
        .string = "Some ",
    };
    var fragment32 = lxw.lxw_rich_string_tuple{
        .format = bold,
        .string = "bold text",
    };
    var fragment33 = lxw.lxw_rich_string_tuple{
        .format = null,
        .string = " centered",
    };
    var rich_string3 =
        [_:null]?*lxw.lxw_rich_string_tuple{
            &fragment31,
            &fragment32,
            &fragment33,
            null,
        };
    _ = lxw.worksheet_write_rich_string(
        worksheet,
        4,
        0,
        &rich_string3,
        center,
    );

    // Example 4: Math example with superscript
    var fragment41 = lxw.lxw_rich_string_tuple{
        .format = italic,
        .string = "j =k",
    };
    var fragment42 = lxw.lxw_rich_string_tuple{
        .format = superscript,
        .string = "(n-1)",
    };
    var rich_string4 =
        [_:null]?*lxw.lxw_rich_string_tuple{
            &fragment41,
            &fragment42,
            null,
        };

    _ = lxw.worksheet_write_rich_string(
        worksheet,
        6,
        0,
        &rich_string4,
        center,
    );

    _ = lxw.workbook_close(workbook);
}
