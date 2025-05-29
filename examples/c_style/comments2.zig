//
// An example of writing cell comments to a worksheet using libxlsxwriter.
//
// Each of the worksheets demonstrates different features of cell comments.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-comments2.xlsx");
    const worksheet1 = lxw.workbook_add_worksheet(workbook, null);
    const worksheet2 = lxw.workbook_add_worksheet(workbook, null);
    const worksheet3 = lxw.workbook_add_worksheet(workbook, null);
    const worksheet4 = lxw.workbook_add_worksheet(workbook, null);
    const worksheet5 = lxw.workbook_add_worksheet(workbook, null);
    const worksheet6 = lxw.workbook_add_worksheet(workbook, null);
    const worksheet7 = lxw.workbook_add_worksheet(workbook, null);

    const text_wrap = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_text_wrap(text_wrap);
    _ = lxw.format_set_align(text_wrap, lxw.LXW_ALIGN_VERTICAL_TOP);

    // Example 1: Simple cell comment without formatting
    _ = lxw.worksheet_set_column(worksheet1, 2, 2, 25, null);
    _ = lxw.worksheet_set_row(worksheet1, 2, 50, null);

    _ = lxw.worksheet_write_string(
        worksheet1,
        2,
        2,
        "Hold the mouse over this cell to see the comment.",
        text_wrap,
    );
    _ = lxw.worksheet_write_comment(worksheet1, 2, 2, "This is a comment.");

    // Example 2: Visible and hidden comments
    _ = lxw.worksheet_set_column(worksheet2, 2, 2, 25, null);
    _ = lxw.worksheet_set_row(worksheet2, 2, 50, null);

    _ = lxw.worksheet_write_string(
        worksheet2,
        2,
        2,
        "This cell comment is visible.",
        text_wrap,
    );

    var options2 = lxw.lxw_comment_options{
        .visible = lxw.LXW_COMMENT_DISPLAY_VISIBLE,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = lxw.worksheet_write_comment_opt(worksheet2, 2, 2, "Hello.", &options2);

    _ = lxw.worksheet_write_string(
        worksheet2,
        5,
        2,
        "This cell comment isn't visible until you pass the mouse over it (the default).",
        text_wrap,
    );
    _ = lxw.worksheet_write_comment(worksheet2, 5, 2, "Hello.");

    // Example 3: Worksheet level comment visibility
    _ = lxw.worksheet_set_column(worksheet3, 2, 2, 25, null);
    _ = lxw.worksheet_set_row(worksheet3, 2, 50, null);
    _ = lxw.worksheet_set_row(worksheet3, 5, 50, null);
    _ = lxw.worksheet_set_row(worksheet3, 8, 50, null);

    _ = lxw.worksheet_show_comments(worksheet3);

    _ = lxw.worksheet_write_string(
        worksheet3,
        2,
        2,
        "This cell comment is visible, explicitly.",
        text_wrap,
    );

    var options3a = lxw.lxw_comment_options{
        .visible = lxw.LXW_COMMENT_DISPLAY_VISIBLE,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = lxw.worksheet_write_comment_opt(worksheet3, 2, 2, "Hello", &options3a);

    _ = lxw.worksheet_write_string(
        worksheet3,
        5,
        2,
        "This cell comment is also visible because we used worksheet_show_comments().",
        text_wrap,
    );
    _ = lxw.worksheet_write_comment(worksheet3, 5, 2, "Hello");

    _ = lxw.worksheet_write_string(
        worksheet3,
        8,
        2,
        "However, we can still override it locally.",
        text_wrap,
    );

    var options3b = lxw.lxw_comment_options{
        .visible = lxw.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = lxw.worksheet_write_comment_opt(worksheet3, 8, 2, "Hello", &options3b);

    // Example 4: Comment box dimensions
    _ = lxw.worksheet_set_column(worksheet4, 2, 2, 25, null);
    _ = lxw.worksheet_set_row(worksheet4, 2, 50, null);
    _ = lxw.worksheet_set_row(worksheet4, 5, 50, null);
    _ = lxw.worksheet_set_row(worksheet4, 8, 50, null);
    _ = lxw.worksheet_set_row(worksheet4, 15, 50, null);
    _ = lxw.worksheet_set_row(worksheet4, 18, 50, null);

    _ = lxw.worksheet_show_comments(worksheet4);

    _ = lxw.worksheet_write_string(
        worksheet4,
        2,
        2,
        "This cell comment is default size.",
        text_wrap,
    );
    _ = lxw.worksheet_write_comment(worksheet4, 2, 2, "Hello");

    _ = lxw.worksheet_write_string(
        worksheet4,
        5,
        2,
        "This cell comment is twice as wide.",
        text_wrap,
    );

    var options4a = lxw.lxw_comment_options{
        .visible = lxw.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 2.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = lxw.worksheet_write_comment_opt(worksheet4, 5, 2, "Hello", &options4a);

    _ = lxw.worksheet_write_string(
        worksheet4,
        8,
        2,
        "This cell comment is twice as high.",
        text_wrap,
    );

    var options4b = lxw.lxw_comment_options{
        .visible = lxw.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 2.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = lxw.worksheet_write_comment_opt(worksheet4, 8, 2, "Hello", &options4b);

    _ = lxw.worksheet_write_string(
        worksheet4,
        15,
        2,
        "This cell comment is scaled in both directions.",
        text_wrap,
    );

    var options4c = lxw.lxw_comment_options{
        .visible = lxw.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.2,
        .y_scale = 0.5,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = lxw.worksheet_write_comment_opt(worksheet4, 15, 2, "Hello", &options4c);

    _ = lxw.worksheet_write_string(
        worksheet4,
        18,
        2,
        "This cell comment has width and height specified in pixels.",
        text_wrap,
    );

    var options4d = lxw.lxw_comment_options{
        .visible = lxw.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 200,
        .height = 50,
    };
    _ = lxw.worksheet_write_comment_opt(worksheet4, 18, 2, "Hello", &options4d);

    // Example 5: Comment positioning
    _ = lxw.worksheet_set_column(worksheet5, 2, 2, 25, null);
    _ = lxw.worksheet_set_row(worksheet5, 2, 50, null);
    _ = lxw.worksheet_set_row(worksheet5, 5, 50, null);
    _ = lxw.worksheet_set_row(worksheet5, 8, 50, null);

    _ = lxw.worksheet_show_comments(worksheet5);

    _ = lxw.worksheet_write_string(
        worksheet5,
        2,
        2,
        "This cell comment is in the default position.",
        text_wrap,
    );
    _ = lxw.worksheet_write_comment(worksheet5, 2, 2, "Hello");

    _ = lxw.worksheet_write_string(
        worksheet5,
        5,
        2,
        "This cell comment has been moved to another cell.",
        text_wrap,
    );

    var options5a = lxw.lxw_comment_options{
        .visible = lxw.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 3,
        .start_col = 4,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = lxw.worksheet_write_comment_opt(worksheet5, 5, 2, "Hello", &options5a);

    _ = lxw.worksheet_write_string(
        worksheet5,
        8,
        2,
        "This cell comment has been shifted within its default cell.",
        text_wrap,
    );

    var options5b = lxw.lxw_comment_options{
        .visible = lxw.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 30,
        .y_offset = 12,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = lxw.worksheet_write_comment_opt(worksheet5, 8, 2, "Hello", &options5b);

    // Example 6: Comment colors
    _ = lxw.worksheet_set_column(worksheet6, 2, 2, 25, null);
    _ = lxw.worksheet_set_row(worksheet6, 2, 50, null);
    _ = lxw.worksheet_set_row(worksheet6, 5, 50, null);
    _ = lxw.worksheet_set_row(worksheet6, 8, 50, null);

    _ = lxw.worksheet_show_comments(worksheet6);

    _ = lxw.worksheet_write_string(
        worksheet6,
        2,
        2,
        "This cell comment has a different color.",
        text_wrap,
    );

    var options6a = lxw.lxw_comment_options{
        .visible = lxw.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x008000, // Green
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = lxw.worksheet_write_comment_opt(worksheet6, 2, 2, "Hello", &options6a);

    _ = lxw.worksheet_write_string(
        worksheet6,
        5,
        2,
        "This cell comment has the default color.",
        text_wrap,
    );
    _ = lxw.worksheet_write_comment(worksheet6, 5, 2, "Hello");

    _ = lxw.worksheet_write_string(
        worksheet6,
        8,
        2,
        "This cell comment has a different color.",
        text_wrap,
    );

    var options6b = lxw.lxw_comment_options{
        .visible = lxw.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0xFF6600,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = null,
        .width = 0,
        .height = 0,
    };
    _ = lxw.worksheet_write_comment_opt(worksheet6, 8, 2, "Hello", &options6b);

    // Example 7: Comment author
    _ = lxw.worksheet_set_column(worksheet7, 2, 2, 30, null);
    _ = lxw.worksheet_set_row(worksheet7, 2, 50, null);
    _ = lxw.worksheet_set_row(worksheet7, 5, 60, null);

    _ = lxw.worksheet_write_string(
        worksheet7,
        2,
        2,
        "Move the mouse over this cell and you will see 'Cell C3 " ++
            "commented by' (blank) in the status bar at the bottom.",
        text_wrap,
    );
    _ = lxw.worksheet_write_comment(worksheet7, 2, 2, "Hello");

    _ = lxw.worksheet_write_string(
        worksheet7,
        5,
        2,
        "Move the mouse over this cell and you will see 'Cell C6 " ++
            "commented by libxlsxwriter' in the status bar at the bottom.",
        text_wrap,
    );

    var options7a = lxw.lxw_comment_options{
        .visible = lxw.LXW_COMMENT_DISPLAY_HIDDEN,
        .x_scale = 1.0,
        .y_scale = 1.0,
        .color = 0x0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
        .author = "libxlsxwriter",
        .width = 0,
        .height = 0,
    };
    _ = lxw.worksheet_write_comment_opt(worksheet7, 5, 2, "Hello", &options7a);

    _ = lxw.workbook_close(workbook);
}
