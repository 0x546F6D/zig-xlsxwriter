pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
    defer workbook.deinit() catch {};

    const worksheet1 = try workbook.addWorkSheet(null);
    const worksheet2 = try workbook.addWorkSheet(null);
    const worksheet3 = try workbook.addWorkSheet(null);
    const worksheet4 = try workbook.addWorkSheet(null);
    const worksheet5 = try workbook.addWorkSheet(null);
    const worksheet6 = try workbook.addWorkSheet(null);
    const worksheet7 = try workbook.addWorkSheet(null);
    const worksheet8 = try workbook.addWorkSheet(null);
    var worksheet: WorkSheet = undefined;

    const text_wrap = try workbook.addFormat();
    text_wrap.setTextWrap();
    text_wrap.setAlign(.vertical_top);

    // Example 1: Simple cell comment without formatting
    worksheet = worksheet1;
    try worksheet.setColumn(2, 2, 25, .none);
    try worksheet.setRow(2, 50, .none);

    try worksheet.writeString(
        2,
        2,
        "Hold the mouse over this cell to see the comment.",
        text_wrap,
    );
    try worksheet.writeComment(2, 2, "This is a comment.");

    // Example 2: Visible and hidden comments
    worksheet = worksheet2;
    try worksheet.setColumn(2, 2, 25, .none);
    try worksheet.setRow(2, 50, .none);

    try worksheet.writeString(
        2,
        2,
        "This cell comment is visible.",
        text_wrap,
    );

    const options2 = xwz.CommentOptions{
        .visible = .visible,
    };
    try worksheet.writeCommentOpt(2, 2, "Hello.", options2);

    try worksheet.writeString(
        5,
        2,
        "This cell comment isn't visible until you pass the mouse over it (the default).",
        text_wrap,
    );
    try worksheet.writeComment(5, 2, "Hello.");

    // Example 3: Worksheet level comment visibility
    worksheet = worksheet3;
    try worksheet.setColumn(2, 2, 25, .none);
    try worksheet.setRow(2, 50, .none);
    try worksheet.setRow(5, 50, .none);
    try worksheet.setRow(8, 50, .none);

    worksheet.showComments();

    try worksheet.writeString(
        2,
        2,
        "This cell comment is visible, explicitly.",
        text_wrap,
    );

    const options3a = xwz.CommentOptions{
        .visible = .visible,
    };
    try worksheet.writeCommentOpt(2, 2, "Hello", options3a);

    try worksheet.writeString(
        5,
        2,
        "This cell comment is also visible because we used worksheet_show_comments().",
        text_wrap,
    );
    try worksheet.writeComment(5, 2, "Hello");

    try worksheet.writeString(
        8,
        2,
        "However, we can still override it locally.",
        text_wrap,
    );

    const options3b = xwz.CommentOptions{
        .visible = .hidden,
    };
    try worksheet.writeCommentOpt(8, 2, "Hello", options3b);

    // Example 4: Comment box dimensions
    worksheet = worksheet4;
    try worksheet.setColumn(2, 2, 25, .none);
    try worksheet.setRow(2, 50, .none);
    try worksheet.setRow(5, 50, .none);
    try worksheet.setRow(8, 50, .none);
    try worksheet.setRow(15, 50, .none);
    try worksheet.setRow(18, 50, .none);

    worksheet.showComments();

    try worksheet.writeString(
        2,
        2,
        "This cell comment is default size.",
        text_wrap,
    );
    try worksheet.writeCommentOpt(2, 2, "Hello", .default);

    try worksheet.writeString(
        5,
        2,
        "This cell comment is twice as wide.",
        text_wrap,
    );

    const options4a = xwz.CommentOptions{
        .x_scale = 2.0,
    };
    try worksheet.writeCommentOpt(5, 2, "Hello", options4a);

    try worksheet.writeString(
        8,
        2,
        "This cell comment is twice as high.",
        text_wrap,
    );

    const options4b = xwz.CommentOptions{
        .y_scale = 2.0,
    };
    try worksheet.writeCommentOpt(8, 2, "Hello", options4b);

    try worksheet.writeString(
        15,
        2,
        "This cell comment is scaled in both directions.",
        text_wrap,
    );

    const options4c = xwz.CommentOptions{
        .x_scale = 1.2,
        .y_scale = 0.5,
    };
    try worksheet.writeCommentOpt(15, 2, "Hello", options4c);

    try worksheet.writeString(
        18,
        2,
        "This cell comment has width and height specified in pixels.",
        text_wrap,
    );

    const options4d = xwz.CommentOptions{
        .width = 200,
        .height = 50,
    };
    try worksheet.writeCommentOpt(18, 2, "Hello", options4d);

    // Example 5: Comment positioning
    worksheet = worksheet5;
    try worksheet.setColumn(2, 2, 25, .none);
    try worksheet.setRow(2, 50, .none);
    try worksheet.setRow(5, 50, .none);
    try worksheet.setRow(8, 50, .none);

    worksheet.showComments();

    try worksheet.writeString(
        2,
        2,
        "This cell comment is in the default position.",
        text_wrap,
    );
    try worksheet.writeComment(2, 2, "Hello");

    try worksheet.writeString(
        5,
        2,
        "This cell comment has been moved to another cell.",
        text_wrap,
    );

    const options5a = xwz.CommentOptions{
        .start_row = 3,
        .start_col = 4,
    };
    try worksheet.writeCommentOpt(5, 2, "Hello", options5a);

    try worksheet.writeString(
        8,
        2,
        "This cell comment has been shifted within its default cell.",
        text_wrap,
    );

    const options5b = xwz.CommentOptions{
        .x_offset = 30,
        .y_offset = 12,
    };
    try worksheet.writeCommentOpt(8, 2, "Hello", options5b);

    // Example 6: Comment colors
    worksheet = worksheet6;
    try worksheet.setColumn(2, 2, 25, .none);
    try worksheet.setRow(2, 50, .none);
    try worksheet.setRow(5, 50, .none);
    try worksheet.setRow(8, 50, .none);

    worksheet.showComments();

    try worksheet.writeString(
        2,
        2,
        "This cell comment has a different color.",
        text_wrap,
    );

    const options6a = xwz.CommentOptions{
        .color = @enumFromInt(0x008000), // Green
    };
    try worksheet.writeCommentOpt(2, 2, "Hello", options6a);

    try worksheet.writeString(
        5,
        2,
        "This cell comment has the default color.",
        text_wrap,
    );
    try worksheet.writeComment(5, 2, "Hello");

    try worksheet.writeString(
        8,
        2,
        "This cell comment has a different color.",
        text_wrap,
    );

    const options6b = xwz.CommentOptions{
        .color = @enumFromInt(0xFF6600),
    };
    try worksheet.writeCommentOpt(8, 2, "Hello", options6b);

    // Example 7: Comment author
    worksheet = worksheet7;
    try worksheet.setColumn(2, 2, 30, .none);
    try worksheet.setRow(2, 50, .none);
    try worksheet.setRow(5, 60, .none);

    try worksheet.writeString(
        2,
        2,
        "Move the mouse over this cell and you will see 'Cell C3 " ++
            "commented by' (blank) in the status bar at the bottom.",
        text_wrap,
    );
    try worksheet.writeComment(2, 2, "Hello");

    try worksheet.writeString(
        5,
        2,
        "Move the mouse over this cell and you will see 'Cell C6 " ++
            "commented by libxlsxwriter' in the status bar at the bottom.",
        text_wrap,
    );

    const options7a = xwz.CommentOptions{
        .author = "libxlsxwriter",
    };
    try worksheet.writeCommentOpt(5, 2, "Hello", options7a);

    // Example 8. Demonstrates the need to explicitly set the row height.
    worksheet = worksheet8;
    try worksheet.setColumn(2, 2, 25, .none);
    try worksheet.setRow(2, 80, .none);

    worksheet.showComments();

    var cell = xwz.cell("C3");
    try worksheet.writeString(
        cell.row,
        cell.col,
        "The height of this row has been adjusted explicitly using " ++
            "worksheet_set_row(). The size of the comment box is " ++
            "adjusted accordingly by libxlsxwriter",
        text_wrap,
    );
    try worksheet.writeComment(cell.row, cell.col, "Hello");

    cell = xwz.cell("C6");
    try worksheet.writeString(
        cell.row,
        cell.col,
        "The height of this row has been adjusted by Excel when the " ++
            "file is opened due to the text wrap property being set. " ++
            "Unfortunately this means that the height of the row is " ++
            "unknown to libxlsxwriter at run time and thus the comment " ++
            "box is stretched as well.\n\n" ++
            "Use worksheet_set_row() to specify the row height explicitly " ++
            "to avoid this problem.",
        text_wrap,
    );
    try worksheet.writeComment(cell.row, cell.col, "Hello");
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
const WorkSheet = xwz.WorkSheet;
