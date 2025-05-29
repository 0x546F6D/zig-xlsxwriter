// An example of inserting images into a worksheet using the libxlsxwriter
// library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//

const std = @import("std");
const lxw = @import("lxw");
const mktmp = @import("mktmp");

// Embed the logo image directly into the executable
const logo_data = @embedFile("logo.png");

pub fn main() !void {
    // Create a temporary file for the logo using the TmpFile API
    var arena = std.heap.ArenaAllocator.init(
        std.heap.page_allocator,
    );
    defer arena.deinit();
    const allocator = arena.allocator();

    var tmp_file = try mktmp.TmpFile.create(
        allocator,
        "logo_",
    );
    defer tmp_file.cleanUp();

    // Write the embedded data to the temporary file
    try tmp_file.write(logo_data);

    // Create a new workbook and add a worksheet.
    const workbook = lxw.workbook_new("zig-images.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Change some of the column widths for clarity.
    // COLS("A:A") expands to first_col, last_col
    const first_col = lxw.lxw_name_to_col("A:A");
    const last_col = lxw.lxw_name_to_col_2("A:A");
    _ = lxw.worksheet_set_column(worksheet, first_col, last_col, 30, null);

    // Insert an image.
    // CELL("A2") expands to row, col
    const row_a2 = lxw.lxw_name_to_row("A2");
    const col_a2 = lxw.lxw_name_to_col("A2");
    _ = lxw.worksheet_write_string(worksheet, row_a2, col_a2, "Insert an image in a cell:", null);

    const row_b2 = lxw.lxw_name_to_row("B2");
    const col_b2 = lxw.lxw_name_to_col("B2");
    _ = lxw.worksheet_insert_image(worksheet, row_b2, col_b2, tmp_file.path.ptr);

    // Insert an image offset in the cell.
    const row_a12 = lxw.lxw_name_to_row("A12");
    const col_a12 = lxw.lxw_name_to_col("A12");
    _ = lxw.worksheet_write_string(worksheet, row_a12, col_a12, "Insert an offset image:", null);

    var options1 = lxw.lxw_image_options{
        .x_offset = 15,
        .y_offset = 10,
        .x_scale = 0,
        .y_scale = 0,
        .object_position = 0,
        .description = null,
        .url = null,
        .tip = null,
    };

    const row_b12 = lxw.lxw_name_to_row("B12");
    const col_b12 = lxw.lxw_name_to_col("B12");
    _ = lxw.worksheet_insert_image_opt(worksheet, row_b12, col_b12, tmp_file.path.ptr, &options1);

    // Insert an image with scaling.
    const row_a22 = lxw.lxw_name_to_row("A22");
    const col_a22 = lxw.lxw_name_to_col("A22");
    _ = lxw.worksheet_write_string(worksheet, row_a22, col_a22, "Insert a scaled image:", null);

    var options2 = lxw.lxw_image_options{
        .x_offset = 0,
        .y_offset = 0,
        .x_scale = 0.5,
        .y_scale = 0.5,
        .object_position = 0,
        .description = null,
        .url = null,
        .tip = null,
    };

    const row_b22 = lxw.lxw_name_to_row("B22");
    const col_b22 = lxw.lxw_name_to_col("B22");
    _ = lxw.worksheet_insert_image_opt(worksheet, row_b22, col_b22, tmp_file.path.ptr, &options2);

    // Insert an image with a hyperlink.
    const row_a32 = lxw.lxw_name_to_row("A32");
    const col_a32 = lxw.lxw_name_to_col("A32");
    _ = lxw.worksheet_write_string(worksheet, row_a32, col_a32, "Insert an image with a hyperlink:", null);

    var options3 = lxw.lxw_image_options{
        .x_offset = 0,
        .y_offset = 0,
        .x_scale = 0,
        .y_scale = 0,
        .object_position = 0,
        .description = null,
        .url = "https://github.com/jmcnamara",
        .tip = null,
    };

    const row_b32 = lxw.lxw_name_to_row("B32");
    const col_b32 = lxw.lxw_name_to_col("B32");
    _ = lxw.worksheet_insert_image_opt(worksheet, row_b32, col_b32, tmp_file.path.ptr, &options3);

    _ = lxw.workbook_close(workbook);
}
