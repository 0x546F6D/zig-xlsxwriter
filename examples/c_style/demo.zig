//
// A simple example of some of the features of the libxlsxwriter library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
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
    const workbook = lxw.workbook_new("zig-demo.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Add a format.
    const format = lxw.workbook_add_format(workbook);

    // Set the bold property for the format
    _ = lxw.format_set_bold(format);

    // Change the column width for clarity.
    _ = lxw.worksheet_set_column(worksheet, 0, 0, 20, null);

    // Write some simple text.
    _ = lxw.worksheet_write_string(worksheet, 0, 0, "Hello", null);

    // Text with formatting.
    _ = lxw.worksheet_write_string(worksheet, 1, 0, "World", format);

    // Write some numbers.
    _ = lxw.worksheet_write_number(worksheet, 2, 0, 123, null);
    _ = lxw.worksheet_write_number(worksheet, 3, 0, 123.456, null);

    // Insert an image using the temporary file path
    // Convert the path to a null-terminated C string pointer
    const c_path = @as(
        [*c]const u8,
        @ptrCast(tmp_file.path.ptr),
    );
    _ = lxw.worksheet_insert_image(worksheet, 1, 2, c_path);

    _ = lxw.workbook_close(workbook);
}
