//
//
// An example of adding macros to a libxlsxwriter file using a VBA project
// file extracted from an existing Excel .xlsm file.
//
// The vba_extract.py utility from the libxlsxwriter examples directory can be
// used to extract the vbaProject.bin file.
//
// This example connects the macro to a button (the only Excel/VBA form object
// supported by libxlsxwriter) but that isn't a requirement for adding a macro
// file to the workbook.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//

const std = @import("std");
const lxw = @import("lxw");
const mktmp = @import("mktmp");

// Embed the VBA project binary
const vba_data = @embedFile("vbaProject.bin");

pub fn main() !void {
    // Create a temporary file for the VBA project using the TmpFile API
    var arena = std.heap.ArenaAllocator.init(std.heap.page_allocator);
    defer arena.deinit();
    const allocator = arena.allocator();

    var tmp_file = try mktmp.TmpFile.create(allocator, "vba_data");
    defer tmp_file.cleanUp();

    // Write the embedded data to the temporary file
    try tmp_file.write(vba_data);

    // Note the xlsm extension of the filename
    const workbook = lxw.workbook_new("zig-macro.xlsm");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    _ = lxw.worksheet_set_column(worksheet, 0, 0, 30, null);

    // Add a macro file extracted from an Excel workbook
    // Convert the path to a null-terminated C string pointer
    const c_path = @as([*c]const u8, @ptrCast(tmp_file.path.ptr));
    _ = lxw.workbook_add_vba_project(workbook, c_path);

    _ = lxw.worksheet_write_string(worksheet, 2, 0, "Press the button to say hello.", null);

    var options = lxw.lxw_button_options{
        .caption = "Press Me",
        .macro = "say_hello",
        .width = 80,
        .height = 30,
    };

    _ = lxw.worksheet_insert_button(worksheet, 2, 1, &options);

    _ = lxw.workbook_close(workbook);

    // The temporary file will be automatically cleaned up by the defer statement
}
