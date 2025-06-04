// Embed the VBA project binary
const vba_data = @embedFile("vbaProject.bin");

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsm_path = try h.getXlsmPath(alloc, @src().file);
    defer alloc.free(xlsm_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsm_path, null);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // get vbaProject path
    const asset_path = try h.getAssetPath(alloc, "vbaProject.bin");
    defer alloc.free(asset_path);

    try worksheet.setColumn(.{}, 30, .default, null);

    // Add a macro file extracted from an Excel workbook
    try workbook.addVbaProject(asset_path);

    try worksheet.writeString(.{ .row = 2 }, "Press the button to say hello.", .default);

    const options = xwz.ButtonOptions{
        .caption = "Press Me",
        .macro = "say_hello",
        .width = 80,
        .height = 30,
    };

    try worksheet.insertButton(.{ .row = 2, .col = 1 }, options);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
