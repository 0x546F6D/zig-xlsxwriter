pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    const asset_path = try h.getAssetPath(alloc, "watermark.png");
    defer alloc.free(asset_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    // Set a worksheet header with the watermark image.
    const header_options = xwz.HeaderFooterOptions{ .image_center = asset_path };

    try worksheet.setHeaderOpt("&C&[Picture]", header_options);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
