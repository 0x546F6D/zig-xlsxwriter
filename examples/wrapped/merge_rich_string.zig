pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add worksheets.
    var workbook = try xwz.initWorkBook(alloc, xlsx_path.ptr);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    // Configure a format for the merged range.
    const merge_format = try workbook.addFormat();
    merge_format.setAlign(.center);
    merge_format.setAlign(.vertical_center);
    merge_format.setBorder(.thin);

    // Configure formats for the rich string.
    const red = try workbook.addFormat();
    red.setFontColor(.red);

    const blue = try workbook.addFormat();
    blue.setFontColor(.blue);
    // for comparison with original
    blue.setFontColor(@enumFromInt(0x0000FF));

    // Create the fragments for the rich string.
    const fragment1: xwz.RichStringTuple = .{
        .string = "This is ",
    };
    const fragment2: xwz.RichStringTuple = .{
        .format = red,
        .string = "red",
    };
    const fragment3: xwz.RichStringTuple = .{
        .string = " and this is ",
    };
    const fragment4: xwz.RichStringTuple = .{
        .format = blue,
        .string = "blue",
    };

    const rich_string = &.{
        fragment1,
        fragment2,
        fragment3,
        fragment4,
    };

    // Write an empty string to the merged range.
    try worksheet.mergeRange(1, 1, 4, 4, "", merge_format);

    // We then overwrite the first merged cell with a rich string. Note that
    // we must also pass the cell format used in the merged cells format at
    // the end.
    try worksheet.writeRichString(1, 1, rich_string, merge_format);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
