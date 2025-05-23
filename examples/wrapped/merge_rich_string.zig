pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add worksheets.
    const workbook = try xlsxwriter.initWorkBook(xlsx_path.ptr);
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
    const fragment1: xlsxwriter.RichStringTuple = .{
        .string = "This is ",
    };
    const fragment2: xlsxwriter.RichStringTuple = .{
        .format_c = red.format_c,
        .string = "red",
    };
    const fragment3: xlsxwriter.RichStringTuple = .{
        .string = " and this is ",
    };
    const fragment4: xlsxwriter.RichStringTuple = .{
        .format_c = blue.format_c,
        .string = "blue",
    };

    const rich_string: xlsxwriter.RichStringType = &.{
        &fragment1,
        &fragment2,
        &fragment3,
        &fragment4,
        null,
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
const xlsxwriter = @import("xlsxwriter");
const WorkSheet = @import("xlsxwriter").WorkSheet;
const Format = @import("xlsxwriter").Format;
