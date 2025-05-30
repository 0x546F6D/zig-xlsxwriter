pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);

    // Add a worksheet with a user defined sheet name
    const worksheet1 = try workbook.addWorkSheet("Demo");

    // Add a worksheet with Excel's default sheet name: Sheet2
    const worksheet2 = try workbook.addWorkSheet(null);

    // Add some cell formats
    const myformat1 = try workbook.addFormat();
    const myformat2 = try workbook.addFormat();

    // Set the bold property for the first format
    myformat1.setBold();

    // Set a number format for the second format
    myformat2.setNumFormat("$#,##0.00");

    // Widen the first column to make the text clearer
    try worksheet1.setColumn(0, 0, 20, .default);

    // Write some unformatted data
    try worksheet1.writeString(0, 0, "Peach", .default);
    try worksheet1.writeString(1, 0, "Plum", .default);

    // Write formatted data
    try worksheet1.writeString(2, 0, "Pear", myformat1);

    // Formats can be reused
    try worksheet1.writeString(3, 0, "Persimmon", myformat1);

    // Write some numbers
    try worksheet1.writeNumber(5, 0, 123, .default);
    try worksheet1.writeNumber(6, 0, 4567.555, myformat2);

    // Write to the second worksheet
    try worksheet2.writeString(0, 0, "Some text", myformat1);

    // Close the workbook, save the file and free any memory
    workbook.deinit() catch |err| {
        std.debug.print("Error in workbook.deinit().\nError = {s}\n", .{xwz.strError(err)});
        return err;
    };
}

const std = @import("std");
var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
