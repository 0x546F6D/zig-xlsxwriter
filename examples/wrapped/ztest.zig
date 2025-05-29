pub fn main() !void {

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook("out/ztest.xlsx");
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);
    _ = worksheet;

    std.debug.print("error num = {}\n", .{@intFromError(xwz.XlsxError.ZipParameterError)});
    std.debug.print("error string = {s}\n", .{xwz.strError(xwz.XlsxError.SheetnameStartEndApostrophe)});
    std.debug.print("error string = {s}\n", .{xwz.strError(xwz.XlsxError.AddChart)});
}

const std = @import("std");
const xwz = @import("xlsxwriter");
