pub fn main() !void {

    // Create a workbook and add a worksheet.
    const workbook = try xlsxwriter.initWorkBook("out/ztest.xlsx");
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);
    _ = worksheet;

    std.debug.print("error num = {}\n", .{@intFromError(xlsxwriter.XlsxError.ZipParameterError)});
    std.debug.print("error string = {s}\n", .{xlsxwriter.strError(xlsxwriter.XlsxError.SheetnameStartEndApostrophe)});
    std.debug.print("error string = {s}\n", .{xlsxwriter.strError(xlsxwriter.XlsxError.AddChart)});
}

const std = @import("std");
const xlsxwriter = @import("xlsxwriter");
