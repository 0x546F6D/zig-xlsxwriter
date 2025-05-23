const row_max: u32 = 1000;
const col_max: u16 = 50;

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Set the workbook options
    const options: xlsxwriter.WorkBookOptions = .{
        .constant_memory = xlsxwriter.xw_true,
        .tmpdir = null,
        .use_zip64 = xlsxwriter.xw_false,
        .output_buffer = null,
        .output_buffer_size = null,
    };

    // Create a new workbook with options
    const workbook = try xlsxwriter.initWorkBookOpt(xlsx_path.ptr, options);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    var row: u32 = 0;
    while (row < row_max) : (row += 1) {
        var col: u16 = 0;
        while (col < col_max) : (col += 1) {
            try worksheet.writeNumber(row, col, 123.45, .none);
        }
    }
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");
const WorkSheet = @import("xlsxwriter").WorkSheet;
const Format = @import("xlsxwriter").Format;
