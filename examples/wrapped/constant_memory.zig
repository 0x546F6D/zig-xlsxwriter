const row_max: u32 = 1000;
const col_max: u16 = 50;

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a new workbook with options
    const workbook = try xwz.initWorkBook(null, xlsx_path, .{ .constant_memory = true });
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    var row: u32 = 0;
    while (row < row_max) : (row += 1) {
        var col: u16 = 0;
        while (col < col_max) : (col += 1) {
            try worksheet.writeNumber(.{ .row = row, .col = col }, 123.45, .default);
        }
    }
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
