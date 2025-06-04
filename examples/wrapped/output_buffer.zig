var output_buffer: [*:0]const u8 = undefined;
var output_buffer_size: usize = undefined;

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a new workbook with options
    const workbook = try xwz.initWorkBook(
        null,
        null,
        .{
            .output_buffer = &output_buffer,
            .output_buffer_size = &output_buffer_size,
        },
    );
    const worksheet = try workbook.addWorkSheet(null);

    try worksheet.writeString(.{}, "Hello", .default);
    try worksheet.writeNumber(.{ .row = 1 }, 123, .default);

    try workbook.deinit();

    // Do something with the XLSX data in the output buffer.
    const file = try std.fs.createFileAbsoluteZ(xlsx_path, .{});
    defer file.close();

    try file.writeAll(output_buffer[0..output_buffer_size]);

    // Free the buffer
    std.c.free(@constCast(output_buffer));
}

const std = @import("std");
var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
