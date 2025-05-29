// var output_buffer: xwz.WorkBookOutputBuffer = undefined;
var output_buffer: [*:0]const u8 = undefined;
// var output_buffer: [*c]const u8 = undefined;
var output_buffer_size: usize = undefined;

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Set the worksheet options.
    const options: xwz.WorkBookOptions = .{
        .constant_memory = true,
        .tmpdir = null,
        .use_zip64 = false,
        .output_buffer = &output_buffer,
        .output_buffer_size = &output_buffer_size,
    };

    // Create a new workbook with options
    var workbook = try xwz.initWorkBookOpt(alloc, null, options);
    const worksheet = try workbook.addWorkSheet(null);

    try worksheet.writeString(0, 0, "Hello", .none);
    try worksheet.writeNumber(1, 0, 123, .none);

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
