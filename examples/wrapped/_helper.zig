pub fn getXlsxPath(alloc: std.mem.Allocator, example_name: []const u8) ![:0]const u8 {
    const exe_dir = try std.fs.selfExeDirPathAlloc(alloc);
    defer alloc.free(exe_dir);

    const xlsx_name = try std.fmt.allocPrint(
        alloc,
        "{s}.xlsx",
        .{example_name[0 .. example_name.len - 4]},
    );
    defer alloc.free(xlsx_name);

    return try std.fs.path.joinZ(
        alloc,
        &[_][]const u8{ exe_dir, "..", "..", "examples", "wrapped", "out", xlsx_name },
    );
}

pub fn getAssetPath(alloc: std.mem.Allocator, asset_name: []const u8) ![:0]const u8 {
    const exe_dir = try std.fs.selfExeDirPathAlloc(alloc);
    defer alloc.free(exe_dir);

    return try std.fs.path.joinZ(
        alloc,
        &[_][]const u8{ exe_dir, "..", "..", "examples", "assets", asset_name },
    );
}

const std = @import("std");

// common replace:
// %s/xlsxwriter.workbook_add_format(workbook)/try workbook.addFormat()
// %s/_ = xlsxwriter.worksheet_set_column(\(\n *\)\=worksheet\d*, /try worksheet.setColumn(
// %s/_ = xlsxwriter.worksheet_set_column_opt(\(\n *\)\=worksheet\d*, /try worksheet.setColumnOpt(
// %s/_ = xlsxwriter.worksheet_set_row(\(\n *\)\=worksheet\d*, /try worksheet.setRow(
// %s/_ = xlsxwriter.worksheet_set_row_opt(\(\n *\)\=worksheet\d*, /try worksheet.setRowOpt(
// %s/_ = xlsxwriter.worksheet_write_string(\(\n *\)\=worksheet\d*,/try worksheet.writeString(
// %s/_ = xlsxwriter.worksheet_write_number(\(\n *\)\=worksheet\d*,/try worksheet.writeNumber(
// %s/_ = xlsxwriter.worksheet_write_formula(\(\n *\)\=worksheet\d*,/try worksheet.writeFormula(
//
//
//
//
//
//
//
//
