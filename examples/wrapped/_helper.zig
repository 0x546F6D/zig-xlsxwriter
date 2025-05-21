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
