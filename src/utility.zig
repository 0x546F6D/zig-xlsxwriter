pub inline fn nameToRow(cell_name: [*c]const u8) u32 {
    return c.lxw_name_to_row(cell_name);
}

// Convert the second row of an Excel range ref to a zero indexed number.
pub inline fn nameToRow2(cell_name: [*c]const u8) u32 {
    return c.lxw_name_to_row_2(cell_name);
}

pub inline fn nameToCol(cell_name: [*c]const u8) u16 {
    return c.lxw_name_to_col(cell_name);
}

// Convert the second column of an Excel range ref to a zero indexed number.
pub inline fn nameToCol2(cell_name: [*c]const u8) u16 {
    return c.lxw_name_to_col_2(cell_name);
}

// pub extern fn lxw_get_filehandle(buf: [*c][*c]u8, size: [*c]usize, tmpdir: [*c]const u8) [*c]FILE;
// pub inline fn getFileHandle (buf: [*c][*c]u8, size: [*c]usize, tmpdir: [*c]const u8) [*c]FILE

const c = @import("xlsxwriter_c");
