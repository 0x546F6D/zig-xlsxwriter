pub inline fn nameToRow(cell_name: [*c]const u8) u32 {
    return c.lxw_name_to_row(cell_name);
}

pub inline fn nameToRow2(cell_name: [*c]const u8) u32 {
    return c.lxw_name_to_row_2(cell_name);
}

pub inline fn nameToCol(cell_name: [*c]const u8) u16 {
    return c.lxw_name_to_col(cell_name);
}

pub inline fn nameToCol2(cell_name: [*c]const u8) u16 {
    return c.lxw_name_to_col_2(cell_name);
}

const c = @import("xlsxwriter_c");
