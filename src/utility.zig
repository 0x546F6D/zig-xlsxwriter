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

pub const Cell = struct { row: u32, col: u16 };
pub inline fn cell(cell_name: [*c]const u8) Cell {
    return Cell{
        .row = nameToRow(cell_name),
        .col = nameToCol(cell_name),
    };
}

pub const Cols = struct { first_col: u16, last_col: u16 };
pub inline fn cols(cell_name: [*c]const u8) Cols {
    return Cols{
        .first_col = nameToCol(cell_name),
        .last_col = nameToCol2(cell_name),
    };
}

pub const Range = struct { first_row: u32, first_col: u16, last_row: u32, last_col: u16 };
pub inline fn range(cell_name: [*c]const u8) Range {
    return Range{
        .first_row = nameToRow(cell_name),
        .first_col = nameToCol(cell_name),
        .last_row = nameToRow2(cell_name),
        .last_col = nameToCol2(cell_name),
    };
}

// pub extern fn lxw_get_filehandle(buf: [*c][*c]u8, size: [*c]usize, tmpdir: [*c]const u8) [*c]FILE;
// pub inline fn getFileHandle (buf: [*c][*c]u8, size: [*c]usize, tmpdir: [*c]const u8) [*c]FILE

const c = @import("xlsxwriter_c");
