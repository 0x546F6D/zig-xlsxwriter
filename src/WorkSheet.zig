const WorkSheet = @This();

worksheet_c: *c.lxw_worksheet,

pub inline fn setColumn(
    self: WorkSheet,
    first_col: u16,
    last_col: u16,
    width: f64,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_set_column(
        self.worksheet_c,
        first_col,
        last_col,
        width,
        format.format_c,
    ));
}

pub inline fn writeString(
    self: WorkSheet,
    row: u32,
    col: u16,
    string: [*c]const u8,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_string(
        self.worksheet_c,
        row,
        col,
        string,
        format.format_c,
    ));
}

pub inline fn writeNumber(
    self: WorkSheet,
    row: u32,
    col: u16,
    number: f64,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_number(
        self.worksheet_c,
        row,
        col,
        number,
        format.format_c,
    ));
}

pub inline fn writeDateTime(
    self: WorkSheet,
    row: u32,
    col: u16,
    datetime: *c.lxw_datetime,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_datetime(
        self.worksheet_c,
        row,
        col,
        datetime,
        format.format_c,
    ));
}

pub inline fn writeFormula(
    self: WorkSheet,
    row: u32,
    col: u16,
    formula: [*c]const u8,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_formula(
        self.worksheet_c,
        row,
        col,
        formula,
        format.format_c,
    ));
}

pub inline fn writeArrayFormula(
    self: WorkSheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    formula: [*c]const u8,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_array_formula(
        self.worksheet_c,
        first_row,
        first_col,
        last_row,
        last_col,
        formula,
        format.format_c,
    ));
}

pub inline fn insertChart(
    self: WorkSheet,
    row: u32,
    col: u16,
    chart: Chart,
) XlsxError!void {
    try check(c.worksheet_insert_chart(
        self.worksheet_c,
        row,
        col,
        chart.chart_c,
    ));
}

const c = @import("xlsxwriter_c");
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;
const Format = @import("Format.zig");
const Chart = @import("Chart.zig");
