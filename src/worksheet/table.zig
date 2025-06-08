pub const TableStyleType = enum(u8) {
    default = c.LXW_TABLE_STYLE_TYPE_DEFAULT,
    light = c.LXW_TABLE_STYLE_TYPE_LIGHT,
    medium = c.LXW_TABLE_STYLE_TYPE_MEDIUM,
    dark = c.LXW_TABLE_STYLE_TYPE_DARK,
};
pub const TableTotalFunctions = enum(u8) {
    none = c.LXW_TABLE_FUNCTION_NONE,
    average = c.LXW_TABLE_FUNCTION_AVERAGE,
    count_nums = c.LXW_TABLE_FUNCTION_COUNT_NUMS,
    count = c.LXW_TABLE_FUNCTION_COUNT,
    max = c.LXW_TABLE_FUNCTION_MAX,
    min = c.LXW_TABLE_FUNCTION_MIN,
    std_dev = c.LXW_TABLE_FUNCTION_STD_DEV,
    sum = c.LXW_TABLE_FUNCTION_SUM,
    variance = c.LXW_TABLE_FUNCTION_VAR,
};

pub const TableColumn = struct {
    header: ?[*:0]const u8 = null,
    formula: ?[*:0]const u8 = null,
    total_string: ?[*:0]const u8 = null,
    total_function: TableTotalFunctions = .none,
    header_format: Format = .default,
    format: Format = .default,
    total_value: f64 = 0,

    pub const empty: TableColumn = .{
        .header = null,
        .formula = null,
        .total_string = null,
        .total_function = .none,
        .header_format = .default,
        .format = .default,
        .total_value = 0,
    };

    inline fn toC(self: TableColumn) c.lxw_table_column {
        return c.lxw_table_column{
            .header = self.header,
            .formula = self.formula,
            .total_string = self.total_string,
            .total_function = @intFromEnum(self.total_function),
            .header_format = self.header_format.format_c,
            .format = self.format.format_c,
            .total_value = self.total_value,
        };
    }
};

pub const TableOptions = struct {
    name: ?[*:0]const u8 = null,
    no_header_row: bool = false,
    no_autofilter: bool = false,
    no_banded_rows: bool = false,
    banded_columns: bool = false,
    first_column: bool = false,
    last_column: bool = false,
    style_type: TableStyleType = .default,
    style_type_number: u8 = 0,
    total_row: bool = false,
    columns: []const TableColumn = &.{},

    pub const default: TableOptions = .{
        .name = null,
        .no_header_row = false,
        .no_autofilter = false,
        .no_banded_rows = false,
        .banded_columns = false,
        .first_column = false,
        .last_column = false,
        .style_type = .default,
        .style_type_number = 0,
        .total_row = false,
        .columns = &.{},
    };

    inline fn toC(self: TableOptions, columns: [:null]const ?*c.lxw_table_column) c.lxw_table_options {
        return c.lxw_table_options{
            .name = self.name,
            .no_header_row = @intFromBool(self.no_header_row),
            .no_autofilter = @intFromBool(self.no_autofilter),
            .no_banded_rows = @intFromBool(self.no_banded_rows),
            .banded_columns = @intFromBool(self.banded_columns),
            .first_column = @intFromBool(self.first_column),
            .last_column = @intFromBool(self.last_column),
            .style_type = @intFromEnum(self.style_type),
            .style_type_number = self.style_type_number,
            .total_row = @intFromBool(self.total_row),
            .columns = @ptrCast(@constCast(columns)),
        };
    }
};

// pub extern fn worksheet_add_table(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, options: [*c]lxw_table_options) lxw_error;
pub inline fn addTable(
    self: WorkSheet,
    range: Range,
    options: TableOptions,
) !void {
    const allocator: std.mem.Allocator = if (self.alloc) |allocator|
        allocator
    else
        return XlsxError.AddTableNoAlloc;

    // convert table_columns to [*c][*c]lxw_table_column
    var table_column_array = try allocator.alloc(c.lxw_table_column, options.columns.len);
    defer allocator.free(table_column_array);
    var table_column_c = try allocator.allocSentinel(?*c.lxw_table_column, options.columns.len, null);
    defer allocator.free(table_column_c);

    for (options.columns, 0..) |column, i| {
        table_column_array[i] = column.toC();
        table_column_c[i] = &table_column_array[i];
    }

    try check(c.worksheet_add_table(
        self.worksheet_c,
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
        @constCast(&options.toC(table_column_c)),
    ));
}

pub const TableColumnNoAlloc = extern struct {
    header: ?[*:0]const u8 = null,
    formula: ?[*:0]const u8 = null,
    total_string: ?[*:0]const u8 = null,
    total_function: TableTotalFunctions = .none,
    header_format_c: ?*c.lxw_format = null,
    format_c: ?*c.lxw_format = null,
    total_value: f64 = 0,

    pub const default: TableColumnNoAlloc = .{
        .header = null,
        .formula = null,
        .total_string = null,
        .total_function = .none,
        .header_format_c = null,
        .format_c = null,
        .total_value = 0,
    };
};
pub const TableColumnNoAllocArray = [:null]const ?*const TableColumnNoAlloc;

pub const TableOptionsNoAlloc = struct {
    name: ?[*:0]const u8 = null,
    no_header_row: Bool = Bool.false,
    no_autofilter: Bool = Bool.false,
    no_banded_rows: Bool = Bool.false,
    banded_columns: Bool = Bool.false,
    first_column: Bool = Bool.false,
    last_column: Bool = Bool.false,
    style_type: TableStyleType = .default,
    style_type_number: u8 = 0,
    total_row: Bool = Bool.false,
    columns: TableColumnNoAllocArray = &.{},

    pub const default: TableOptionsNoAlloc = .{
        .name = null,
        .no_header_row = Bool.false,
        .no_autofilter = Bool.false,
        .no_banded_rows = Bool.false,
        .banded_columns = Bool.false,
        .first_column = Bool.false,
        .last_column = Bool.false,
        .style_type = .default,
        .style_type_number = 0,
        .total_row = Bool.false,
        .columns = &.{},
    };

    inline fn toC(self: TableOptionsNoAlloc) c.lxw_table_options {
        return c.lxw_table_options{
            .name = self.name,
            .no_header_row = @intFromEnum(self.no_header_row),
            .no_autofilter = @intFromEnum(self.no_autofilter),
            .no_banded_rows = @intFromEnum(self.no_banded_rows),
            .banded_columns = @intFromEnum(self.banded_columns),
            .first_column = @intFromEnum(self.first_column),
            .last_column = @intFromEnum(self.last_column),
            .style_type = @intFromEnum(self.style_type),
            .style_type_number = self.style_type_number,
            .total_row = @intFromEnum(self.total_row),
            .columns = @ptrCast(@constCast(self.columns)),
        };
    }
};

// pub extern fn worksheet_add_table(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw _col_t, options: [*c]lxw_table_options) lxw_error;
pub inline fn addTableNoAlloc(
    self: WorkSheet,
    range: Range,
    options: TableOptionsNoAlloc,
) XlsxError!void {
    try check(c.worksheet_add_table(
        self.worksheet_c,
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
        @constCast(&options.toC()),
    ));
}

const std = @import("std");
const c = @import("lxw");
const xlsxwriter = @import("../xlsxwriter.zig");
const Bool = xlsxwriter.Boolean;
const WorkSheet = @import("../WorkSheet.zig");
const XlsxError = @import("../errors.zig").XlsxError;
const check = @import("../errors.zig").checkResult;
const utility = @import("../utility.zig");
const Cell = utility.Cell;
const Cols = utility.Cols;
const Range = utility.Range;
const Format = @import("../Format.zig");
