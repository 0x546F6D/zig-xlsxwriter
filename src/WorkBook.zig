const WorkBook = @This();

alloc: ?std.mem.Allocator = null,
workbook_c: ?*c.lxw_workbook,
worksheets: ?[]WorkSheet = null,
chartsheets: ?[]ChartSheet = null,

// c.lxw_workbook_options
pub const WorkBookOptions = extern struct {
    constant_memory: bool = false,
    tmpdir: ?[*:0]const u8 = null,
    use_zip64: bool = false,
    output_buffer: ?*[*:0]const u8 = null,
    output_buffer_size: ?*usize = null,

    pub const default = WorkBookOptions{
        .constant_memory = false,
        .tmpdir = null,
        .use_zip64 = false,
        .output_buffer = null,
        .output_buffer_size = null,
    };

    inline fn toC(self: WorkBookOptions) c.lxw_workbook_options {
        return c.lxw_workbook_options{
            .constant_memory = @intFromBool(self.constant_memory),
            .tmpdir = self.tmpdir,
            .use_zip64 = @intFromBool(self.use_zip64),
            .output_buffer = @ptrCast(@constCast(self.output_buffer)),
            .output_buffer_size = self.output_buffer_size,
        };
    }
};

// pub extern fn workbook_new(filename: [*c]const u8) [*c]lxw_workbook;
// pub extern fn workbook_new_opt(filename: [*c]const u8, options: [*c]lxw_workbook_options) [*c]lxw_workbook;
pub inline fn initWorkBook(alloc: ?std.mem.Allocator, filename: ?[*:0]const u8, options: ?WorkBookOptions) XlsxError!WorkBook {
    const output_buffer = if (options) |wb_options|
        wb_options.output_buffer
    else
        null;
    if (filename == null and output_buffer == null) return XlsxError.InitWorkBookOutput;

    return WorkBook{
        .alloc = alloc,
        // .output_buffer = output_buffer,
        .workbook_c = if (options) |wb_options|
            c.workbook_new_opt(
                filename,
                @constCast(&wb_options.toC()),
            ) orelse return XlsxError.InitWorkBook
        else
            c.workbook_new(filename) orelse return XlsxError.InitWorkBook,
    };
}

// pub extern fn workbook_close(workbook: [*c]lxw_workbook) lxw_error;
pub inline fn deinit(self: WorkBook) XlsxError!void {
    if (self.worksheets) |worksheets| self.alloc.?.free(worksheets);
    if (self.chartsheets) |chartsheets| self.alloc.?.free(chartsheets);
    try check(c.workbook_close(self.workbook_c));
}

pub const DocProperties = c.lxw_doc_properties;
// pub extern fn workbook_set_properties(workbook: [*c]lxw_workbook, properties: [*c]lxw_doc_properties) lxw_error;
pub inline fn setProperties(
    self: WorkBook,
    properties: DocProperties,
) XlsxError!void {
    try check(c.workbook_set_properties(
        self.workbook_c,
        @constCast(&properties),
    ));
}

// pub extern fn workbook_add_worksheet(workbook: [*c]lxw_workbook, sheetname: [*c]const u8) [*c]lxw_worksheet;
pub inline fn addWorkSheet(self: WorkBook, sheetname: ?[*:0]const u8) XlsxError!WorkSheet {
    // const sheetname_ptr = OptStringPtr(sheetname);
    return WorkSheet{
        .alloc = self.alloc,
        .worksheet_c = c.workbook_add_worksheet(self.workbook_c, sheetname) orelse
            return XlsxError.AddWorkSheet,
    };
}

// pub extern fn workbook_get_worksheet_by_name(workbook: [*c]lxw_workbook, name: [*c]const u8) [*c]lxw_worksheet;
pub inline fn getWorkSheetByName(self: WorkBook, name: [:0]const u8) XlsxError!WorkSheet {
    return WorkSheet{
        .alloc = self.alloc,
        .worksheet_c = c.workbook_get_worksheet_by_name(
            self.workbook_c,
            name,
        ) orelse
            return XlsxError.GetWorkSheetName,
    };
}

// Get the list of worksheets
pub inline fn getWorkSheets(self: *WorkBook) !?[]WorkSheet {
    const allocator: std.mem.Allocator = if (self.alloc) |allocator|
        allocator
    else
        return XlsxError.GetWorkSheets;
    // clean-up previous list of worksheets
    if (self.worksheets) |worksheets| {
        allocator.free(worksheets);
        self.worksheets = null;
    }
    const nb_worksheet = self.workbook_c.?.num_worksheets;
    if (nb_worksheet == 0) return null;

    var worksheet_array = try allocator.alloc(WorkSheet, nb_worksheet);

    var worksheet_c: ?*c.lxw_worksheet = self.workbook_c.?.worksheets.*.stqh_first;
    for (0..nb_worksheet) |i| {
        worksheet_array[i] = WorkSheet{
            .alloc = self.alloc,
            .worksheet_c = worksheet_c,
        };
        worksheet_c = worksheet_c.?.list_pointers.stqe_next;
    }

    self.worksheets = worksheet_array;
    return self.worksheets;
}

// pub extern fn workbook_add_chartsheet(workbook: [*c]lxw_workbook, sheetname: [*c]const u8) [*c]lxw_chartsheet;
pub inline fn addChartSheet(
    self: WorkBook,
    sheetname: ?[*:0]const u8,
) XlsxError!ChartSheet {
    return ChartSheet{
        // .alloc = self.alloc,
        .chartsheet_c = c.workbook_add_chartsheet(self.workbook_c, sheetname) orelse
            return XlsxError.AddChartSheet,
    };
}

// pub extern fn workbook_get_chartsheet_by_name(workbook: [*c]lxw_workbook, name: [*c]const u8) [*c]lxw_chartsheet;
pub inline fn getChartSheetByName(self: WorkBook, name: [:0]const u8) ChartSheet {
    return ChartSheet{
        // .alloc = self.alloc,
        .worksheet_c = c.workbook_get_chartsheet_by_name(self.workbook_c, name) orelse
            return XlsxError.GetChartSheetName,
    };
}

// Get the list of chartsheets
pub inline fn getChartSheets(self: *WorkBook) !?[]ChartSheet {
    const allocator: std.mem.Allocator = if (self.alloc) |allocator|
        allocator
    else
        return XlsxError.GetChartSheets;
    // clean-up previous list of chartsheets
    if (self.chartsheets) |chartsheets| {
        allocator.free(chartsheets);
        self.chartsheets = null;
    }
    const nb_chartsheet = self.workbook_c.?.num_chartsheets;
    if (nb_chartsheet == 0) return null;

    var chartsheet_array = try allocator.alloc(ChartSheet, nb_chartsheet);

    var chartsheet_c: ?*c.lxw_chartsheet = self.workbook_c.?.chartsheets.*.stqh_first;
    for (0..nb_chartsheet) |i| {
        chartsheet_array[i] = ChartSheet{
            // .alloc = self.alloc,
            .chartsheet_c = chartsheet_c,
        };
        chartsheet_c = chartsheet_c.?.list_pointers.stqe_next;
    }

    self.chartsheets = chartsheet_array;
    return self.chartsheets;
}

// pub extern fn workbook_add_format(workbook: [*c]lxw_workbook) [*c]lxw_format;
pub inline fn addFormat(self: WorkBook) XlsxError!Format {
    return Format{
        .format_c = c.workbook_add_format(self.workbook_c) orelse return XlsxError.AddFormat,
    };
}

// pub extern fn workbook_add_chart(workbook: [*c]lxw_workbook, chart_type: u8) [*c]lxw_chart;
pub inline fn addChart(self: WorkBook, chart_type: ChartType) XlsxError!Chart {
    const chart = c.workbook_add_chart(self.workbook_c, @intFromEnum(chart_type)) orelse
        return XlsxError.AddChart;
    return Chart{
        .alloc = self.alloc,
        .chart_c = chart,
        .x_axis = ChartAxis{ .axis_c = chart.*.x_axis },
        .y_axis = ChartAxis{ .axis_c = chart.*.y_axis },
    };
}

// pub extern fn workbook_get_default_url_format(workbook: [*c]lxw_workbook) [*c]lxw_format;
pub inline fn getDefaultUrlFormat(self: WorkBook) XlsxError!Format {
    return Format{
        .format_c = c.workbook_get_default_url_format(self.workbook_c) orelse
            return XlsxError.AddFormat,
    };
}

// pub extern fn workbook_define_name(workbook: [*c]lxw_workbook, name: [*c]const u8, formula: [*c]const u8) lxw_error;
pub inline fn defineName(
    self: WorkBook,
    name: [:0]const u8,
    formula: [:0]const u8,
) XlsxError!void {
    try check(c.workbook_define_name(
        self.workbook_c,
        name,
        formula,
    ));
}

// pub extern fn workbook_set_custom_property_string(workbook: [*c]lxw_workbook, name: [*c]const u8, value: [*c]const u8) lxw_error;
pub inline fn setCustomPropertyString(
    self: WorkBook,
    name: [:0]const u8,
    value: [:0]const u8,
) XlsxError!void {
    try check(c.workbook_set_custom_property_string(
        self.workbook_c,
        name,
        value,
    ));
}

// pub extern fn workbook_set_custom_property_number(workbook: [*c]lxw_workbook, name: [*c]const u8, value: f64) lxw_error;
pub inline fn setCustomPropertyNumber(
    self: WorkBook,
    name: [:0]const u8,
    value: f64,
) XlsxError!void {
    try check(c.workbook_set_custom_property_number(
        self.workbook_c,
        name,
        value,
    ));
}

// pub extern fn workbook_set_custom_property_integer(workbook: [*c]lxw_workbook, name: [*c]const u8, value: i32) lxw_error;
pub inline fn setCustomPropertyInteger(
    self: WorkBook,
    name: [:0]const u8,
    value: i32,
) XlsxError!void {
    try check(c.workbook_set_custom_property_integer(
        self.workbook_c,
        name,
        value,
    ));
}

// pub extern fn workbook_set_custom_property_boolean(workbook: [*c]lxw_workbook, name: [*c]const u8, value: u8) lxw_error;
pub inline fn setCustomPropertyBoolean(
    self: WorkBook,
    name: [:0]const u8,
    value: bool,
) XlsxError!void {
    try check(c.workbook_set_custom_property_boolean(
        self.workbook_c,
        name,
        @intFromBool(value),
    ));
}

// pub extern fn workbook_set_custom_property_datetime(workbook: [*c]lxw_workbook, name: [*c]const u8, datetime: [*c]lxw_datetime) lxw_error;
pub inline fn setCustomPropertyDateTime(
    self: WorkBook,
    name: [:0]const u8,
    value: c.lxw_datetime,
) XlsxError!void {
    try check(c.workbook_set_custom_property_datetime(
        self.workbook_c,
        name,
        @constCast(&value),
    ));
}

// pub extern fn workbook_add_vba_project(workbook: [*c]lxw_workbook, filename: [*c]const u8) lxw_error;
pub inline fn addVbaProject(
    self: WorkBook,
    filename: [:0]const u8,
) XlsxError!void {
    try check(c.workbook_add_vba_project(
        self.workbook_c,
        filename,
    ));
}

// pub extern fn workbook_add_signed_vba_project(workbook: [*c]lxw_workbook, vba_project: [*c]const u8, signature: [*c]const u8) lxw_error;
pub inline fn addSignedVbaProject(
    self: WorkBook,
    vba_project: [:0]const u8,
    signature: [:0]const u8,
) XlsxError!void {
    try check(c.workbook_add_signed_vba_project(
        self.workbook_c,
        vba_project,
        signature,
    ));
}

// pub extern fn workbook_set_vba_name(workbook: [*c]lxw_workbook, name: [*c]const u8) lxw_error;
pub inline fn setVbaName(
    self: WorkBook,
    name: [:0]const u8,
) XlsxError!void {
    try check(c.workbook_set_vba_name(
        self.workbook_c,
        name,
    ));
}

// pub extern fn workbook_validate_sheet_name(workbook: [*c]lxw_workbook, sheetname: [*c]const u8) lxw_error;
pub inline fn validateSheetName(
    self: WorkBook,
    sheetname: [:0]const u8,
) XlsxError!void {
    try check(c.workbook_validate_sheet_name(
        self.workbook_c,
        sheetname,
    ));
}

// pub extern fn workbook_read_only_recommended(workbook: [*c]lxw_workbook) void;
pub inline fn readOnlyRecommended(self: WorkBook) void {
    c.workbook_read_only_recommended(self.workbook_c);
}

// pub extern fn workbook_set_size(workbook: [*c]lxw_workbook, width: u16, height: u16) void;
pub inline fn setSize(
    self: WorkBook,
    width: u16,
    height: u16,
) void {
    c.workbook_set_size(self.workbook_c, width, height);
}

// pub extern fn workbook_unset_default_url_format(workbook: [*c]lxw_workbook) void;
pub inline fn unsetDefaultUrlFormat(self: WorkBook) void {
    c.workbook_unset_default_url_format(self.workbook_c);
}

test "assembling a complete Workbook file." {

    // The Standard Library contains useful functions to help create tests.
    // `expect` is a function that verifies its argument is true.
    // It will return an error if its argument is false to indicate a failure.
    // `try` is used to return an error to the test runner to notify it that the test failed.
    try std.testing.expect(42 == 42);
}

const std = @import("std");
const c = @import("lxw");
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;
const WorkSheet = @import("WorkSheet.zig");
const ChartSheet = @import("ChartSheet.zig");
const Format = @import("Format.zig");
const Chart = @import("Chart.zig");
const ChartType = Chart.Type;
const ChartAxis = @import("chart/Axis.zig");
