const WorkBook = @This();

alloc: std.mem.Allocator,
workbook_c: ?*c.lxw_workbook,
worksheets: ?[]WorkSheet = null,

// pub extern fn workbook_new(filename: [*c]const u8) [*c]lxw_workbook;
pub inline fn new(alloc: std.mem.Allocator, filename: ?CString) XlsxError!WorkBook {
    return WorkBook{
        .alloc = alloc,
        .workbook_c = c.workbook_new(filename) orelse return XlsxError.NewWorkBook,
    };
}

// pub const WorkBookOutputBuffer = [*:0]const u8;
// c.lxw_workbook_options
pub const WorkBookOptions = struct {
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
            .output_buffer = @ptrCast(self.output_buffer),
            .output_buffer_size = self.output_buffer_size,
        };
    }
};

// pub extern fn workbook_new_opt(filename: [*c]const u8, options: [*c]lxw_workbook_options) [*c]lxw_workbook;
pub inline fn newOpt(
    alloc: std.mem.Allocator,
    filename: ?CString,
    options: WorkBookOptions,
) XlsxError!WorkBook {
    return WorkBook{
        .alloc = alloc,
        .workbook_c = c.workbook_new_opt(
            filename,
            @constCast(&options.toC()),
        ) orelse return XlsxError.NewWorkBook,
    };
}

// pub extern fn workbook_close(workbook: [*c]lxw_workbook) lxw_error;
pub inline fn deinit(self: WorkBook) XlsxError!void {
    if (self.worksheets) |worksheets| self.alloc.free(worksheets);
    try check(c.workbook_close(self.workbook_c));
}

// pub extern fn workbook_add_worksheet(workbook: [*c]lxw_workbook, sheetname: [*c]const u8) [*c]lxw_worksheet;
pub inline fn addWorkSheet(self: WorkBook, sheetname: ?CString) XlsxError!WorkSheet {
    return WorkSheet{
        .alloc = self.alloc,
        .worksheet_c = c.workbook_add_worksheet(self.workbook_c, sheetname) orelse return XlsxError.AddWorkSheet,
    };
}

// pub extern fn workbook_add_format(workbook: [*c]lxw_workbook) [*c]lxw_format;
pub inline fn addFormat(self: WorkBook) XlsxError!Format {
    return Format{
        .format_c = c.workbook_add_format(self.workbook_c) orelse return XlsxError.AddFormat,
    };
}

// pub extern fn workbook_add_chart(workbook: [*c]lxw_workbook, chart_type: u8) [*c]lxw_chart;
pub inline fn addChart(self: WorkBook, chart_type: Chart.Type) XlsxError!Chart {
    return Chart{
        .chart_c = c.workbook_add_chart(self.workbook_c, @intFromEnum(chart_type)) orelse return XlsxError.AddChart,
    };
}

// pub extern fn workbook_get_default_url_format(workbook: [*c]lxw_workbook) [*c]lxw_format;
pub inline fn getDefaultUrlFormat(self: WorkBook) XlsxError!Format {
    return Format{
        .format_c = c.workbook_get_default_url_format(self.workbook_c) orelse return XlsxError.AddFormat,
    };
}

// pub extern fn workbook_get_worksheet_by_name(workbook: [*c]lxw_workbook, name: [*c]const u8) [*c]lxw_worksheet;
pub inline fn getWorkSheetByName(self: WorkBook, name: ?CString) WorkSheet {
    return WorkSheet{
        .alloc = self.alloc,
        .worksheet_c = c.workbook_get_worksheet_by_name(self.workbook_c, name),
    };
}

// Get the list of worksheets
pub inline fn getWorkSheets(self: *WorkBook) !?[]WorkSheet {
    // clean-up previous list of worksheets
    if (self.worksheets) |worksheets| {
        self.alloc.free(worksheets);
        self.worksheets = null;
    }
    const nb_worksheet = self.workbook_c.?.num_worksheets;
    if (nb_worksheet == 0) return null;

    var worksheet_array = try self.alloc.alloc(WorkSheet, nb_worksheet);

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

// pub extern fn workbook_define_name(workbook: [*c]lxw_workbook, name: [*c]const u8, formula: [*c]const u8) lxw_error;
pub inline fn defineName(
    self: WorkBook,
    name: ?CString,
    formula: ?CString,
) XlsxError!void {
    try check(c.workbook_define_name(
        self.workbook_c,
        name,
        formula,
    ));
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

// pub extern fn workbook_set_custom_property_string(workbook: [*c]lxw_workbook, name: [*c]const u8, value: [*c]const u8) lxw_error;
pub inline fn setCustomPropertyString(
    self: WorkBook,
    name: ?CString,
    value: ?CString,
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
    name: ?CString,
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
    name: ?CString,
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
    name: ?CString,
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
    name: ?CString,
    value: c.lxw_datetime,
) XlsxError!void {
    try check(c.workbook_set_custom_property_datetime(
        self.workbook_c,
        name,
        @constCast(&value),
    ));
}

// pub extern fn workbook_add_chartsheet(workbook: [*c]lxw_workbook, sheetname: [*c]const u8) [*c]lxw_chartsheet;

// pub extern fn workbook_get_worksheet_by_name(workbook: [*c]lxw_workbook, name: [*c]const u8) [*c]lxw_worksheet;
// pub extern fn workbook_get_chartsheet_by_name(workbook: [*c]lxw_workbook, name: [*c]const u8) [*c]lxw_chartsheet;
// pub extern fn workbook_validate_sheet_name(workbook: [*c]lxw_workbook, sheetname: [*c]const u8) lxw_error;
// pub extern fn workbook_add_vba_project(workbook: [*c]lxw_workbook, filename: [*c]const u8) lxw_error;
// pub extern fn workbook_add_signed_vba_project(workbook: [*c]lxw_workbook, vba_project: [*c]const u8, signature: [*c]const u8) lxw_error;
// pub extern fn workbook_set_vba_name(workbook: [*c]lxw_workbook, name: [*c]const u8) lxw_error;
// pub extern fn workbook_read_only_recommended(workbook: [*c]lxw_workbook) void;
// pub extern fn workbook_set_size(workbook: [*c]lxw_workbook, width: u16, height: u16) void;
// pub extern fn workbook_unset_default_url_format(workbook: [*c]lxw_workbook) void;

test "assembling a complete Workbook file." {

    // The Standard Library contains useful functions to help create tests.
    // `expect` is a function that verifies its argument is true.
    // It will return an error if its argument is false to indicate a failure.
    // `try` is used to return an error to the test runner to notify it that the test failed.
    try std.testing.expect(42 == 42);
}

const std = @import("std");
const c = @import("lxw");
const CString = @import("xlsxwriter.zig").CString;
const CStringArray = @import("xlsxwriter.zig").CStringArray;
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;
const WorkSheet = @import("WorkSheet.zig");
const Format = @import("Format.zig");
const Chart = @import("Chart.zig");
