const WorkBook = @This();

workbook_c: ?*c.lxw_workbook,

// pub extern fn workbook_new(filename: [*c]const u8) [*c]lxw_workbook;
pub inline fn new(filename: [*c]const u8) XlsxError!WorkBook {
    return WorkBook{
        .workbook_c = c.workbook_new(filename) orelse return XlsxError.NewWorkBook,
    };
}

pub const WorkBookOptions = c.lxw_workbook_options;
// pub extern fn workbook_new_opt(filename: [*c]const u8, options: [*c]lxw_workbook_options) [*c]lxw_workbook;
pub inline fn newOpt(filename: [*c]const u8, options: WorkBookOptions) XlsxError!WorkBook {
    return WorkBook{
        .workbook_c = c.workbook_new_opt(filename, @ptrCast(@constCast(&options))) orelse return XlsxError.NewWorkBook,
    };
}

// pub extern fn workbook_close(workbook: [*c]lxw_workbook) lxw_error;
pub inline fn deinit(self: WorkBook) XlsxError!void {
    try check(c.workbook_close(self.workbook_c));
}

// pub extern fn workbook_add_worksheet(workbook: [*c]lxw_workbook, sheetname: [*c]const u8) [*c]lxw_worksheet;
pub inline fn addWorkSheet(self: WorkBook, sheetname: [*c]const u8) XlsxError!WorkSheet {
    return WorkSheet{
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

// pub extern fn workbook_add_chartsheet(workbook: [*c]lxw_workbook, sheetname: [*c]const u8) [*c]lxw_chartsheet;

// pub extern fn workbook_set_properties(workbook: [*c]lxw_workbook, properties: [*c]lxw_doc_properties) lxw_error;
// pub extern fn workbook_set_custom_property_string(workbook: [*c]lxw_workbook, name: [*c]const u8, value: [*c]const u8) lxw_error;
// pub extern fn workbook_set_custom_property_number(workbook: [*c]lxw_workbook, name: [*c]const u8, value: f64) lxw_error;
// pub extern fn workbook_set_custom_property_integer(workbook: [*c]lxw_workbook, name: [*c]const u8, value: i32) lxw_error;
// pub extern fn workbook_set_custom_property_boolean(workbook: [*c]lxw_workbook, name: [*c]const u8, value: u8) lxw_error;
// pub extern fn workbook_set_custom_property_datetime(workbook: [*c]lxw_workbook, name: [*c]const u8, datetime: [*c]lxw_datetime) lxw_error;
// pub extern fn workbook_define_name(workbook: [*c]lxw_workbook, name: [*c]const u8, formula: [*c]const u8) lxw_error;
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
const c = @import("xlsxwriter_c");
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;
const WorkSheet = @import("WorkSheet.zig");
const Format = @import("Format.zig");
const Chart = @import("Chart.zig");
