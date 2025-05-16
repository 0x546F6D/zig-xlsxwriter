const WorkBook = @This();

workbook_c: *c.lxw_workbook,

pub fn init(filename: [*c]const u8) XlsxError!WorkBook {
    return WorkBook{
        .workbook_c = c.workbook_new(filename) orelse return XlsxError.NewWorkBook,
    };
}

pub fn deinit(self: WorkBook) XlsxError!void {
    try check(c.workbook_close(self.workbook_c));
}

pub fn addWorkSheet(self: WorkBook, sheetname: [*c]const u8) XlsxError!WorkSheet {
    return WorkSheet{
        .worksheet_c = c.workbook_add_worksheet(self.workbook_c, sheetname) orelse return XlsxError.AddWorkSheet,
    };
}

pub fn addFormat(self: WorkBook) XlsxError!Format {
    return Format{
        .format_c = c.workbook_add_format(self.workbook_c) orelse return XlsxError.AddFormat,
    };
}

pub fn addChart(self: WorkBook, chart_type: Chart.Type) XlsxError!Chart {
    return Chart{
        .chart_c = c.workbook_add_chart(self.workbook_c, @intFromEnum(chart_type)) orelse return XlsxError.AddChart,
    };
}

const c = @import("xlsxwriter_c");
pub const XlsxError = @import("errors.zig").XlsxError;
pub const check = @import("errors.zig").checkResult;
pub const WorkSheet = @import("WorkSheet.zig");
pub const Format = @import("Format.zig");
pub const Chart = @import("Chart.zig");
