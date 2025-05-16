const WorkBook = @This();

workbook_c: *c.lxw_workbook,

pub inline fn init(filename: [*c]const u8) XlsxError!WorkBook {
    return WorkBook{
        .workbook_c = c.workbook_new(filename) orelse return XlsxError.NewWorkBook,
    };
}

pub inline fn deinit(self: WorkBook) XlsxError!void {
    try check(c.workbook_close(self.workbook_c));
}

pub inline fn addWorkSheet(self: WorkBook, sheetname: [*c]const u8) XlsxError!WorkSheet {
    return WorkSheet{
        .worksheet_c = c.workbook_add_worksheet(self.workbook_c, sheetname) orelse return XlsxError.AddWorkSheet,
    };
}

pub inline fn addFormat(self: WorkBook) XlsxError!Format {
    return Format{
        .format_c = c.workbook_add_format(self.workbook_c) orelse return XlsxError.AddFormat,
    };
}

pub inline fn addChart(self: WorkBook, chart_type: Chart.Type) XlsxError!Chart {
    return Chart{
        .chart_c = c.workbook_add_chart(self.workbook_c, @intFromEnum(chart_type)) orelse return XlsxError.AddChart,
    };
}

const c = @import("xlsxwriter_c");
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;
const WorkSheet = @import("WorkSheet.zig");
const Format = @import("Format.zig");
const Chart = @import("Chart.zig");
