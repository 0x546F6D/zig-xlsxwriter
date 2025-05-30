const ChartSheet = @This();

// alloc: ?std.mem.Allocator,
chartsheet_c: ?*c.lxw_chartsheet,

// pub extern fn chartsheet_activate(chartsheet: [*c]lxw_chartsheet) void;
pub inline fn activate(self: ChartSheet) void {
    c.chartsheet_activate(self.chartsheet_c);
}

// pub extern fn chartsheet_set_chart(chartsheet: [*c]lxw_chartsheet, chart: [*c]lxw_chart) lxw_error;
pub inline fn setChart(
    self: ChartSheet,
    chart: Chart,
) XlsxError!void {
    try check(c.chartsheet_set_chart(
        self.chartsheet_c,
        chart.chart_c,
    ));
}

// pub extern fn chartsheet_set_chart_opt(chartsheet: [*c]lxw_chartsheet, chart: [*c]lxw_chart, user_options: [*c]lxw_chart_options) lxw_error;
pub inline fn setChartOpt(
    self: ChartSheet,
    chart: Chart,
    user_options: ChartOptions,
) XlsxError!void {
    try check(c.chartsheet_set_chart_opt(
        self.chartsheet_c,
        chart.chart_c,
        @constCast(&user_options.toC()),
    ));
}

// pub extern fn chartsheet_select(chartsheet: [*c]lxw_chartsheet) void;
pub inline fn select(self: ChartSheet) XlsxError!void {
    try check(c.chartsheet_select(self.chartsheet_c));
}

// pub extern fn chartsheet_hide(chartsheet: [*c]lxw_chartsheet) void;
pub inline fn hide(self: ChartSheet) XlsxError!void {
    try check(c.chartsheet_hide(self.chartsheet_c));
}

// pub extern fn chartsheet_set_first_sheet(chartsheet: [*c]lxw_chartsheet) void;
pub inline fn setFirstSheet(self: ChartSheet) XlsxError!void {
    try check(c.chartsheet_set_first_sheet(self.chartsheet_c));
}

// pub extern fn chartsheet_set_tab_color(chartsheet: [*c]lxw_chartsheet, color: lxw_color_t) void;
pub inline fn setTabColor(
    self: ChartSheet,
    color: DefinedColors,
) XlsxError!void {
    try check(c.chartsheet_set_tab_color(
        self.chartsheet_c,
        @intFromEnum(color),
    ));
}

// pub extern fn chartsheet_protect(chartsheet: [*c]lxw_chartsheet, password: [*c]const u8, options: [*c]lxw_protection) void;
pub inline fn protect(
    self: ChartSheet,
    password: ?CString,
    options: Protection,
) void {
    c.chartsheet_protect(
        self.chartsheet_c,
        password,
        @constCast(&options.toC()),
    );
}

// pub extern fn chartsheet_set_zoom(chartsheet: [*c]lxw_chartsheet, scale: u16) void;
pub inline fn setZoom(self: ChartSheet, scale: u16) void {
    c.chartsheet_set_zoom(self.chartsheet_c, scale);
}

// pub extern fn chartsheet_set_landscape(chartsheet: [*c]lxw_chartsheet) void;
pub inline fn setLandscape(self: ChartSheet) void {
    c.chartsheet_set_landscape(self.chartsheet_c);
}

// pub extern fn chartsheet_set_portrait(chartsheet: [*c]lxw_chartsheet) void;
pub inline fn setPortrait(self: ChartSheet) void {
    c.chartsheet_set_landscape(self.chartsheet_c);
}

// pub extern fn chartsheet_set_paper(chartsheet: [*c]lxw_chartsheet, paper_type: u8) void;
// paper_type in chartsheet.h is actually paper_size in chartsheet.c
pub inline fn setPaper(self: ChartSheet, paper_size: u8) void {
    c.chartsheet_set_paper(self.chartsheet_c, paper_size);
}

// pub extern fn chartsheet_set_margins(chartsheet: [*c]lxw_chartsheet, left: f64, right: f64, top: f64, bottom: f64) void;
pub inline fn setMargins(
    self: ChartSheet,
    left: f64,
    right: f64,
    top: f64,
    bottom: f64,
) void {
    c.chartsheet_set_margins(
        self.chartsheet_c,
        left,
        right,
        top,
        bottom,
    );
}

// pub extern fn chartsheet_set_header(chartsheet: [*c]lxw_chartsheet, string: [*c]const u8) lxw_error;
pub inline fn setHeader(self: ChartSheet, string: ?CString) XlsxError!void {
    try check(c.chartsheet_set_header(self.chartsheet_c, string));
}

// pub extern fn chartsheet_set_footer(chartsheet: [*c]lxw_chartsheet, string: [*c]const u8) lxw_error;
pub inline fn setFooter(self: ChartSheet, string: ?CString) XlsxError!void {
    try check(c.chartsheet_set_footer(self.chartsheet_c, string));
}

// pub extern fn chartsheet_set_header_opt(chartsheet: [*c]lxw_chartsheet, string: [*c]const u8, options: [*c]lxw_header_footer_options) lxw_error;
pub inline fn setHeaderOpt(
    self: ChartSheet,
    string: ?CString,
    options: HeaderFooterOptions,
) XlsxError!void {
    try check(c.chartsheet_set_header_opt(
        self.chartsheet_c,
        string,
        @constCast(&options),
    ));
}

// pub extern fn chartsheet_set_footer_opt(chartsheet: [*c]lxw_chartsheet, string: [*c]const u8, options: [*c]lxw_header_footer_options) lxw_error;
pub inline fn setFooterOpt(
    self: ChartSheet,
    string: ?CString,
    options: HeaderFooterOptions,
) XlsxError!void {
    try check(c.chartsheet_set_footer_opt(
        self.chartsheet_c,
        string,
        @constCast(&options),
    ));
}

const c = @import("lxw");
const xlsxwriter = @import("xlsxwriter.zig");
const CString = xlsxwriter.CString;
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;
const Chart = @import("Chart.zig");
const ChartOptions = Chart.Options;
const DefinedColors = @import("Format.zig").DefinedColors;
const Protection = @import("WorkSheet.zig").Protection;
const HeaderFooterOptions = @import("WorkSheet.zig").HeaderFooterOptions;
