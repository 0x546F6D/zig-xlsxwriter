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

// pub extern fn chartsheet_select(chartsheet: [*c]lxw_chartsheet) void;
// pub extern fn chartsheet_hide(chartsheet: [*c]lxw_chartsheet) void;
// pub extern fn chartsheet_set_first_sheet(chartsheet: [*c]lxw_chartsheet) void;
// pub extern fn chartsheet_set_tab_color(chartsheet: [*c]lxw_chartsheet, color: lxw_color_t) void;
// pub extern fn chartsheet_protect(chartsheet: [*c]lxw_chartsheet, password: [*c]const u8, options: [*c]lxw_protection) void;
// pub extern fn chartsheet_set_zoom(chartsheet: [*c]lxw_chartsheet, scale: u16) void;
// pub extern fn chartsheet_set_landscape(chartsheet: [*c]lxw_chartsheet) void;
// pub extern fn chartsheet_set_portrait(chartsheet: [*c]lxw_chartsheet) void;
// pub extern fn chartsheet_set_paper(chartsheet: [*c]lxw_chartsheet, paper_type: u8) void;
// pub extern fn chartsheet_set_margins(chartsheet: [*c]lxw_chartsheet, left: f64, right: f64, top: f64, bottom: f64) void;
// pub extern fn chartsheet_set_header(chartsheet: [*c]lxw_chartsheet, string: [*c]const u8) lxw_error;
// pub extern fn chartsheet_set_footer(chartsheet: [*c]lxw_chartsheet, string: [*c]const u8) lxw_error;
// pub extern fn chartsheet_set_header_opt(chartsheet: [*c]lxw_chartsheet, string: [*c]const u8, options: [*c]lxw_header_footer_options) lxw_error;
// pub extern fn chartsheet_set_footer_opt(chartsheet: [*c]lxw_chartsheet, string: [*c]const u8, options: [*c]lxw_header_footer_options) lxw_error;
// pub extern fn lxw_chartsheet_new(init_data: [*c]lxw_worksheet_init_data) [*c]lxw_chartsheet;
// pub extern fn lxw_chartsheet_free(chartsheet: [*c]lxw_chartsheet) void;
// pub extern fn lxw_chartsheet_assemble_xml_file(chartsheet: [*c]lxw_chartsheet) void;

const c = @import("lxw");
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;
const Chart = @import("Chart.zig");
