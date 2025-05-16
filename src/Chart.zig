const Chart = @This();

chart_c: *c.lxw_chart,

pub const Series = *c.lxw_chart_series;
pub const Font = c.lxw_chart_font;

pub const Type = enum(u8) {
    area = c.LXW_CHART_AREA,
    area_stacked = c.LXW_CHART_AREA_STACKED,
    area_stacked_percent = c.LXW_CHART_AREA_STACKED_PERCENT,
    bar = c.LXW_CHART_BAR,
    bar_stacked = c.LXW_CHART_BAR_STACKED,
    bar_stacked_percent = c.LXW_CHART_BAR_STACKED_PERCENT,
    column = c.LXW_CHART_COLUMN,
    column_stacked = c.LXW_CHART_COLUMN_STACKED,
    column_stacked_percent = c.LXW_CHART_COLUMN_STACKED_PERCENT,
    doughnut = c.LXW_CHART_DOUGHNUT,
    line = c.LXW_CHART_LINE,
    line_stacked = c.LXW_CHART_LINE_STACKED,
    line_stacked_percent = c.LXW_CHART_LINE_STACKED_PERCENT,
    pie = c.LXW_CHART_PIE,
    scatter = c.LXW_CHART_SCATTER,
    scatter_straight = c.LXW_CHART_SCATTER_STRAIGHT,
    scatter_straight_with_markers = c.LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS,
    scatter_smooth = c.LXW_CHART_SCATTER_SMOOTH,
    scatter_smooth_with_markers = c.LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS,
    radar = c.LXW_CHART_RADAR,
    radar_with_markers = c.LXW_CHART_RADAR_WITH_MARKERS,
    radar_filled = c.LXW_CHART_RADAR_FILLED,
};

pub inline fn addSeries(self: Chart, categories: [*c]const u8, values: [*c]const u8) XlsxError!Series {
    return c.chart_add_series(self.chart_c, categories, values) orelse XlsxError.ChartAddSeries;
}

pub inline fn titleSetName(self: Chart, name: [*c]const u8, font: ?*Font) void {
    c.chart_title_set_name(self.chart_c, name);
    if (font) |f|
        c.chart_title_set_name_font(self.chart_c, f);
}

const c = @import("xlsxwriter_c");
const XlsxError = @import("errors.zig").XlsxError;
