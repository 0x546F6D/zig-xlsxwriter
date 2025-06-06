const Chart = @This();

alloc: ?std.mem.Allocator,
chart_c: ?*c.lxw_chart,
x_axis: ChartAxis,
y_axis: ChartAxis,

// lxw_chart_options
pub const Options = struct {
    x_offset: i32 = 0,
    y_offset: i32 = 0,
    x_scale: f64 = 0,
    y_scale: f64 = 0,
    object_position: ObjectPosition = .default,
    description: ?[*:0]const u8 = null,
    decorative: bool = false,

    pub inline fn toC(self: Options) c.lxw_chart_options {
        return c.lxw_chart_options{
            .x_offset = self.x_offset,
            .y_offset = self.y_offset,
            .x_scale = self.x_scale,
            .y_scale = self.y_scale,
            .object_position = @intFromEnum(self.object_position),
            .description = self.description,
            .decorative = @intFromBool(self.decorative),
        };
    }
};

// lxw_chart_pattern
pub const Pattern = struct {
    fg_color: DefinedColors = .default,
    bg_color: DefinedColors = .default,
    type: PatternType = .none,

    pub const default = Pattern{
        .fg_color = .default,
        .bg_color = .default,
        .type = .none,
    };

    pub inline fn toC(self: Pattern) c.lxw_chart_pattern {
        return c.lxw_chart_pattern{
            .fg_color = @intFromEnum(self.fg_color),
            .bg_color = @intFromEnum(self.bg_color),
            .type = @intFromEnum(self.type),
        };
    }
};

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

pub const LineDashType = enum(u8) {
    solid = c.LXW_CHART_LINE_DASH_SOLID,
    round_dot = c.LXW_CHART_LINE_DASH_ROUND_DOT,
    square_dot = c.LXW_CHART_LINE_DASH_SQUARE_DOT,
    dash = c.LXW_CHART_LINE_DASH_DASH,
    dash_dot = c.LXW_CHART_LINE_DASH_DASH_DOT,
    long_dash = c.LXW_CHART_LINE_DASH_LONG_DASH,
    long_dash_dot = c.LXW_CHART_LINE_DASH_LONG_DASH_DOT,
    long_dash_dot_dot = c.LXW_CHART_LINE_DASH_LONG_DASH_DOT_DOT,
    dot = c.LXW_CHART_LINE_DASH_DOT,
    system_dash_dot = c.LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT,
    system_dash_dot_dot = c.LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT_DOT,
};

pub const PatternType = enum(u8) {
    none = c.LXW_CHART_PATTERN_NONE,
    percent_5 = c.LXW_CHART_PATTERN_PERCENT_5,
    percent_10 = c.LXW_CHART_PATTERN_PERCENT_10,
    percent_20 = c.LXW_CHART_PATTERN_PERCENT_20,
    percent_25 = c.LXW_CHART_PATTERN_PERCENT_25,
    percent_30 = c.LXW_CHART_PATTERN_PERCENT_30,
    percent_40 = c.LXW_CHART_PATTERN_PERCENT_40,
    percent_50 = c.LXW_CHART_PATTERN_PERCENT_50,
    percent_60 = c.LXW_CHART_PATTERN_PERCENT_60,
    percent_70 = c.LXW_CHART_PATTERN_PERCENT_70,
    percent_75 = c.LXW_CHART_PATTERN_PERCENT_75,
    percent_80 = c.LXW_CHART_PATTERN_PERCENT_80,
    percent_90 = c.LXW_CHART_PATTERN_PERCENT_90,
    downward_diagonal = c.LXW_CHART_PATTERN_LIGHT_DOWNWARD_DIAGONAL,
    upward_diagonal = c.LXW_CHART_PATTERN_LIGHT_UPWARD_DIAGONAL,
    dark_downward_diagonal = c.LXW_CHART_PATTERN_DARK_DOWNWARD_DIAGONAL,
    dark_upward_diagonal = c.LXW_CHART_PATTERN_DARK_UPWARD_DIAGONAL,
    wide_downward_diagonal = c.LXW_CHART_PATTERN_WIDE_DOWNWARD_DIAGONAL,
    wide_upward_diagonal = c.LXW_CHART_PATTERN_WIDE_UPWARD_DIAGONAL,
    light_vertical = c.LXW_CHART_PATTERN_LIGHT_VERTICAL,
    light_horizontal = c.LXW_CHART_PATTERN_LIGHT_HORIZONTAL,
    narrow_vertical = c.LXW_CHART_PATTERN_NARROW_VERTICAL,
    narrow_horizontal = c.LXW_CHART_PATTERN_NARROW_HORIZONTAL,
    dark_vertical = c.LXW_CHART_PATTERN_DARK_VERTICAL,
    dark_horizontal = c.LXW_CHART_PATTERN_DARK_HORIZONTAL,
    dashed_downward_diagonal = c.LXW_CHART_PATTERN_DASHED_DOWNWARD_DIAGONAL,
    dashed_upward_diagonal = c.LXW_CHART_PATTERN_DASHED_UPWARD_DIAGONAL,
    dashed_horizontal = c.LXW_CHART_PATTERN_DASHED_HORIZONTAL,
    dashed_vertical = c.LXW_CHART_PATTERN_DASHED_VERTICAL,
    small_confetti = c.LXW_CHART_PATTERN_SMALL_CONFETTI,
    large_confetti = c.LXW_CHART_PATTERN_LARGE_CONFETTI,
    zigzag = c.LXW_CHART_PATTERN_ZIGZAG,
    wave = c.LXW_CHART_PATTERN_WAVE,
    diagonal_brick = c.LXW_CHART_PATTERN_DIAGONAL_BRICK,
    horizontal_brick = c.LXW_CHART_PATTERN_HORIZONTAL_BRICK,
    weave = c.LXW_CHART_PATTERN_WEAVE,
    plaid = c.LXW_CHART_PATTERN_PLAID,
    divot = c.LXW_CHART_PATTERN_DIVOT,
    dotted_grid = c.LXW_CHART_PATTERN_DOTTED_GRID,
    dotted_diamond = c.LXW_CHART_PATTERN_DOTTED_DIAMOND,
    shingle = c.LXW_CHART_PATTERN_SHINGLE,
    sphere = c.LXW_CHART_PATTERN_SPHERE,
    trellis = c.LXW_CHART_PATTERN_TRELLIS,
    small_grid = c.LXW_CHART_PATTERN_SMALL_GRID,
    large_grid = c.LXW_CHART_PATTERN_LARGE_GRID,
    small_check = c.LXW_CHART_PATTERN_SMALL_CHECK,
    large_check = c.LXW_CHART_PATTERN_LARGE_CHECK,
    outlined_diamond = c.LXW_CHART_PATTERN_OUTLINED_DIAMOND,
    solid_diamond = c.LXW_CHART_PATTERN_SOLID_DIAMOND,
};

pub const SubType = enum(u8) {
    none = c.LXW_CHART_SUBTYPE_NONE,
    stacked = c.LXW_CHART_SUBTYPE_STACKED,
    stacked_percent = c.LXW_CHART_SUBTYPE_STACKED_PERCENT,
};

pub const Grouping = enum(u8) {
    clustered = c.LXW_GROUPING_CLUSTERED,
    standard = c.LXW_GROUPING_STANDARD,
    percent_stacked = c.LXW_GROUPING_PERCENTSTACKED,
    stacked = c.LXW_GROUPING_STACKED,
};

// pub extern fn chart_title_set_name(chart: [*c]lxw_chart, name: [*c]const u8) void;
pub inline fn titleSetName(self: Chart, name: [:0]const u8) void {
    c.chart_title_set_name(self.chart_c, name);
}

// pub extern fn chart_title_set_name_range(chart: [*c]lxw_chart, sheetname: [*c]const u8, row: lxw_row_t, col: lxw_col_t) void;
pub inline fn titleSetNameRange(
    self: Chart,
    name: [:0]const u8,
    cell: Cell,
) void {
    c.chart_title_set_name_range(
        self.chart_c,
        name,
        cell.row,
        cell.col,
    );
}

pub const Font = struct {
    name: ?[*:0]const u8 = null,
    size: f64 = 0,
    bold: bool = true,
    italic: bool = false,
    underline: bool = false,
    rotation: i32 = 0,
    color: DefinedColors = .default,
    pitch_family: u8 = 0,
    charset: u8 = 0,
    baseline: i8 = 0,

    pub const default = Font{
        .name = null,
        .size = 0,
        .bold = false,
        .italic = false,
        .underline = false,
        .rotation = 0,
        .color = .default,
        .pitch_family = 0,
        .charset = 0,
        .baseline = 0,
    };

    pub inline fn toC(self: Font) c.lxw_chart_font {
        return c.struct_lxw_chart_font{
            .name = self.name,
            .size = self.size,
            // bold by default, need to explicitely disable it
            .bold = if (self.bold) @intFromBool(self.bold) else @intFromEnum(Bool.explicit_false),
            .italic = @intFromBool(self.italic),
            .underline = @intFromBool(self.underline),
            .rotation = self.rotation,
            .color = @intFromEnum(self.color),
            .pitch_family = self.pitch_family,
            .charset = self.charset,
            .baseline = self.baseline,
        };
    }
};

// pub extern fn chart_title_set_name_font(chart: [*c]lxw_chart, font: [*c]lxw_chart_font) void;
pub inline fn titleSetNameFont(self: Chart, font: Font) void {
    c.chart_title_set_name_font(
        self.chart_c,
        @constCast(&font.toC()),
    );
}

// pub extern fn chart_title_off(chart: [*c]lxw_chart) void;
pub inline fn titleOff(self: Chart) void {
    c.chart_title_off(self.chart_c);
}

// lxw_chart_layout
pub const Layout = struct {
    x: f64 = 0,
    y: f64 = 0,
    width: f64 = 0,
    height: f64 = 0,
    has_inner: bool = false,

    pub const default = Layout{
        .x = 0,
        .y = 0,
        .width = 0,
        .height = 0,
        .has_inner = false,
    };

    inline fn toC(self: Layout) c.lxw_chart_layout {
        return c.lxw_chart_layout{
            .x = self.x,
            .y = self.y,
            .width = self.width,
            .height = self.height,
            .has_inner = @intFromBool(self.has_inner),
        };
    }
};

// pub extern fn chart_title_set_layout(chart: [*c]lxw_chart, layout: [*c]lxw_chart_layout) void;
pub inline fn titleSetLayout(self: Chart, layout: Layout) void {
    c.chart_title_set_name_font(
        self.chart_c,
        @constCast(&layout.toC()),
    );
}

// pub extern fn chart_title_set_overlay(chart: [*c]lxw_chart, overlay: u8) void;
pub inline fn titleSetOverlay(self: Chart, overlay: bool) void {
    c.chart_title_set_overlay(
        self.chart_c,
        @intFromBool(overlay),
    );
}

// pub extern fn chart_set_style(chart: [*c]lxw_chart, style_id: u8) void;
pub inline fn setStyle(self: Chart, style_id: u8) void {
    c.chart_set_style(
        self.chart_c,
        style_id,
    );
}

// pub extern fn chart_set_table(chart: [*c]lxw_chart) void;
pub inline fn setTable(self: Chart) void {
    c.chart_set_table(self.chart_c);
}

// pub extern fn chart_set_table_grid(chart: [*c]lxw_chart, horizontal: u8, vertical: u8, outline: u8, legend_keys: u8) void;
pub inline fn setTableGrid(
    self: Chart,
    horizontal: bool,
    vertical: bool,
    outline: bool,
    legend_keys: bool,
) void {
    c.chart_set_table_grid(
        self.chart_c,
        @intFromBool(horizontal),
        @intFromBool(vertical),
        @intFromBool(outline),
        @intFromBool(legend_keys),
    );
}

// pub extern fn chart_set_table_font(chart: [*c]lxw_chart, font: [*c]lxw_chart_font) void;
pub inline fn setTableFont(self: Chart, font: Font) void {
    c.chart_set_table_font(
        self.chart_c,
        @constCast(&font.toC()),
    );
}

// pub extern fn chart_set_up_down_bars(chart: [*c]lxw_chart) void;
pub inline fn setUpDownBars(self: Chart) void {
    c.chart_set_up_down_bars(self.chart_c);
}

// lxw_chart_line
pub const Line = struct {
    color: DefinedColors = .default,
    none: bool = false,
    width: f32 = 0,
    dash_type: LineDashType = .solid,
    transparency: u8 = 0,

    pub const default = Line{
        .color = .default,
        .none = false,
        .width = 0,
        .dash_type = .solid,
        .transparency = 0,
    };

    pub inline fn toC(self: Line) c.lxw_chart_line {
        return c.lxw_chart_line{
            .color = @intFromEnum(self.color),
            .none = @intFromBool(self.none),
            .width = self.width,
            .dash_type = @intFromEnum(self.dash_type),
            .transparency = self.transparency,
        };
    }
};

// lxw_chart_fill
pub const Fill = struct {
    color: DefinedColors = .default,
    none: bool = false,
    transparency: u8 = 0,

    pub const default = Fill{
        .color = .default,
        .none = false,
        .transparency = 0,
    };

    pub inline fn toC(self: Fill) c.lxw_chart_fill {
        return c.lxw_chart_fill{
            .color = @intFromEnum(self.color),
            .none = @intFromBool(self.none),
            .transparency = self.transparency,
        };
    }
};

// pub extern fn chart_set_up_down_bars_format(chart: [*c]lxw_chart, up_bar_line: [*c]lxw_chart_line, up_bar_fill: [*c]lxw_chart_fill, down_bar_line: [*c]lxw_chart_line, down_bar_fill: [*c]lxw_chart_fill) void;
pub inline fn setUpDownBarsFormat(
    self: Chart,
    up_bar_line: ?Line,
    up_bar_fill: ?Fill,
    down_bar_line: ?Line,
    down_bar_fill: ?Fill,
) void {
    c.chart_set_up_down_bars_format(
        self.chart_c,
        if (up_bar_line) |line| @constCast(&line.toC()) else null,
        if (up_bar_fill) |fill| @constCast(&fill.toC()) else null,
        if (down_bar_line) |line| @constCast(&line.toC()) else null,
        if (down_bar_fill) |fill| @constCast(&fill.toC()) else null,
        // @constCast(&up_bar_line.toC()),
        // @constCast(&up_bar_fill.toC()),
        // @constCast(&down_bar_line.toC()),
        // @constCast(&down_bar_fill.toC()),
    );
}

// pub extern fn chart_set_drop_lines(chart: [*c]lxw_chart, line: [*c]lxw_chart_line) void;
pub inline fn setDropLines(
    self: Chart,
    line: ?Line,
) void {
    c.chart_set_drop_lines(
        self.chart_c,
        if (line) |line_o| @constCast(&line_o.toC()) else null,
    );
}

// pub extern fn chart_set_high_low_lines(chart: [*c]lxw_chart, line: [*c]lxw_chart_line) void;
pub inline fn setHighLowLines(
    self: Chart,
    line: ?Line,
) void {
    c.chart_set_high_low_lines(
        self.chart_c,
        if (line) |line_o| @constCast(&line_o.toC()) else null,
    );
}

// pub extern fn chart_set_series_overlap(chart: [*c]lxw_chart, overlap: i8) void;
pub inline fn setSeriesOverlap(self: Chart, overlap: i8) void {
    c.chart_set_series_overlap(self.chart_c, overlap);
}

// pub extern fn chart_set_series_gap(chart: [*c]lxw_chart, gap: u16) void;
pub inline fn setSeriesGap(self: Chart, gap: u16) void {
    c.chart_set_series_gap(self.chart_c, gap);
}

pub const ChartBlank = enum(u8) {
    gap = c.LXW_CHART_BLANKS_AS_GAP,
    zero = c.LXW_CHART_BLANKS_AS_ZERO,
    connected = c.LXW_CHART_BLANKS_AS_CONNECTED,
};
// pub extern fn chart_show_blanks_as(chart: [*c]lxw_chart, option: u8) void;
pub inline fn showBlanksAs(
    self: Chart,
    option: ChartBlank,
) void {
    c.chart_show_blanks_as(
        self.chart_c,
        @intFromEnum(option),
    );
}

// pub extern fn chart_show_hidden_data(chart: [*c]lxw_chart) void;
pub inline fn showHiddenData(self: Chart) void {
    c.chart_show_hidden_data(self.chart_c);
}

// pub extern fn chart_set_rotation(chart: [*c]lxw_chart, rotation: u16) void;
pub inline fn setRotation(self: Chart, rotation: u16) void {
    c.chart_set_rotation(self.chart_c, rotation);
}

// pub extern fn chart_set_hole_size(chart: [*c]lxw_chart, size: u8) void;
pub inline fn setHoleSize(self: Chart, size: u16) void {
    c.chart_set_hole_size(self.chart_c, size);
}

pub const LegendPosition = enum(u8) {
    none = c.LXW_CHART_LEGEND_NONE,
    right = c.LXW_CHART_LEGEND_RIGHT,
    left = c.LXW_CHART_LEGEND_LEFT,
    top = c.LXW_CHART_LEGEND_TOP,
    bottom = c.LXW_CHART_LEGEND_BOTTOM,
    top_right = c.LXW_CHART_LEGEND_TOP_RIGHT,
    overlay_right = c.LXW_CHART_LEGEND_OVERLAY_RIGHT,
    overlay_left = c.LXW_CHART_LEGEND_OVERLAY_LEFT,
    overlay_top_right = c.LXW_CHART_LEGEND_OVERLAY_TOP_RIGHT,
};

// pub extern fn chart_legend_set_position(chart: [*c]lxw_chart, position: u8) void;
pub inline fn legendSetPosition(
    self: Chart,
    position: LegendPosition,
) void {
    c.chart_legend_set_position(
        self.chart_c,
        @intFromEnum(position),
    );
}

// pub extern fn chart_legend_set_layout(chart: [*c]lxw_chart, layout: [*c]lxw_chart_layout) void;
pub inline fn legendSetLayout(self: Chart, layout: Layout) void {
    c.chart_legend_set_layout(
        self.chart_c,
        @constCast(&layout.toC()),
    );
}

// pub extern fn chart_legend_set_font(chart: [*c]lxw_chart, font: [*c]lxw_chart_font) void;
pub inline fn legendsetFont(self: Chart, font: Font) void {
    c.chart_legend_set_font(
        self.chart_c,
        @constCast(&font.toC()),
    );
}

pub const DeleteSeries = [:-1]i16;
// pub extern fn chart_legend_delete_series(chart: [*c]lxw_chart, delete_series: [*c]i16) lxw_error;
pub inline fn legendDeleteSeries(
    self: Chart,
    delete_series: DeleteSeries,
) void {
    c.chart_legend_delete_series(
        self.chart_c,
        @ptrCast(@constCast(&delete_series)),
    );
}

// pub extern fn chart_chartarea_set_line(chart: [*c]lxw_chart, line: [*c]lxw_chart_line) void;
pub inline fn chartAreaSetLine(
    self: Chart,
    line: Line,
) void {
    c.chart_chartarea_set_line(
        self.chart_c,
        @constCast(&line.toC()),
    );
}

// pub extern fn chart_chartarea_set_fill(chart: [*c]lxw_chart, fill: [*c]lxw_chart_fill) void;
pub inline fn chartAreaSetFill(
    self: Chart,
    fill: Fill,
) void {
    c.chart_chartarea_set_fill(
        self.chart_c,
        @constCast(&fill.toC()),
    );
}

// pub extern fn chart_chartarea_set_pattern(chart: [*c]lxw_chart, pattern: [*c]lxw_chart_pattern) void;
pub inline fn chartAreaSetPattern(
    self: Chart,
    pattern: PatternType,
) void {
    c.chart_chartarea_set_pattern(
        self.chart_c,
        @intFromEnum(pattern),
    );
}

// pub extern fn chart_plotarea_set_line(chart: [*c]lxw_chart, line: [*c]lxw_chart_line) void;
pub inline fn plotAreaSetLine(
    self: Chart,
    line: Line,
) void {
    c.chart_plotarea_set_line(
        self.chart_c,
        @constCast(&line.toC()),
    );
}

// pub extern fn chart_plotarea_set_fill(chart: [*c]lxw_chart, fill: [*c]lxw_chart_fill) void;
pub inline fn plotAreaSetFill(
    self: Chart,
    fill: Fill,
) void {
    c.chart_plotarea_set_fill(
        self.chart_c,
        @constCast(&fill.toC()),
    );
}

// pub extern fn chart_plotarea_set_pattern(chart: [*c]lxw_chart, pattern: [*c]lxw_chart_pattern) void;
pub inline fn plotAreaSetPattern(
    self: Chart,
    pattern: PatternType,
) void {
    c.chart_plotarea_set_pattern(
        self.chart_c,
        @intFromEnum(pattern),
    );
}

// pub extern fn chart_plotarea_set_layout(chart: [*c]lxw_chart, layout: [*c]lxw_chart_layout) void;
pub inline fn plotAreaSetLayout(
    self: Chart,
    layout: Layout,
) void {
    c.chart_plotarea_set_layout(
        self.chart_c,
        @constCast(&layout.toC()),
    );
}

// pub extern fn chart_add_series(chart: [*c]lxw_chart, categories: [*c]const u8, values: [*c]const u8) [*c]lxw_chart_series;
pub inline fn addSeries(self: Chart, categories: ?[*:0]const u8, values: ?[*:0]const u8) XlsxError!ChartSeries {
    const series = c.chart_add_series(
        self.chart_c,
        categories,
        values,
    ) orelse
        return XlsxError.ChartAddSeries;
    return ChartSeries{
        .alloc = self.alloc,
        .chartseries_c = series,
        .x_error_bars = .{ .error_bars_c = series.*.x_error_bars },
        .y_error_bars = .{ .error_bars_c = series.*.y_error_bars },
    };
}

// pub extern fn chart_axis_get(chart: [*c]lxw_chart, axis_type: lxw_chart_axis_type) [*c]lxw_chart_axis;
pub inline fn getAxis(
    self: Chart,
    axis_type: ChartAxis,
) ChartAxis {
    return ChartAxis{
        .axis_c = c.chart_axis_get(
            self.chart_c,
            axis_type,
        ),
    };
}

const std = @import("std");
const c = @import("lxw");
const xlsxwriter = @import("xlsxwriter.zig");
const CStringArray = xlsxwriter.CStringArray;
const Bool = xlsxwriter.Boolean;
const XlsxError = @import("errors.zig").XlsxError;
const Cell = @import("utility.zig").Cell;
const DefinedColors = @import("format.zig").DefinedColors;
pub const ChartSeries = @import("ChartSeries.zig");
pub const ChartPoint = ChartSeries.Point;
pub const ChartPointNoAlloc = ChartSeries.PointNoAlloc;
pub const ChartPointNoAllocArray = ChartSeries.PointNoAllocArray;
pub const ChartDataLabel = ChartSeries.DataLabel;
pub const ChartAxis = @import("ChartAxis.zig");
pub const ChartAxisType = ChartAxis.Type;
const ObjectPosition = @import("WorkSheet.zig").ObjectPosition;
