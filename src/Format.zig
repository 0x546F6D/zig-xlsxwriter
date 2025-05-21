const Format = @This();

format_c: ?*c.lxw_format,

pub const none: Format = .{
    .format_c = null,
};

pub const DefinedColor = enum(c_int) {
    black = c.LXW_COLOR_BLACK,
    blue = c.LXW_COLOR_BLUE,
    brown = c.LXW_COLOR_BROWN,
    cyan = c.LXW_COLOR_CYAN,
    gray = c.LXW_COLOR_GRAY,
    green = c.LXW_COLOR_GREEN,
    lime = c.LXW_COLOR_LIME,
    magenta = c.LXW_COLOR_MAGENTA,
    navy = c.LXW_COLOR_NAVY,
    orange = c.LXW_COLOR_ORANGE,
    // MAGENTA and PINK have the some color code
    // pink = c.LXW_COLOR_PINK,
    pink = 0xFFC0CB,
    purple = c.LXW_COLOR_PURPLE,
    red = c.LXW_COLOR_RED,
    silver = c.LXW_COLOR_SILVER,
    white = c.LXW_COLOR_WHITE,
    yellow = c.LXW_COLOR_YELLOW,
};

pub inline fn setBold(self: Format) void {
    c.format_set_bold(self.format_c);
}

pub inline fn setItalic(self: Format) void {
    c.format_set_italic(self.format_c);
}

pub inline fn setNumFormat(self: Format, num_format: [*c]const u8) void {
    c.format_set_num_format(self.format_c, num_format);
}

pub const Alignments = enum(c_int) {
    none = c.LXW_ALIGN_NONE,
    left = c.LXW_ALIGN_LEFT,
    center = c.LXW_ALIGN_CENTER,
    right = c.LXW_ALIGN_RIGHT,
    fill = c.LXW_ALIGN_FILL,
    justify = c.LXW_ALIGN_JUSTIFY,
    center_across = c.LXW_ALIGN_CENTER_ACROSS,
    distributed = c.LXW_ALIGN_DISTRIBUTED,
    vertical_top = c.LXW_ALIGN_VERTICAL_TOP,
    vertical_bottom = c.LXW_ALIGN_VERTICAL_BOTTOM,
    vertical_center = c.LXW_ALIGN_VERTICAL_CENTER,
    vertical_justify = c.LXW_ALIGN_VERTICAL_JUSTIFY,
    vertical_distributed = c.LXW_ALIGN_VERTICAL_DISTRIBUTED,
};

// pub extern fn format_set_align(format: [*c]lxw_format, alignment: u8) void;
pub inline fn setAlign(self: Format, alignment: Alignments) void {
    c.format_set_align(self.format_c, @intFromEnum(alignment));
}

// pub extern fn format_set_font_name(format: [*c]lxw_format, font_name: [*c]const u8) void;
// pub extern fn format_set_font_size(format: [*c]lxw_format, size: f64) void;
// pub extern fn format_set_font_color(format: [*c]lxw_format, color: lxw_color_t) void;
// pub extern fn format_set_bold(format: [*c]lxw_format) void;
// pub extern fn format_set_italic(format: [*c]lxw_format) void;
// pub extern fn format_set_underline(format: [*c]lxw_format, style: u8) void;
// pub extern fn format_set_font_strikeout(format: [*c]lxw_format) void;
// pub extern fn format_set_font_script(format: [*c]lxw_format, style: u8) void;
// pub extern fn format_set_font_family(format: [*c]lxw_format, value: u8) void;
// pub extern fn format_set_font_charset(format: [*c]lxw_format, value: u8) void;
// pub extern fn format_set_num_format(format: [*c]lxw_format, num_format: [*c]const u8) void;
// pub extern fn format_set_num_format_index(format: [*c]lxw_format, index: u8) void;
// pub extern fn format_set_unlocked(format: [*c]lxw_format) void;
// pub extern fn format_set_hidden(format: [*c]lxw_format) void;
// pub extern fn format_set_text_wrap(format: [*c]lxw_format) void;
// pub extern fn format_set_rotation(format: [*c]lxw_format, angle: i16) void;
// pub extern fn format_set_indent(format: [*c]lxw_format, level: u8) void;
// pub extern fn format_set_shrink(format: [*c]lxw_format) void;
// pub extern fn format_set_pattern(format: [*c]lxw_format, index: u8) void;
// pub extern fn format_set_bg_color(format: [*c]lxw_format, color: lxw_color_t) void;
// pub extern fn format_set_fg_color(format: [*c]lxw_format, color: lxw_color_t) void;
// pub extern fn format_set_border(format: [*c]lxw_format, style: u8) void;
// pub extern fn format_set_bottom(format: [*c]lxw_format, style: u8) void;
// pub extern fn format_set_top(format: [*c]lxw_format, style: u8) void;
// pub extern fn format_set_left(format: [*c]lxw_format, style: u8) void;
// pub extern fn format_set_right(format: [*c]lxw_format, style: u8) void;
// pub extern fn format_set_border_color(format: [*c]lxw_format, color: lxw_color_t) void;
// pub extern fn format_set_bottom_color(format: [*c]lxw_format, color: lxw_color_t) void;
// pub extern fn format_set_top_color(format: [*c]lxw_format, color: lxw_color_t) void;
// pub extern fn format_set_left_color(format: [*c]lxw_format, color: lxw_color_t) void;
// pub extern fn format_set_right_color(format: [*c]lxw_format, color: lxw_color_t) void;
// pub extern fn format_set_diag_type(format: [*c]lxw_format, @"type": u8) void;
// pub extern fn format_set_diag_border(format: [*c]lxw_format, style: u8) void;
// pub extern fn format_set_diag_color(format: [*c]lxw_format, color: lxw_color_t) void;
// pub extern fn format_set_quote_prefix(format: [*c]lxw_format) void;
// pub extern fn format_set_font_outline(format: [*c]lxw_format) void;
// pub extern fn format_set_font_shadow(format: [*c]lxw_format) void;
// pub extern fn format_set_font_scheme(format: [*c]lxw_format, font_scheme: [*c]const u8) void;
// pub extern fn format_set_font_condense(format: [*c]lxw_format) void;
// pub extern fn format_set_font_extend(format: [*c]lxw_format) void;
// pub extern fn format_set_reading_order(format: [*c]lxw_format, value: u8) void;
// pub extern fn format_set_theme(format: [*c]lxw_format, value: u8) void;
// pub extern fn format_set_hyperlink(format: [*c]lxw_format) void;
// pub extern fn format_set_color_indexed(format: [*c]lxw_format, value: u8) void;
// pub extern fn format_set_font_only(format: [*c]lxw_format) void;

const c = @import("xlsxwriter_c");
