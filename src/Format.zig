const Format = @This();

format_c: ?*c.lxw_format,

pub const none: Format = .{
    .format_c = null,
};

// pub extern fn format_set_bold(format: [*c]lxw_format) void;
pub inline fn setBold(self: Format) void {
    c.format_set_bold(self.format_c);
}

// pub extern fn format_set_italic(format: [*c]lxw_format) void;
pub inline fn setItalic(self: Format) void {
    c.format_set_italic(self.format_c);
}

// pub extern fn format_set_num_format(format: [*c]lxw_format, num_format: [*c]const u8) void;
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

// pub extern fn format_set_font_strikeout(format: [*c]lxw_format) void;
pub inline fn setFontStrikeout(self: Format) void {
    c.format_set_font_strikeout(self.format_c);
}

pub const Underlines = enum(c_int) {
    none = c.LXW_UNDERLINE_NONE,
    single = c.LXW_UNDERLINE_SINGLE,
    double = c.LXW_UNDERLINE_DOUBLE,
    single_accounting = c.LXW_UNDERLINE_SINGLE_ACCOUNTING,
    double_accounting = c.LXW_UNDERLINE_DOUBLE_ACCOUNTING,
};

// pub extern fn format_set_underline(format: [*c]lxw_format, style: u8) void;
pub inline fn setUnderline(self: Format, style: Underlines) void {
    c.format_set_underline(self.format_c, @intFromEnum(style));
}

pub const DefinedColors = enum(c_int) {
    // Some colors are changed to reflect Excel default RGB code
    black = c.LXW_COLOR_BLACK,
    white = c.LXW_COLOR_WHITE,
    // blue = c.LXW_COLOR_BLUE,
    blue = 0x0070C0,
    brown = c.LXW_COLOR_BROWN,
    cyan = c.LXW_COLOR_CYAN,
    gray = c.LXW_COLOR_GRAY,
    // green = c.LXW_COLOR_GREEN,
    green = 0x00B050,
    // lime = c.LXW_COLOR_LIME,
    lime = 0xE4FFB5,
    magenta = c.LXW_COLOR_MAGENTA,
    navy = c.LXW_COLOR_NAVY,
    // orange = c.LXW_COLOR_ORANGE,
    orange = 0xFFC000,
    // MAGENTA and PINK have the some color code
    // pink = c.LXW_COLOR_PINK,
    pink = 0xFFC0CB,
    // purple = c.LXW_COLOR_PURPLE,
    purple = 0x7030A0,
    red = c.LXW_COLOR_RED,
    silver = c.LXW_COLOR_SILVER,
    yellow = c.LXW_COLOR_YELLOW,
    // some more colors taken from excel default palette
    // Standart Colors
    dark_red = 0xC00000,
    light_green = 0x92D050,
    light_blue = 0x00B0F0,
    dark_blue = 0x002060,
    // Additional Colors
    rose = 0xFF7474,
    light_rose = 0xFFC6C6,
    light_yellow = 0xFFFB8A,
    bright_green = 0x29B95C,
    sky_blue = 0xD0FFFF,
    turquoise = 0x4EC9FF,
    light_turquoise = 0x76E3FF,
    periwinkle = 0x7C8AD2,
    lavender = 0xC189F7,
    light_lavender = 0xF6C3FF,
    indigo = 0x6373BA,
    gold = 0xFFCE3C,
    light_brown = 0xA66500,
    light_gray = 0xD4D4D4,
    dark_gray = 0x212121,
    darker_gray = 0x4B4B4B,
    // Theme Colors:
    // color number stands for color accent/text number
    // "d" stands for dark %, "l" for light %
    white_d5 = 0xF2F2F2,
    white_d15 = 0xD9D9D9,
    white_d25 = 0xBFBFBF,
    white_d35 = 0xA6A6A6,
    // they are both equivalent to gray
    // white_d50 = 0x808080,
    // black_l50 = 0x808080,
    black_l35 = 0x595959,
    black_l25 = 0x404040,
    black_l15 = 0x262626,
    black_l5 = 0x0D0D0D,
    tan = 0xEEECE1,
    tan_d10 = 0xDDD9C4,
    tan_d25 = 0xC4BD97,
    tan_d50 = 0x948A54,
    tan_d75 = 0x494529,
    tan_d90 = 0x1D1B10,
    dark_blue2 = 0x1F497D,
    dark_blue2_l80 = 0xC5D9F1,
    dark_blue2_l60 = 0x8DB4E2,
    dark_blue2_l40 = 0x538DD5,
    dark_blue2_d25 = 0x16365C,
    dark_blue2_d50 = 0x0F243E,
    blue1 = 0x4F81BD,
    blue1_l80 = 0xDCE6F1,
    blue1_l60 = 0xB8CCE4,
    blue1_l40 = 0x95B3D7,
    blue1_d25 = 0x366092,
    blue1_d50 = 0x244062,
    red2 = 0xC0504D,
    red2_l80 = 0xF2DCDB,
    red2_l60 = 0xE6B8B7,
    red2_l40 = 0xDA9694,
    red2_d25 = 0x963634,
    red2_d50 = 0x632523,
    olive_green3 = 0x9BBB59,
    olive_green3_l80 = 0xEBF1DE,
    olive_green3_l60 = 0xD8E4BC,
    olive_green3_l40 = 0xC4D79B,
    olive_green3_d25 = 0x76933C,
    olive_green3_d50 = 0x4F6228,
    purple4 = 0x8064A2,
    purple4_l80 = 0xE4DFEC,
    purple4_l60 = 0xCCC0DA,
    purple4_l40 = 0xB1A0C7,
    purple4_d25 = 0x60497A,
    purple4_d50 = 0x403151,
    aqua5 = 0x4BACC6,
    aqua5_l80 = 0xDAEEF3,
    aqua5_l60 = 0xB7DEE8,
    aqua5_l40 = 0x92CDDC,
    aqua5_d25 = 0x31869B,
    aqua5_d50 = 0x215967,
    orange6 = 0xF79646,
    orange6_l80 = 0xFDE9D9,
    orange6_l60 = 0xFCD5B4,
    orange6_l40 = 0xFABF8F,
    orange6_d25 = 0xE26B0A,
    orange6_d50 = 0x974706,
    _,
};

// pub extern fn format_set_font_color(format: [*c]lxw_format, color: lxw_color_t) void;
pub inline fn setFontColor(self: Format, color: DefinedColors) void {
    c.format_set_font_color(self.format_c, @intFromEnum(color));
}

// pub extern fn format_set_bg_color(format: [*c]lxw_format, color: lxw_color_t) void;
pub inline fn setBgColor(self: Format, color: DefinedColors) void {
    c.format_set_bg_color(self.format_c, @intFromEnum(color));
}

// pub extern fn format_set_fg_color(format: [*c]lxw_format, color: lxw_color_t) void;
pub inline fn setFgColor(self: Format, color: DefinedColors) void {
    c.format_set_fg_color(self.format_c, @intFromEnum(color));
}

pub const Scripts = enum(c_int) {
    superscript = c.LXW_FONT_SUPERSCRIPT,
    subscript = c.LXW_FONT_SUBSCRIPT,
};

// pub extern fn format_set_font_script(format: [*c]lxw_format, style: u8) void;
pub inline fn setFontScript(self: Format, style: Scripts) void {
    c.format_set_font_script(self.format_c, @intFromEnum(style));
}

pub const Borders = enum(c_int) {
    none = c.LXW_BORDER_NONE,
    thin = c.LXW_BORDER_THIN,
    medium = c.LXW_BORDER_MEDIUM,
    dashed = c.LXW_BORDER_DASHED,
    dotted = c.LXW_BORDER_DOTTED,
    thick = c.LXW_BORDER_THICK,
    double = c.LXW_BORDER_DOUBLE,
    hair = c.LXW_BORDER_HAIR,
    medium_dashed = c.LXW_BORDER_MEDIUM_DASHED,
    dash_dot = c.LXW_BORDER_DASH_DOT,
    medium_dash_dot = c.LXW_BORDER_MEDIUM_DASH_DOT,
    dash_dot_dot = c.LXW_BORDER_DASH_DOT_DOT,
    medium_dash_dot_dot = c.LXW_BORDER_MEDIUM_DASH_DOT_DOT,
    slant_dash_dot = c.LXW_BORDER_SLANT_DASH_DOT,
};

// pub extern fn format_set_border(format: [*c]lxw_format, style: u8) void;
pub inline fn setBorder(self: Format, style: Borders) void {
    c.format_set_border(self.format_c, @intFromEnum(style));
}

// pub extern fn format_set_border_color(format: [*c]lxw_format, color: lxw_color_t) void;
pub inline fn setBorderColor(self: Format, color: DefinedColors) void {
    c.format_set_border_color(self.format_c, @intFromEnum(color));
}

// pub extern fn format_set_text_wrap(format: [*c]lxw_format) void;
pub inline fn setTextWrap(self: Format) void {
    c.format_set_text_wrap(self.format_c);
}

// pub extern fn format_set_indent(format: [*c]lxw_format, level: u8) void;
pub inline fn setIndent(self: Format, level: u8) void {
    c.format_set_indent(self.format_c, level);
}

// pub extern fn format_set_font_name(format: [*c]lxw_format, font_name: [*c]const u8) void;
// pub extern fn format_set_font_size(format: [*c]lxw_format, size: f64) void;
// pub extern fn format_set_font_family(format: [*c]lxw_format, value: u8) void;
// pub extern fn format_set_font_charset(format: [*c]lxw_format, value: u8) void;
// pub extern fn format_set_num_format_index(format: [*c]lxw_format, index: u8) void;
// pub extern fn format_set_unlocked(format: [*c]lxw_format) void;
// pub extern fn format_set_hidden(format: [*c]lxw_format) void;
// pub extern fn format_set_rotation(format: [*c]lxw_format, angle: i16) void;
// pub extern fn format_set_shrink(format: [*c]lxw_format) void;
// pub extern fn format_set_pattern(format: [*c]lxw_format, index: u8) void;
// pub extern fn format_set_bottom(format: [*c]lxw_format, style: u8) void;
// pub extern fn format_set_top(format: [*c]lxw_format, style: u8) void;
// pub extern fn format_set_left(format: [*c]lxw_format, style: u8) void;
// pub extern fn format_set_right(format: [*c]lxw_format, style: u8) void;
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
