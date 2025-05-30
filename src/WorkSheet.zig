const WorkSheet = @This();

alloc: ?std.mem.Allocator,
worksheet_c: ?*c.lxw_worksheet,

// set functions
pub const RowColOptions = struct {
    hidden: bool = false,
    level: u8 = 0,
    collapsed: bool = false,

    inline fn toC(self: RowColOptions) c.lxw_row_col_options {
        return c.lxw_row_col_options{
            .hidden = @intFromBool(self.hidden),
            .level = self.level,
            .collapsed = @intFromBool(self.collapsed),
        };
    }
};

// pub extern fn worksheet_set_row(worksheet: [*c]lxw_worksheet, row: lxw_row_t, height: f64, format: [*c]lxw_format) lxw_error;
pub inline fn setRow(
    self: WorkSheet,
    row: u32,
    height: f64,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_set_row(
        self.worksheet_c,
        row,
        height,
        format.format_c,
    ));
}

// pub extern fn worksheet_set_row_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, height: f64, format: [*c]lxw_format, options: [*c]lxw_row_col_options) lxw_error;
pub inline fn setRowOpt(
    self: WorkSheet,
    row: u32,
    height: f64,
    format: Format,
    options: RowColOptions,
) XlsxError!void {
    try check(c.worksheet_set_row_opt(
        self.worksheet_c,
        row,
        height,
        format.format_c,
        @constCast(&options.toC()),
    ));
}

// pub extern fn worksheet_set_row_pixels(worksheet: [*c]lxw_worksheet, row: lxw_row_t, pixels: u32, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_set_row_pixels_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, pixels: u32, format: [*c]lxw_format, options: [*c]lxw_row_col_options) lxw_error;

// pub extern fn worksheet_set_column(worksheet: [*c]lxw_worksheet, first_col: lxw_col_t, last_col: lxw_col_t, width: f64, format: [*c]lxw_format) lxw_error;
pub inline fn setColumn(
    self: WorkSheet,
    first_col: u16,
    last_col: u16,
    width: f64,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_set_column(
        self.worksheet_c,
        first_col,
        last_col,
        width,
        format.format_c,
    ));
}

// pub extern fn worksheet_set_column_opt(worksheet: [*c]lxw_worksheet, first_col: lxw_col_t, last_col: lxw_col_t, width: f64, format: [*c]lxw_format, options: [*c]lxw_row_col_options) lxw_error;
pub inline fn setColumnOpt(
    self: WorkSheet,
    first_col: u16,
    last_col: u16,
    width: f64,
    format: Format,
    options: RowColOptions,
) XlsxError!void {
    try check(c.worksheet_set_column_opt(
        self.worksheet_c,
        first_col,
        last_col,
        width,
        format.format_c,
        @constCast(&options.toC()),
    ));
}

// pub extern fn worksheet_set_column_pixels(worksheet: [*c]lxw_worksheet, first_col: lxw_col_t, last_col: lxw_col_t, pixels: u32, format: [*c]lxw_format) lxw_error;
pub inline fn setColumnPixels(
    self: WorkSheet,
    first_col: u16,
    last_col: u16,
    pixels: u32,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_set_column_pixels(
        self.worksheet_c,
        first_col,
        last_col,
        pixels,
        format.format_c,
    ));
}

// pub extern fn worksheet_set_column_pixels_opt(worksheet: [*c]lxw_worksheet, first_col: lxw_col_t, last_col: lxw_col_t, pixels: u32, format: [*c]lxw_format, options: [*c]lxw_row_col_options) lxw_error;
pub inline fn setColumnPixelsOpt(
    self: WorkSheet,
    first_col: u16,
    last_col: u16,
    pixels: u32,
    format: Format,
    options: RowColOptions,
) XlsxError!void {
    try check(c.worksheet_set_column_pixels_opt(
        self.worksheet_c,
        first_col,
        last_col,
        pixels,
        format.format_c,
        @constCast(&options.toC()),
    ));
}

// pub extern fn worksheet_set_h_pagebreaks(worksheet: [*c]lxw_worksheet, breaks: [*c]lxw_row_t) lxw_error;
pub const RowBreaks = [*:0]const u32;
pub inline fn setHPageBreaks(
    self: WorkSheet,
    breaks: RowBreaks,
) XlsxError!void {
    try check(c.worksheet_set_h_pagebreaks(
        self.worksheet_c,
        @constCast(breaks),
    ));
}

// pub extern fn worksheet_set_v_pagebreaks(worksheet: [*c]lxw_worksheet, breaks: [*c]lxw_col_t) lxw_error;
pub const ColBreaks = [*:0]const u16;
pub inline fn setVPageBreaks(
    self: WorkSheet,
    breaks: ColBreaks,
) XlsxError!void {
    try check(c.worksheet_set_v_pagebreaks(
        self.worksheet_c,
        @constCast(breaks),
    ));
}

pub const Margins = extern struct {
    left: f64 = 0,
    right: f64 = 0,
    top: f64 = 0,
    bottom: f64 = 0,
};
// pub extern fn worksheet_set_margins(worksheet: [*c]lxw_worksheet, left: f64, right: f64, top: f64, bottom: f64) void;
pub inline fn setMargins(
    self: WorkSheet,
    margins: Margins,
) void {
    c.worksheet_set_margins(
        self.worksheet_c,
        margins.left,
        margins.right,
        margins.top,
        margins.bottom,
    );
}

// pub extern fn worksheet_set_default_row(worksheet: [*c]lxw_worksheet, height: f64, hide_unused_rows: u8) void;
pub inline fn setDefaultRow(
    self: WorkSheet,
    height: f64,
    hide_unused_rows: bool,
) void {
    c.worksheet_set_default_row(
        self.worksheet_c,
        height,
        @intFromBool(hide_unused_rows),
    );
}

// pub extern fn worksheet_set_selection(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t) lxw_error;
pub inline fn setSelection(
    self: WorkSheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
) XlsxError!void {
    try check(c.worksheet_set_selection(
        self.worksheet_c,
        first_row,
        first_col,
        last_row,
        last_col,
    ));
}

// pub extern fn worksheet_set_top_left_cell(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t) void;
pub inline fn setTopLeftCell(
    self: WorkSheet,
    row: u32,
    col: u16,
) void {
    c.worksheet_set_top_left_cell(self.worksheet_c, row, col);
}

// pub extern fn worksheet_set_landscape(worksheet: [*c]lxw_worksheet) void;
pub inline fn setLandscape(self: WorkSheet) void {
    c.worksheet_set_landscape(self.worksheet_c);
}

// pub extern fn worksheet_set_portrait(worksheet: [*c]lxw_worksheet) void;
pub inline fn setPortrait(self: WorkSheet) void {
    c.worksheet_set_portrait(self.worksheet_c);
}

// pub extern fn worksheet_set_page_view(worksheet: [*c]lxw_worksheet) void;
pub inline fn setPageView(self: WorkSheet) void {
    c.worksheet_set_page_view(self.worksheet_c);
}

// pub extern fn worksheet_set_paper(worksheet: [*c]lxw_worksheet, paper_type: u8) void;
// paper_type in worksheet.h is actually paper_size in worksheet.c
pub inline fn setPaper(self: WorkSheet, paper_size: u8) void {
    c.worksheet_set_paper(self.worksheet_c, paper_size);
}

// pub extern fn worksheet_set_zoom(worksheet: [*c]lxw_worksheet, scale: u16) void;
pub inline fn setZoom(self: WorkSheet, scale: u16) void {
    c.worksheet_set_zoom(self.worksheet_c, scale);
}

// pub extern fn worksheet_set_start_page(worksheet: [*c]lxw_worksheet, start_page: u16) void;
pub inline fn setStartPage(self: WorkSheet, start_page: u16) void {
    c.worksheet_set_start_page(self.worksheet_c, start_page);
}

// pub extern fn worksheet_set_print_scale(worksheet: [*c]lxw_worksheet, scale: u16) void;
pub inline fn setPrintScale(self: WorkSheet, scale: u16) void {
    c.worksheet_set_print_scale(self.worksheet_c, scale);
}

// pub extern fn worksheet_set_vba_name(worksheet: [*c]lxw_worksheet, name: [*c]const u8) lxw_error;
pub inline fn setVbaName(
    self: WorkSheet,
    name: ?CString,
) XlsxError!void {
    try check(c.worksheet_set_vba_name(self.worksheet_c, name));
}

// pub extern fn worksheet_set_comments_author(worksheet: [*c]lxw_worksheet, author: [*c]const u8) void;
pub inline fn setCommentsAuthor(self: WorkSheet, author: ?CString) void {
    c.worksheet_set_comments_author(self.worksheet_c, author);
}

pub const ObjectProperties = c.lxw_object_properties;
// pub extern fn worksheet_set_error_cell(worksheet: [*c]lxw_worksheet, object_props: [*c]lxw_object_properties, ref_id: u32) void;
pub inline fn setErrorCell(self: WorkSheet, object_props: ObjectProperties) void {
    c.worksheet_set_comments_author(
        self.worksheet_c,
        @constCast(&object_props),
    );
}

// write functions
// pub extern fn worksheet_write_number(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, number: f64, format: [*c]lxw_format) lxw_error;
pub inline fn writeNumber(
    self: WorkSheet,
    row: u32,
    col: u16,
    number: f64,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_number(
        self.worksheet_c,
        row,
        col,
        number,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_string(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, string: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn writeString(
    self: WorkSheet,
    row: u32,
    col: u16,
    string: ?CString,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_string(
        self.worksheet_c,
        row,
        col,
        string,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_formula(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn writeFormula(
    self: WorkSheet,
    row: u32,
    col: u16,
    formula: ?CString,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_formula(
        self.worksheet_c,
        row,
        col,
        formula,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_array_formula(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn writeArrayFormula(
    self: WorkSheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    formula: ?CString,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_array_formula(
        self.worksheet_c,
        first_row,
        first_col,
        last_row,
        last_col,
        formula,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_dynamic_formula(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn writeDynamicFormula(
    self: WorkSheet,
    row: u32,
    col: u16,
    formula: ?CString,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_dynamic_formula(
        self.worksheet_c,
        row,
        col,
        formula,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_dynamic_array_formula(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn writeDynamicArrayFormula(
    self: WorkSheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    formula: ?CString,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_dynamic_array_formula(
        self.worksheet_c,
        first_row,
        first_col,
        last_row,
        last_col,
        formula,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_array_formula_num(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: f64) lxw_error;
// pub extern fn worksheet_write_dynamic_array_formula_num(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: f64) lxw_error;
// pub extern fn worksheet_write_dynamic_formula_num(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: f64) lxw_error;

// pub extern fn worksheet_write_datetime(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, datetime: [*c]lxw_datetime, format: [*c]lxw_format) lxw_error;
pub inline fn writeDateTime(
    self: WorkSheet,
    row: u32,
    col: u16,
    datetime: c.lxw_datetime,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_datetime(
        self.worksheet_c,
        row,
        col,
        @constCast(&datetime),
        format.format_c,
    ));
}

// pub extern fn worksheet_write_unixtime(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, unixtime: i64, format: [*c]lxw_format) lxw_error;
pub inline fn writeUnixTime(
    self: WorkSheet,
    row: u32,
    col: u16,
    unixtime: i64,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_unixtime(
        self.worksheet_c,
        row,
        col,
        unixtime,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_url(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, url: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn writeUrl(
    self: WorkSheet,
    row: u32,
    col: u16,
    url: ?CString,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_url(
        self.worksheet_c,
        row,
        col,
        url,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_url_opt(worksheet: [*c]lxw_worksheet, row_num: lxw_row_t, col_num: lxw_col_t, url: [*c]const u8, format: [*c]lxw_format, string: [*c]const u8, tooltip: [*c]const u8) lxw_error;
pub inline fn writeUrlOpt(
    self: WorkSheet,
    row: u32,
    col: u16,
    url: ?CString,
    format: Format,
    string: ?CString,
    tooltip: ?CString,
) XlsxError!void {
    try check(c.worksheet_write_url_opt(
        self.worksheet_c,
        row,
        col,
        url,
        format.format_c,
        string,
        tooltip,
    ));
}

pub const RichStringTuple = struct {
    format: Format = .default,
    string: ?CString,

    pub const default = RichStringTuple{
        .format = .default,
        .string = null,
    };

    inline fn toC(self: RichStringTuple) c.lxw_rich_string_tuple {
        return c.lxw_rich_string_tuple{
            .format = self.format.format_c,
            .string = self.string,
        };
    }
};

// pub extern fn worksheet_write_rich_string(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, rich_string: [*c][*c]lxw_rich_string_tuple, format: [*c]lxw_format) lxw_error;
pub inline fn writeRichString(
    self: WorkSheet,
    row: u32,
    col: u16,
    rich_string: []const RichStringTuple,
    format: Format,
) !void {
    const allocator: std.mem.Allocator = if (self.alloc) |allocator|
        allocator
    else
        return XlsxError.WriteRichString;

    var rich_string_array = try allocator.alloc(c.lxw_rich_string_tuple, rich_string.len);
    defer allocator.free(rich_string_array);
    var rich_string_c = try allocator.allocSentinel(?*c.lxw_rich_string_tuple, rich_string.len, null);
    defer allocator.free(rich_string_c);

    for (rich_string, 0..) |tuple, i| {
        rich_string_array[i] = tuple.toC();
        rich_string_c[i] = &rich_string_array[i];
    }

    try check(c.worksheet_write_rich_string(
        self.worksheet_c,
        row,
        col,
        @ptrCast(rich_string_c),
        format.format_c,
    ));
}

pub const RichStringTupleNoAlloc = extern struct {
    format_c: ?*c.lxw_format = null,
    string: [*c]const u8,
};

pub const RichStringNoAllocArray: type = [:null]const ?*const RichStringTupleNoAlloc;
// pub extern fn worksheet_write_rich_string(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, rich_string: [*c][*c]lxw_rich_string_t uple, format: [*c]lxw_format) lxw_error;
pub inline fn writeRichStringNoAlloc(
    self: WorkSheet,
    row: u32,
    col: u16,
    rich_string: RichStringNoAllocArray,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_rich_string(
        self.worksheet_c,
        row,
        col,
        @ptrCast(@constCast(rich_string)),
        format.format_c,
    ));
}

// pub extern fn worksheet_write_comment(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, string: [*c]const u8) lxw_error;
pub inline fn writeComment(
    self: WorkSheet,
    row: u32,
    col: u16,
    string: ?CString,
) XlsxError!void {
    try check(c.worksheet_write_comment(
        self.worksheet_c,
        row,
        col,
        string,
    ));
}

pub const CommentDisplayTypes = enum(u8) {
    default = c.LXW_COMMENT_DISPLAY_DEFAULT,
    hidden = c.LXW_COMMENT_DISPLAY_HIDDEN,
    visible = c.LXW_COMMENT_DISPLAY_VISIBLE,
};

pub const CommentOptions = extern struct {
    visible: CommentDisplayTypes = .default,
    author: ?CString = null,
    width: u16 = 0,
    height: u16 = 0,
    x_scale: f64 = 1,
    y_scale: f64 = 1,
    color: Format.DefinedColors = .default,
    font_name: ?CString = null,
    font_size: f64 = 8,
    font_family: u8 = 0,
    start_row: u32 = 0,
    start_col: u16 = 0,
    x_offset: i32 = 0,
    y_offset: i32 = 0,

    pub const default = CommentOptions{
        .visible = .default,
        .author = null,
        .width = 0,
        .height = 0,
        .x_scale = 1,
        .y_scale = 1,
        .color = .default,
        .font_name = null,
        .font_size = 8,
        .font_family = 0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
    };

    inline fn toC(self: CommentOptions) c.lxw_comment_options {
        return c.lxw_comment_options{
            .visible = @intFromEnum(self.visible),
            .author = self.author,
            .width = self.width,
            .height = self.height,
            .x_scale = self.x_scale,
            .y_scale = self.y_scale,
            .color = @intFromEnum(self.color),
            .font_name = self.font_name,
            .font_size = self.font_size,
            .font_family = self.font_family,
            .start_row = self.start_row,
            .start_col = self.start_col,
            .x_offset = self.x_offset,
            .y_offset = self.y_offset,
        };
    }
};

// pub extern fn worksheet_write_comment_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, string: [*c]const u8, options: [*c]lxw_comment_options) lxw_error;
pub inline fn writeCommentOpt(
    self: WorkSheet,
    row: u32,
    col: u16,
    string: ?CString,
    options: CommentOptions,
) XlsxError!void {
    try check(c.worksheet_write_comment_opt(
        self.worksheet_c,
        row,
        col,
        string,
        @constCast(&options.toC()),
    ));
}

// pub extern fn worksheet_show_comments(worksheet: [*c]lxw_worksheet) void;
pub inline fn showComments(self: WorkSheet) void {
    c.worksheet_show_comments(self.worksheet_c);
}

// pub extern fn worksheet_write_boolean(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, value: c_int, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_write_blank(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_write_formula_num(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: f64) lxw_error;
// pub extern fn worksheet_write_formula_str(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: [*c]const u8) lxw_error;

// image functions
pub const ImageOptions = struct {
    x_offset: i32 = 0,
    y_offset: i32 = 0,
    x_scale: f64 = 0,
    y_scale: f64 = 0,
    object_position: u8 = 0,
    description: ?CString = null,
    decorative: bool = false,
    url: ?CString = null,
    tip: ?CString = null,
    cell_format: Format = .default,

    pub const default = ImageOptions{
        .x_offset = 0,
        .y_offset = 0,
        .x_scale = 0,
        .y_scale = 0,
        .object_position = 0,
        .description = null,
        .decorative = false,
        .url = null,
        .tip = null,
        .cell_format = .default,
    };

    inline fn toC(self: ImageOptions) c.lxw_image_options {
        return c.lxw_image_options{
            .x_offset = self.x_offset,
            .y_offset = self.y_offset,
            .x_scale = self.x_scale,
            .y_scale = self.y_scale,
            .object_position = self.object_position,
            .description = self.description,
            .decorative = @intFromBool(self.decorative),
            .url = self.url,
            .tip = self.tip,
            .cell_format = self.cell_format.format_c,
        };
    }
};

// pub extern fn worksheet_insert_image(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, filename: [*c]const u8) lxw_error;
pub inline fn insertImage(
    self: WorkSheet,
    row: u32,
    col: u16,
    filename: ?CString,
) XlsxError!void {
    try check(c.worksheet_insert_image(
        self.worksheet_c,
        row,
        col,
        filename,
    ));
}

// pub extern fn worksheet_insert_image_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, filename: [*c]const u8, options: [*c]lxw_image_options) lxw_error;
pub inline fn insertImageOpt(
    self: WorkSheet,
    row: u32,
    col: u16,
    filename: ?CString,
    options: ImageOptions,
) XlsxError!void {
    try check(c.worksheet_insert_image_opt(
        self.worksheet_c,
        row,
        col,
        filename,
        @constCast(&options.toC()),
    ));
}

// pub extern fn worksheet_insert_image_buffer(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, image_buffer: [*c]const u8, image_size: usize) lxw_error;
pub inline fn insertImageBuffer(
    self: WorkSheet,
    row: u32,
    col: u16,
    image_buffer: ?CString,
    image_size: usize,
) XlsxError!void {
    try check(c.worksheet_insert_image_buffer(
        self.worksheet_c,
        row,
        col,
        image_buffer,
        image_size,
    ));
}

// pub extern fn worksheet_insert_image_buffer_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, image_buffer: [*c]const u8, image_size: usize, options: [*c]lxw_image_options) lxw_error;
pub inline fn insertImageBufferOpt(
    self: WorkSheet,
    row: u32,
    col: u16,
    image_buffer: ?CString,
    image_size: usize,
    options: ImageOptions,
) XlsxError!void {
    try check(c.worksheet_insert_image_buffer_opt(
        self.worksheet_c,
        row,
        col,
        image_buffer,
        image_size,
        @constCast(&options.toC()),
    ));
}

// pub extern fn worksheet_embed_image(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, filename: [*c]const u8) lxw_error;
pub inline fn embedImage(
    self: WorkSheet,
    row: u32,
    col: u16,
    filename: ?CString,
) XlsxError!void {
    try check(c.worksheet_embed_image(
        self.worksheet_c,
        row,
        col,
        filename,
    ));
}

// pub extern fn worksheet_embed_image_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, filename: [*c]const u8, options: [*c]lxw_image_options) lxw_error;
pub inline fn embedImageOpt(
    self: WorkSheet,
    row: u32,
    col: u16,
    filename: ?CString,
    options: ImageOptions,
) XlsxError!void {
    try check(c.worksheet_embed_image_opt(
        self.worksheet_c,
        row,
        col,
        filename,
        @constCast(&options.toC()),
    ));
}

// pub extern fn worksheet_embed_image_buffer(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, image_buffer: [*c]const u8, image_size: usize) lxw_error;
pub inline fn embedImageBuffer(
    self: WorkSheet,
    row: u32,
    col: u16,
    image_buffer: ?CString,
    image_size: usize,
) XlsxError!void {
    try check(c.worksheet_embed_image_buffer(
        self.worksheet_c,
        row,
        col,
        image_buffer,
        image_size,
    ));
}

// pub extern fn worksheet_embed_image_buffer_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, image_buffer: [*c]const u8, image_size: usize, options: [*c]lxw_image_options) lxw_error;
pub inline fn embedImageBufferOpt(
    self: WorkSheet,
    row: u32,
    col: u16,
    image_buffer: ?CString,
    image_size: usize,
    options: ImageOptions,
) XlsxError!void {
    try check(c.worksheet_embed_image_buffer_opt(
        self.worksheet_c,
        row,
        col,
        image_buffer,
        image_size,
        @constCast(&options.toC()),
    ));
}

// pub extern fn worksheet_set_background(worksheet: [*c]lxw_worksheet, filename: [*c]const u8) lxw_error;
pub inline fn setBackGround(
    self: WorkSheet,
    filename: ?CString,
) XlsxError!void {
    try check(c.worksheet_set_background(
        self.worksheet_c,
        filename,
    ));
}

// pub extern fn worksheet_set_background_buffer(worksheet: [*c]lxw_worksheet, image_buffer: [*c]const u8, image_size: usize) lxw_error;
pub inline fn setBackGroundBuffer(
    self: WorkSheet,
    image_buffer: ?CString,
    image_size: usize,
) XlsxError!void {
    try check(c.worksheet_set_background_buffer(
        self.worksheet_c,
        image_buffer,
        image_size,
    ));
}

// pub extern fn worksheet_set_tab_color(worksheet: [*c]lxw_worksheet, color: lxw_color_t) void;
pub inline fn setTabColor(
    self: WorkSheet,
    color: Format.DefinedColors,
) void {
    c.worksheet_set_tab_color(self.worksheet_c, @intFromEnum(color));
}

// chart functions

pub const ObjectPosition = enum(u8) {
    default = c.LXW_OBJECT_POSITION_DEFAULT,
    move_and_size = c.LXW_OBJECT_MOVE_AND_SIZE,
    move_dont_size = c.LXW_OBJECT_MOVE_DONT_SIZE,
    dont_move_dont_size = c.LXW_OBJECT_DONT_MOVE_DONT_SIZE,
    move_and_size_after = c.LXW_OBJECT_MOVE_AND_SIZE_AFTER,
};

pub const ChartOptions = struct {
    x_offset: i32 = 0,
    y_offset: i32 = 0,
    x_scale: f64 = 0,
    y_scale: f64 = 0,
    object_position: ObjectPosition = .default,
    description: ?CString = null,
    decorative: bool = false,

    inline fn toC(self: ChartOptions) c.lxw_chart_options {
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

// pub extern fn worksheet_insert_chart(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, chart: [*c]lxw_chart) lxw_error;
pub inline fn insertChart(
    self: WorkSheet,
    row: u32,
    col: u16,
    chart: Chart,
) XlsxError!void {
    try check(c.worksheet_insert_chart(
        self.worksheet_c,
        row,
        col,
        chart.chart_c,
    ));
}

// pub extern fn worksheet_insert_chart_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, chart: [*c]lxw_chart, user_options: [*c]lxw_chart_options) lxw_error;
pub inline fn insertChartOpt(
    self: WorkSheet,
    row: u32,
    col: u16,
    chart: Chart,
    user_options: ChartOptions,
) XlsxError!void {
    try check(c.worksheet_insert_chart_opt(
        self.worksheet_c,
        row,
        col,
        chart.chart_c,
        @constCast(&user_options.toC()),
    ));
}

// filter functions
pub const autoFilter = filter.autoFilter;
pub const filterColumn = filter.filterColumn;
pub const filterColumn2 = filter.filterColumn2;
pub const filterList = filter.filterList;

// panes functions
// pub extern fn worksheet_freeze_panes(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t) void;
pub inline fn freezePanes(
    self: WorkSheet,
    top_row: u32,
    left_col: u16,
) void {
    c.worksheet_freeze_panes(
        self.worksheet_c,
        top_row,
        left_col,
    );
}

// pub extern fn worksheet_freeze_panes_opt(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, top_row: lxw_row_t, left_col: lxw_col_t, @"type": u8) void;
pub inline fn freezePanesOpt(
    self: WorkSheet,
    first_row: u32,
    first_col: u16,
    top_row: u32,
    left_col: u16,
    @"type": bool,
) void {
    c.worksheet_freeze_panes_opt(
        self.worksheet_c,
        first_row,
        first_col,
        top_row,
        left_col,
        @intFromBool(@"type"),
    );
}

// pub extern fn worksheet_split_panes(worksheet: [*c]lxw_worksheet, vertical: f64, horizontal: f64) void;
pub inline fn splitPanes(
    self: WorkSheet,
    vertical: f64,
    horizontal: f64,
) void {
    c.worksheet_split_panes(
        self.worksheet_c,
        vertical,
        horizontal,
    );
}

// pub extern fn worksheet_split_panes_opt(worksheet: [*c]lxw_worksheet, vertical: f64, horizontal: f64, top_row: lxw_row_t, left_col: lxw_col_t) void;
pub inline fn splitPanesOpt(
    self: WorkSheet,
    vertical: f64,
    horizontal: f64,
    top_row: u32,
    left_col: u16,
) void {
    c.worksheet_split_panes_opt(
        self.worksheet_c,
        vertical,
        horizontal,
        top_row,
        left_col,
    );
}

// other functions
// pub extern fn worksheet_merge_range(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, string: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn mergeRange(
    self: WorkSheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    string: ?CString,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_merge_range(
        self.worksheet_c,
        first_row,
        first_col,
        last_row,
        last_col,
        string,
        format.format_c,
    ));
}

pub const ValidationBoolean = enum(u8) {
    default = c.LXW_VALIDATION_DEFAULT,
    off = c.LXW_VALIDATION_OFF,
    on = c.LXW_VALIDATION_ON,
};
pub const ValidationTypes = enum(u8) {
    none = c.LXW_VALIDATION_TYPE_NONE,
    integer = c.LXW_VALIDATION_TYPE_INTEGER,
    integer_formula = c.LXW_VALIDATION_TYPE_INTEGER_FORMULA,
    decimal = c.LXW_VALIDATION_TYPE_DECIMAL,
    decimal_formula = c.LXW_VALIDATION_TYPE_DECIMAL_FORMULA,
    list = c.LXW_VALIDATION_TYPE_LIST,
    list_formula = c.LXW_VALIDATION_TYPE_LIST_FORMULA,
    date = c.LXW_VALIDATION_TYPE_DATE,
    date_formula = c.LXW_VALIDATION_TYPE_DATE_FORMULA,
    date_number = c.LXW_VALIDATION_TYPE_DATE_NUMBER,
    time = c.LXW_VALIDATION_TYPE_TIME,
    time_formula = c.LXW_VALIDATION_TYPE_TIME_FORMULA,
    time_number = c.LXW_VALIDATION_TYPE_TIME_NUMBER,
    length = c.LXW_VALIDATION_TYPE_LENGTH,
    length_formula = c.LXW_VALIDATION_TYPE_LENGTH_FORMULA,
    custom_formula = c.LXW_VALIDATION_TYPE_CUSTOM_FORMULA,
    any = c.LXW_VALIDATION_TYPE_ANY,
};
pub const ValidationCriteria = enum(u8) {
    none = c.LXW_VALIDATION_CRITERIA_NONE,
    between = c.LXW_VALIDATION_CRITERIA_BETWEEN,
    not_between = c.LXW_VALIDATION_CRITERIA_NOT_BETWEEN,
    equal_to = c.LXW_VALIDATION_CRITERIA_EQUAL_TO,
    not_equal_to = c.LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO,
    greater_than = c.LXW_VALIDATION_CRITERIA_GREATER_THAN,
    less_than = c.LXW_VALIDATION_CRITERIA_LESS_THAN,
    greater_than_or_equal_to = c.LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO,
    less_than_or_equal_to = c.LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO,
};
pub const ValidationErrorTypes = enum(u8) {
    stop = c.LXW_VALIDATION_ERROR_TYPE_STOP,
    warning = c.LXW_VALIDATION_ERROR_TYPE_WARNING,
    information = c.LXW_VALIDATION_ERROR_TYPE_INFORMATION,
};
pub const DataValidation = struct {
    validate: ValidationTypes = .none,
    criteria: ValidationCriteria = .none,
    ignore_blank: ValidationBoolean = .default,
    show_input: ValidationBoolean = .default,
    show_error: ValidationBoolean = .default,
    error_type: ValidationErrorTypes = .stop,
    dropdown: ValidationBoolean = .default,
    value_number: f64 = 0,
    value_formula: ?CString = null,
    value_list: CStringArray = &.{},
    value_datetime: c.lxw_datetime = .{},
    minimum_number: f64 = 0,
    minimum_formula: ?CString = null,
    minimum_datetime: c.lxw_datetime = .{},
    maximum_number: f64 = 0,
    maximum_formula: ?CString = null,
    maximum_datetime: c.lxw_datetime = .{},
    input_title: ?CString = null,
    input_message: ?CString = null,
    error_title: ?CString = null,
    error_message: ?CString = null,

    pub const default = DataValidation{
        .validate = .none,
        .criteria = .none,
        .ignore_blank = .default,
        .show_input = .default,
        .show_error = .default,
        .error_type = .stop,
        .dropdown = .default,
        .value_number = 0,
        .value_formula = null,
        .value_list = &.{},
        .value_datetime = .{},
        .minimum_number = 0,
        .minimum_formula = null,
        .minimum_datetime = .{},
        .maximum_number = 0,
        .maximum_formula = null,
        .maximum_datetime = .{},
        .input_title = null,
        .input_message = null,
        .error_title = null,
        .error_message = null,
    };

    inline fn toC(self: DataValidation) c.lxw_data_validation {
        return c.lxw_data_validation{
            .validate = @intFromEnum(self.validate),
            .criteria = @intFromEnum(self.criteria),
            .ignore_blank = @intFromEnum(self.ignore_blank),
            .show_input = @intFromEnum(self.show_input),
            .show_error = @intFromEnum(self.show_error),
            .error_type = @intFromEnum(self.error_type),
            .dropdown = @intFromEnum(self.dropdown),
            .value_number = self.value_number,
            .value_formula = self.value_formula,
            .value_list = @ptrCast(@constCast(self.value_list)),
            .value_datetime = self.value_datetime,
            .minimum_number = self.minimum_number,
            .minimum_formula = self.minimum_formula,
            .minimum_datetime = self.minimum_datetime,
            .maximum_number = self.maximum_number,
            .maximum_formula = self.maximum_formula,
            .maximum_datetime = self.maximum_datetime,
            .input_title = self.input_title,
            .input_message = self.input_message,
            .error_title = self.error_title,
            .error_message = self.error_message,
        };
    }
};

// pub extern fn worksheet_data_validation_cell(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, validation: [*c]lxw_data_validation) lxw_error;
pub inline fn dataValidationCell(
    self: WorkSheet,
    row: u32,
    col: u16,
    validation: DataValidation,
) XlsxError!void {
    try check(c.worksheet_data_validation_cell(
        self.worksheet_c,
        row,
        col,
        @constCast(&validation.toC()),
    ));
}

pub const ConditionalFormatTypes = enum(u8) {
    none = c.LXW_CONDITIONAL_TYPE_NONE,
    cell = c.LXW_CONDITIONAL_TYPE_CELL,
    text = c.LXW_CONDITIONAL_TYPE_TEXT,
    time_period = c.LXW_CONDITIONAL_TYPE_TIME_PERIOD,
    average = c.LXW_CONDITIONAL_TYPE_AVERAGE,
    duplicate = c.LXW_CONDITIONAL_TYPE_DUPLICATE,
    unique = c.LXW_CONDITIONAL_TYPE_UNIQUE,
    top = c.LXW_CONDITIONAL_TYPE_TOP,
    bottom = c.LXW_CONDITIONAL_TYPE_BOTTOM,
    blanks = c.LXW_CONDITIONAL_TYPE_BLANKS,
    no_blanks = c.LXW_CONDITIONAL_TYPE_NO_BLANKS,
    errors = c.LXW_CONDITIONAL_TYPE_ERRORS,
    no_errors = c.LXW_CONDITIONAL_TYPE_NO_ERRORS,
    formula = c.LXW_CONDITIONAL_TYPE_FORMULA,
    color_2_scale = c.LXW_CONDITIONAL_2_COLOR_SCALE,
    color_3_scale = c.LXW_CONDITIONAL_3_COLOR_SCALE,
    data_bar = c.LXW_CONDITIONAL_DATA_BAR,
    icon_sets = c.LXW_CONDITIONAL_TYPE_ICON_SETS,
    last = c.LXW_CONDITIONAL_TYPE_LAST,
};

pub const ConditionalCriteria = enum(u8) {
    none = c.LXW_CONDITIONAL_CRITERIA_NONE,
    equal_to = c.LXW_CONDITIONAL_CRITERIA_EQUAL_TO,
    not_equal_to = c.LXW_CONDITIONAL_CRITERIA_NOT_EQUAL_TO,
    greater_than = c.LXW_CONDITIONAL_CRITERIA_GREATER_THAN,
    less_than = c.LXW_CONDITIONAL_CRITERIA_LESS_THAN,
    greater_than_or_equal_to = c.LXW_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO,
    less_than_or_equal_to = c.LXW_CONDITIONAL_CRITERIA_LESS_THAN_OR_EQUAL_TO,
    between = c.LXW_CONDITIONAL_CRITERIA_BETWEEN,
    not_between = c.LXW_CONDITIONAL_CRITERIA_NOT_BETWEEN,
    text_containing = c.LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING,
    text_not_containing = c.LXW_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING,
    text_begins_with = c.LXW_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH,
    text_ends_with = c.LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH,
    time_period_yesterday = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY,
    time_period_today = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TODAY,
    time_period_tomorrow = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TOMORROW,
    time_period_last_7_days = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_7_DAYS,
    time_period_last_week = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_WEEK,
    time_period_this_week = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_WEEK,
    time_period_next_week = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_WEEK,
    time_period_last_month = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_MONTH,
    time_period_this_month = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_MONTH,
    time_period_next_month = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH,
    average_above = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE,
    average_below = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW,
    average_above_or_equal = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE_OR_EQUAL,
    average_below_or_equal = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL,
    average_1_std_dev_above = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_ABOVE,
    average_1_std_dev_below = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_BELOW,
    average_2_std_dev_above = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_ABOVE,
    average_2_std_dev_below = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_BELOW,
    average_3_std_dev_above = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_ABOVE,
    average_3_std_dev_below = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW,
    top_or_bottom_percent = c.LXW_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT,
};

pub const ConditionalFormatRuleTypes = enum(u8) {
    none = c.LXW_CONDITIONAL_RULE_TYPE_NONE,
    minimum = c.LXW_CONDITIONAL_RULE_TYPE_MINIMUM,
    number = c.LXW_CONDITIONAL_RULE_TYPE_NUMBER,
    percent = c.LXW_CONDITIONAL_RULE_TYPE_PERCENT,
    percentile = c.LXW_CONDITIONAL_RULE_TYPE_PERCENTILE,
    formula = c.LXW_CONDITIONAL_RULE_TYPE_FORMULA,
    maximum = c.LXW_CONDITIONAL_RULE_TYPE_MAXIMUM,
    auto_min = c.LXW_CONDITIONAL_RULE_TYPE_AUTO_MIN,
    auto_max = c.LXW_CONDITIONAL_RULE_TYPE_AUTO_MAX,
};

pub const ConditionalFormatBarDirection = enum(u8) {
    context = c.LXW_CONDITIONAL_BAR_DIRECTION_CONTEXT,
    right_to_left = c.LXW_CONDITIONAL_BAR_DIRECTION_RIGHT_TO_LEFT,
    left_to_right = c.LXW_CONDITIONAL_BAR_DIRECTION_LEFT_TO_RIGHT,
};

pub const ConditionalBarAxisPosition = enum(u8) {
    automatic = c.LXW_CONDITIONAL_BAR_AXIS_AUTOMATIC,
    midpoint = c.LXW_CONDITIONAL_BAR_AXIS_MIDPOINT,
    none = c.LXW_CONDITIONAL_BAR_AXIS_NONE,
};

pub const ConditionalIconTypes = enum(u8) {
    icons_3_arrows_colored = c.LXW_CONDITIONAL_ICONS_3_ARROWS_COLORED,
    icons_3_arrows_gray = c.LXW_CONDITIONAL_ICONS_3_ARROWS_GRAY,
    icons_3_flags = c.LXW_CONDITIONAL_ICONS_3_FLAGS,
    icons_3_traffic_lights_unrimmed = c.LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_UNRIMMED,
    icons_3_traffic_lights_rimmed = c.LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_RIMMED,
    icons_3_signs = c.LXW_CONDITIONAL_ICONS_3_SIGNS,
    icons_3_symbols_circled = c.LXW_CONDITIONAL_ICONS_3_SYMBOLS_CIRCLED,
    icons_3_symbols_uncircled = c.LXW_CONDITIONAL_ICONS_3_SYMBOLS_UNCIRCLED,
    icons_4_arrows_colored = c.LXW_CONDITIONAL_ICONS_4_ARROWS_COLORED,
    icons_4_arrows_gray = c.LXW_CONDITIONAL_ICONS_4_ARROWS_GRAY,
    icons_4_red_to_black = c.LXW_CONDITIONAL_ICONS_4_RED_TO_BLACK,
    icons_4_ratings = c.LXW_CONDITIONAL_ICONS_4_RATINGS,
    icons_4_traffic_lights = c.LXW_CONDITIONAL_ICONS_4_TRAFFIC_LIGHTS,
    icons_5_arrows_colored = c.LXW_CONDITIONAL_ICONS_5_ARROWS_COLORED,
    icons_5_arrows_gray = c.LXW_CONDITIONAL_ICONS_5_ARROWS_GRAY,
    icons_5_ratings = c.LXW_CONDITIONAL_ICONS_5_RATINGS,
    icons_5_quarters = c.LXW_CONDITIONAL_ICONS_5_QUARTERS,
};

pub const ConditionalFormat = struct {
    type: ConditionalFormatTypes = .none,
    criteria: ConditionalCriteria = .none,
    value: f64 = 0,
    value_string: ?CString = null,
    format: Format = .default,
    min_value: f64 = 0,
    min_value_string: ?CString = null,
    min_rule_type: ConditionalFormatRuleTypes = .none,
    min_color: Format.DefinedColors = .default,
    mid_value: f64 = 0,
    mid_value_string: ?CString = null,
    mid_rule_type: ConditionalFormatRuleTypes = .none,
    mid_color: Format.DefinedColors = .default,
    max_value: f64 = 0,
    max_value_string: ?CString = null,
    max_rule_type: ConditionalFormatRuleTypes = .none,
    max_color: Format.DefinedColors = .default,
    bar_color: Format.DefinedColors = .default,
    bar_only: bool = false,
    data_bar_2010: bool = false,
    bar_solid: bool = false,
    bar_negative_color: Format.DefinedColors = .default,
    bar_border_color: Format.DefinedColors = .default,
    bar_negative_border_color: Format.DefinedColors = .default,
    bar_negative_color_same: bool = false,
    bar_negative_border_color_same: bool = false,
    bar_no_border: bool = false,
    bar_direction: ConditionalFormatBarDirection = .context,
    bar_axis_position: ConditionalBarAxisPosition = .automatic,
    bar_axis_color: Format.DefinedColors = .default,
    icon_style: ConditionalIconTypes = .icons_3_arrows_colored,
    reverse_icons: bool = false,
    icons_only: bool = false,
    multi_range: ?CString = null,
    stop_if_true: bool = false,

    pub const default: ConditionalFormat = .{
        .type = .none,
        .criteria = .none,
        .value = 0,
        .value_string = null,
        .format = .default,
        .min_value = 0,
        .min_value_string = null,
        .min_rule_type = .none,
        .min_color = .default,
        .mid_value = 0,
        .mid_value_string = null,
        .mid_rule_type = .none,
        .mid_color = .default,
        .max_value = 0,
        .max_value_string = null,
        .max_rule_type = .none,
        .max_color = .default,
        .bar_color = .default,
        .bar_only = false,
        .data_bar_2010 = false,
        .bar_solid = false,
        .bar_negative_color = .default,
        .bar_border_color = .default,
        .bar_negative_border_color = .default,
        .bar_negative_color_same = false,
        .bar_negative_border_color_same = false,
        .bar_no_border = false,
        .bar_direction = .context,
        .bar_axis_position = .automatic,
        .bar_axis_color = .default,
        .icon_style = .icons_3_arrows_colored,
        .reverse_icons = false,
        .icons_only = false,
        .multi_range = null,
        .stop_if_true = false,
    };

    inline fn toC(self: ConditionalFormat) c.lxw_conditional_format {
        return c.lxw_conditional_format{
            .type = @intFromEnum(self.type),
            .criteria = @intFromEnum(self.criteria),
            .value = self.value,
            .value_string = self.value_string,
            .format = self.format.format_c,
            .min_value = self.min_value,
            .min_value_string = self.min_value_string,
            .min_rule_type = @intFromEnum(self.min_rule_type),
            .min_color = @intFromEnum(self.min_color),
            .mid_value = self.mid_value,
            .mid_value_string = self.mid_value_string,
            .mid_rule_type = @intFromEnum(self.mid_rule_type),
            .mid_color = @intFromEnum(self.mid_color),
            .max_value = self.max_value,
            .max_value_string = self.max_value_string,
            .max_rule_type = @intFromEnum(self.max_rule_type),
            .max_color = @intFromEnum(self.max_color),
            .bar_color = @intFromEnum(self.bar_color),
            .bar_only = @intFromBool(self.bar_only),
            .data_bar_2010 = @intFromBool(self.data_bar_2010),
            .bar_solid = @intFromBool(self.bar_solid),
            .bar_negative_color = @intFromEnum(self.bar_negative_color),
            .bar_border_color = @intFromEnum(self.bar_border_color),
            .bar_negative_border_color = @intFromEnum(self.bar_negative_border_color),
            .bar_negative_color_same = @intFromBool(self.bar_negative_color_same),
            .bar_negative_border_color_same = @intFromBool(self.bar_negative_border_color_same),
            .bar_no_border = @intFromBool(self.bar_no_border),
            .bar_direction = @intFromEnum(self.bar_direction),
            .bar_axis_position = @intFromEnum(self.bar_axis_position),
            .bar_axis_color = @intFromEnum(self.bar_axis_color),
            .icon_style = @intFromEnum(self.icon_style),
            .reverse_icons = @intFromBool(self.reverse_icons),
            .icons_only = @intFromBool(self.icons_only),
            .multi_range = self.multi_range,
            .stop_if_true = @intFromBool(self.stop_if_true),
        };
    }
};

// pub extern fn worksheet_conditional_format_range(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, conditional_format: [*c]lxw_conditional_format) lxw_error;
pub inline fn conditionalFormatRange(
    self: WorkSheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    conditional_format: ConditionalFormat,
) XlsxError!void {
    try check(c.worksheet_conditional_format_range(
        self.worksheet_c,
        first_row,
        first_col,
        last_row,
        last_col,
        @constCast(&conditional_format.toC()),
    ));
}

// pub extern fn worksheet_conditional_format_cell(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, conditional_format: [*c]lxw_conditional_format) lxw_error;

pub const TableStyleType = enum(u8) {
    default = c.LXW_TABLE_STYLE_TYPE_DEFAULT,
    light = c.LXW_TABLE_STYLE_TYPE_LIGHT,
    medium = c.LXW_TABLE_STYLE_TYPE_MEDIUM,
    dark = c.LXW_TABLE_STYLE_TYPE_DARK,
};
pub const TableTotalFunctions = enum(u8) {
    none = c.LXW_TABLE_FUNCTION_NONE,
    average = c.LXW_TABLE_FUNCTION_AVERAGE,
    count_nums = c.LXW_TABLE_FUNCTION_COUNT_NUMS,
    count = c.LXW_TABLE_FUNCTION_COUNT,
    max = c.LXW_TABLE_FUNCTION_MAX,
    min = c.LXW_TABLE_FUNCTION_MIN,
    std_dev = c.LXW_TABLE_FUNCTION_STD_DEV,
    sum = c.LXW_TABLE_FUNCTION_SUM,
    variance = c.LXW_TABLE_FUNCTION_VAR,
};

pub const TableColumn = struct {
    header: ?CString = null,
    formula: ?CString = null,
    total_string: ?CString = null,
    total_function: TableTotalFunctions = .none,
    header_format: Format = .default,
    format: Format = .default,
    total_value: f64 = 0,

    pub const empty: TableColumn = .{
        .header = null,
        .formula = null,
        .total_string = null,
        .total_function = .none,
        .header_format = .default,
        .format = .default,
        .total_value = 0,
    };

    inline fn toC(self: TableColumn) c.lxw_table_column {
        return c.lxw_table_column{
            .header = self.header,
            .formula = self.formula,
            .total_string = self.total_string,
            .total_function = @intFromEnum(self.total_function),
            .header_format = self.header_format.format_c,
            .format = self.format.format_c,
            .total_value = self.total_value,
        };
    }
};

pub const TableOptions = struct {
    name: ?CString = null,
    no_header_row: bool = false,
    no_autofilter: bool = false,
    no_banded_rows: bool = false,
    banded_columns: bool = false,
    first_column: bool = false,
    last_column: bool = false,
    style_type: TableStyleType = .default,
    style_type_number: u8 = 0,
    total_row: bool = false,
    columns: []const TableColumn = &.{},

    pub const default: TableOptions = .{
        .name = null,
        .no_header_row = false,
        .no_autofilter = false,
        .no_banded_rows = false,
        .banded_columns = false,
        .first_column = false,
        .last_column = false,
        .style_type = .default,
        .style_type_number = 0,
        .total_row = false,
        .columns = &.{},
    };

    inline fn toC(self: TableOptions, columns: [:null]const ?*c.lxw_table_column) c.lxw_table_options {
        return c.lxw_table_options{
            .name = self.name,
            .no_header_row = @intFromBool(self.no_header_row),
            .no_autofilter = @intFromBool(self.no_autofilter),
            .no_banded_rows = @intFromBool(self.no_banded_rows),
            .banded_columns = @intFromBool(self.banded_columns),
            .first_column = @intFromBool(self.first_column),
            .last_column = @intFromBool(self.last_column),
            .style_type = @intFromEnum(self.style_type),
            .style_type_number = self.style_type_number,
            .total_row = @intFromBool(self.total_row),
            .columns = @ptrCast(@constCast(columns)),
        };
    }
};

// pub extern fn worksheet_add_table(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, options: [*c]lxw_table_options) lxw_error;
pub inline fn addTable(
    self: WorkSheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    options: TableOptions,
) !void {
    const allocator: std.mem.Allocator = if (self.alloc) |allocator|
        allocator
    else
        return XlsxError.AddTable;

    // convert table_columns to [*c][*c]lxw_table_column
    var table_column_array = try allocator.alloc(c.lxw_table_column, options.columns.len);
    defer allocator.free(table_column_array);
    var table_column_c = try allocator.allocSentinel(?*c.lxw_table_column, options.columns.len, null);
    defer allocator.free(table_column_c);

    for (options.columns, 0..) |column, i| {
        table_column_array[i] = column.toC();
        table_column_c[i] = &table_column_array[i];
    }

    try check(c.worksheet_add_table(
        self.worksheet_c,
        first_row,
        first_col,
        last_row,
        last_col,
        @constCast(&options.toC(table_column_c)),
    ));
}

pub const TableColumnNoAlloc = extern struct {
    header: ?CString = null,
    formula: ?CString = null,
    total_string: ?CString = null,
    total_function: TableTotalFunctions = .none,
    header_format_c: ?*c.lxw_format = null,
    format_c: ?*c.lxw_format = null,
    total_value: f64 = 0,

    pub const default: TableColumnNoAlloc = .{
        .header = null,
        .formula = null,
        .total_string = null,
        .total_function = .none,
        .header_format_c = null,
        .format_c = null,
        .total_value = 0,
    };
};
pub const TableColumnNoAllocArray = [:null]const ?*const TableColumnNoAlloc;

pub const TableOptionsNoAlloc = struct {
    name: ?CString = null,
    no_header_row: Bool = Bool.false,
    no_autofilter: Bool = Bool.false,
    no_banded_rows: Bool = Bool.false,
    banded_columns: Bool = Bool.false,
    first_column: Bool = Bool.false,
    last_column: Bool = Bool.false,
    style_type: TableStyleType = .default,
    style_type_number: u8 = 0,
    total_row: Bool = Bool.false,
    columns: TableColumnNoAllocArray = &.{},

    pub const default: TableOptionsNoAlloc = .{
        .name = null,
        .no_header_row = Bool.false,
        .no_autofilter = Bool.false,
        .no_banded_rows = Bool.false,
        .banded_columns = Bool.false,
        .first_column = Bool.false,
        .last_column = Bool.false,
        .style_type = .default,
        .style_type_number = 0,
        .total_row = Bool.false,
        .columns = &.{},
    };

    inline fn toC(self: TableOptionsNoAlloc) c.lxw_table_options {
        return c.lxw_table_options{
            .name = self.name,
            .no_header_row = @intFromEnum(self.no_header_row),
            .no_autofilter = @intFromEnum(self.no_autofilter),
            .no_banded_rows = @intFromEnum(self.no_banded_rows),
            .banded_columns = @intFromEnum(self.banded_columns),
            .first_column = @intFromEnum(self.first_column),
            .last_column = @intFromEnum(self.last_column),
            .style_type = @intFromEnum(self.style_type),
            .style_type_number = self.style_type_number,
            .total_row = @intFromEnum(self.total_row),
            .columns = @ptrCast(@constCast(self.columns)),
        };
    }
};

// pub extern fn worksheet_add_table(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw _col_t, options: [*c]lxw_table_options) lxw_error;
pub inline fn addTableNoAlloc(
    self: WorkSheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    options: TableOptionsNoAlloc,
) XlsxError!void {
    try check(c.worksheet_add_table(
        self.worksheet_c,
        first_row,
        first_col,
        last_row,
        last_col,
        @constCast(&options.toC()),
    ));
}

pub const HeaderFooterOptions = c.lxw_header_footer_options;

// pub extern fn worksheet_set_header(worksheet: [*c]lxw_worksheet, string: [*c]const u8) lxw_error;
pub inline fn setHeader(self: WorkSheet, string: ?CString) XlsxError!void {
    try check(c.worksheet_set_header(self.worksheet_c, string));
}

// pub extern fn worksheet_set_footer(worksheet: [*c]lxw_worksheet, string: [*c]const u8) lxw_error;
pub inline fn setFooter(self: WorkSheet, string: ?CString) XlsxError!void {
    try check(c.worksheet_set_footer(self.worksheet_c, string));
}

// pub extern fn worksheet_set_header_opt(worksheet: [*c]lxw_worksheet, string: [*c]const u8, options: [*c]lxw_header_footer_options) lxw_error;
pub inline fn setHeaderOpt(
    self: WorkSheet,
    string: ?CString,
    options: HeaderFooterOptions,
) XlsxError!void {
    try check(c.worksheet_set_header_opt(
        self.worksheet_c,
        string,
        @constCast(&options),
    ));
}

// pub extern fn worksheet_set_footer_opt(worksheet: [*c]lxw_worksheet, string: [*c]const u8, options: [*c]lxw_header_footer_options) lxw_error;
pub inline fn setFooterOpt(
    self: WorkSheet,
    string: ?CString,
    options: HeaderFooterOptions,
) XlsxError!void {
    try check(c.worksheet_set_footer_opt(
        self.worksheet_c,
        string,
        @constCast(&options),
    ));
}

// pub extern fn worksheet_hide(worksheet: [*c]lxw_worksheet) void;
pub inline fn hide(self: WorkSheet) void {
    c.worksheet_hide(self.worksheet_c);
}

pub const Protection = extern struct {
    no_select_locked_cells: bool = false,
    no_select_unlocked_cells: bool = false,
    format_cells: bool = false,
    format_columns: bool = false,
    format_rows: bool = false,
    insert_columns: bool = false,
    insert_rows: bool = false,
    insert_hyperlinks: bool = false,
    delete_columns: bool = false,
    delete_rows: bool = false,
    sort: bool = false,
    autofilter: bool = false,
    pivot_tables: bool = false,
    scenarios: bool = false,
    objects: bool = false,
    no_content: bool = false,
    no_objects: bool = false,

    pub const default = Protection{
        .no_select_locked_cells = false,
        .no_select_unlocked_cells = false,
        .format_cells = false,
        .format_columns = false,
        .format_rows = false,
        .insert_columns = false,
        .insert_rows = false,
        .insert_hyperlinks = false,
        .delete_columns = false,
        .delete_rows = false,
        .sort = false,
        .autofilter = false,
        .pivot_tables = false,
        .scenarios = false,
        .objects = false,
        .no_content = false,
        .no_objects = false,
    };

    inline fn toC(self: Protection) c.lxw_protection {
        return c.lxw_protection{
            .no_select_locked_cells = @intFromBool(self.no_select_locked_cells),
            .no_select_unlocked_cells = @intFromBool(self.no_select_unlocked_cells),
            .format_cells = @intFromBool(self.format_cells),
            .format_columns = @intFromBool(self.format_columns),
            .format_rows = @intFromBool(self.format_rows),
            .insert_columns = @intFromBool(self.insert_columns),
            .insert_rows = @intFromBool(self.insert_rows),
            .insert_hyperlinks = @intFromBool(self.insert_hyperlinks),
            .delete_columns = @intFromBool(self.delete_columns),
            .delete_rows = @intFromBool(self.delete_rows),
            .sort = @intFromBool(self.sort),
            .autofilter = @intFromBool(self.autofilter),
            .pivot_tables = @intFromBool(self.pivot_tables),
            .scenarios = @intFromBool(self.scenarios),
            .objects = @intFromBool(self.objects),
            .no_content = @intFromBool(self.no_content),
            .no_objects = @intFromBool(self.no_objects),
        };
    }
};

// pub extern fn worksheet_protect(worksheet: [*c]lxw_worksheet, password: [*c]const u8, options: [*c]lxw_protection) void;
pub inline fn protect(
    self: WorkSheet,
    password: ?CString,
    options: Protection,
) void {
    c.worksheet_protect(
        self.worksheet_c,
        password,
        @constCast(&options.toC()),
    );
}

pub const ButtonOptions = extern struct {
    caption: ?CString = null,
    macro: ?CString = null,
    description: ?CString = null,
    width: u16 = 0,
    height: u16 = 0,
    x_scale: f64 = 0,
    y_scale: f64 = 0,
    x_offset: i32 = 0,
    y_offset: i32 = 0,

    pub const default = ButtonOptions{
        .caption = null,
        .macro = null,
        .description = null,
        .width = 0,
        .height = 0,
        .x_scale = 0,
        .y_scale = 0,
        .x_offset = 0,
        .y_offset = 0,
    };
};

// pub extern fn worksheet_insert_button(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, options: [*c]lxw_button_options) lxw_error;
pub inline fn insertButton(
    self: WorkSheet,
    row: u32,
    col: u16,
    options: ButtonOptions,
) XlsxError!void {
    try check(c.worksheet_insert_button(
        self.worksheet_c,
        row,
        col,
        @ptrCast(@constCast(&options)),
    ));
}

// pub extern fn worksheet_data_validation_range(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, validation: [*c]lxw_data_validation) lxw_error;
// pub extern fn worksheet_activate(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_select(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_set_first_sheet(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_print_across(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_gridlines(worksheet: [*c]lxw_worksheet, option: u8) void;
// pub extern fn worksheet_center_horizontally(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_center_vertically(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_print_row_col_headers(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_repeat_rows(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, last_row: lxw_row_t) lxw_error;
// pub extern fn worksheet_repeat_columns(worksheet: [*c]lxw_worksheet, first_col: lxw_col_t, last_col: lxw_col_t) lxw_error;
// pub extern fn worksheet_print_area(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t) lxw_error;
// pub extern fn worksheet_fit_to_pages(worksheet: [*c]lxw_worksheet, width: u16, height: u16) void;
// pub extern fn worksheet_print_black_and_white(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_right_to_left(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_hide_zero(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_outline_settings(worksheet: [*c]lxw_worksheet, visible: u8, symbols_below: u8, symbols_right: u8, auto_style: u8) void;
// pub extern fn worksheet_ignore_errors(worksheet: [*c]lxw_worksheet, @"type": u8, range: [*c]const u8) lxw_error;

const std = @import("std");
const c = @import("lxw");
const xwz = @import("xlsxwriter.zig");
const CString = xwz.CString;
const CStringArray = xwz.CStringArray;
const Bool = xwz.Boolean;
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;
const Format = @import("Format.zig");
const Chart = @import("Chart.zig");
const filter = @import("filter.zig");
