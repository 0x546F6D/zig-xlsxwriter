const WorkSheet = @This();

worksheet_c: ?*c.lxw_worksheet,

// set functions
pub const RowColOptions = c.lxw_row_col_options;

// pub extern fn worksheet_set_row(worksheet: [*c]lxw_worksheet, row: lxw_row_t, height: f64, format: [*c]lxw_format) lxw_error;
pub inline fn setRow(
    self: WorkSheet,
    row: u16,
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
        @constCast(&options),
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
        @constCast(&options),
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
        @constCast(&options),
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

// pub extern fn worksheet_set_background(worksheet: [*c]lxw_worksheet, filename: [*c]const u8) lxw_error;
// pub extern fn worksheet_set_background_buffer(worksheet: [*c]lxw_worksheet, image_buffer: [*c]const u8, image_size: usize) lxw_error;
// pub extern fn worksheet_set_selection(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t) lxw_error;
// pub extern fn worksheet_set_top_left_cell(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t) void;
// pub extern fn worksheet_set_landscape(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_set_portrait(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_set_page_view(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_set_paper(worksheet: [*c]lxw_worksheet, paper_type: u8) void;
// pub extern fn worksheet_set_zoom(worksheet: [*c]lxw_worksheet, scale: u16) void;
// pub extern fn worksheet_set_start_page(worksheet: [*c]lxw_worksheet, start_page: u16) void;
// pub extern fn worksheet_set_print_scale(worksheet: [*c]lxw_worksheet, scale: u16) void;
// pub extern fn worksheet_set_tab_color(worksheet: [*c]lxw_worksheet, color: lxw_color_t) void;
// pub extern fn worksheet_set_default_row(worksheet: [*c]lxw_worksheet, height: f64, hide_unused_rows: u8) void;
// pub extern fn worksheet_set_vba_name(worksheet: [*c]lxw_worksheet, name: [*c]const u8) lxw_error;
// pub extern fn worksheet_set_comments_author(worksheet: [*c]lxw_worksheet, author: [*c]const u8) void;
// pub extern fn worksheet_set_error_cell(worksheet: [*c]lxw_worksheet, object_props: [*c]lxw_object_properties, ref_id: u32) void;

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
    string: [*c]const u8,
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
    formula: [*c]const u8,
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
    formula: [*c]const u8,
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
    formula: [*c]const u8,
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
    formula: [*c]const u8,
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
    url: [*c]const u8,
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
    url: [*c]const u8,
    format: Format,
    string: [*c]const u8,
    tooltip: [*c]const u8,
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

pub const RichStringTuple = extern struct {
    format_c: ?*c.lxw_format = null,
    string: [*c]const u8,
};

pub const RichStringType: type = [:null]const ?*const RichStringTuple;
// pub extern fn worksheet_write_rich_string(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, rich_string: [*c][*c]lxw_rich_string_tuple, format: [*c]lxw_format) lxw_error;
pub inline fn writeRichString(
    self: WorkSheet,
    row: u32,
    col: u16,
    rich_string: RichStringType,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_rich_string(
        self.worksheet_c,
        row,
        col,
        @ptrCast(@constCast(rich_string)),
        // @ptrCast(rich_string_copy),
        format.format_c,
    ));
}

// pub extern fn worksheet_write_boolean(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, value: c_int, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_write_blank(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_write_formula_num(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: f64) lxw_error;
// pub extern fn worksheet_write_formula_str(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: [*c]const u8) lxw_error;
// pub extern fn worksheet_write_comment(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, string: [*c]const u8) lxw_error;
// pub extern fn worksheet_write_comment_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, string: [*c]const u8, options: [*c]lxw_comment_options) lxw_error;

// image functions
// pub extern fn worksheet_insert_image(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, filename: [*c]const u8) lxw_error;
pub inline fn insertImage(
    self: WorkSheet,
    row: u32,
    col: u16,
    filename: [*c]const u8,
) XlsxError!void {
    try check(c.worksheet_insert_image(
        self.worksheet_c,
        row,
        col,
        filename,
    ));
}

pub const ImageOptions = extern struct {
    x_offset: i32 = 0,
    y_offset: i32 = 0,
    x_scale: f64 = 0,
    y_scale: f64 = 0,
    object_position: u8 = 0,
    description: [*c]const u8 = null,
    decorative: u8 = 0,
    url: [*c]const u8 = null,
    tip: [*c]const u8 = null,
    cell_format_c: ?*c.lxw_format = null,
};
// pub extern fn worksheet_insert_image_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, filename: [*c]const u8, options: [*c]lxw_image_options) lxw_error;
pub inline fn insertImageOpt(
    self: WorkSheet,
    row: u32,
    col: u16,
    filename: [*c]const u8,
    options: ImageOptions,
) XlsxError!void {
    try check(c.worksheet_insert_image_opt(
        self.worksheet_c,
        row,
        col,
        filename,
        @ptrCast(@constCast(&options)),
    ));
}

// pub extern fn worksheet_insert_image_buffer(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, image_buffer: [*c]const u8, image_size: usize) lxw_error;
pub inline fn insertImageBuffer(
    self: WorkSheet,
    row: u32,
    col: u16,
    image_buffer: [*c]const u8,
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
    image_buffer: [*c]const u8,
    image_size: usize,
    options: ImageOptions,
) XlsxError!void {
    try check(c.worksheet_insert_image_buffer_opt(
        self.worksheet_c,
        row,
        col,
        image_buffer,
        image_size,
        @ptrCast(@constCast(&options)),
    ));
}

// pub extern fn worksheet_embed_image(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, filename: [*c]const u8) lxw_error;
pub inline fn embedImage(
    self: WorkSheet,
    row: u32,
    col: u16,
    filename: [*c]const u8,
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
    filename: [*c]const u8,
    options: ImageOptions,
) XlsxError!void {
    try check(c.worksheet_embed_image_opt(
        self.worksheet_c,
        row,
        col,
        filename,
        @ptrCast(@constCast(&options)),
    ));
}

// pub extern fn worksheet_embed_image_buffer(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, image_buffer: [*c]const u8, image_size: usize) lxw_error;
pub inline fn embedImageBuffer(
    self: WorkSheet,
    row: u32,
    col: u16,
    image_buffer: [*c]const u8,
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
    image_buffer: [*c]const u8,
    image_size: usize,
    options: ImageOptions,
) XlsxError!void {
    try check(c.worksheet_embed_image_buffer_opt(
        self.worksheet_c,
        row,
        col,
        image_buffer,
        image_size,
        @ptrCast(@constCast(&options)),
    ));
}

// chart functions
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

// filter functions
pub const autoFilter = filter.autoFilter;
pub const filterColumn = filter.filterColumn;
pub const filterColumn2 = filter.filterColumn2;
pub const filterList = filter.filterList;

// panes functions
// pub extern fn worksheet_freeze_panes(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t) void;
// pub extern fn worksheet_split_panes(worksheet: [*c]lxw_worksheet, vertical: f64, horizontal: f64) void;
// pub extern fn worksheet_freeze_panes_opt(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, top_row: lxw_row_t, left_col: lxw_col_t, @"type": u8) void;
// pub extern fn worksheet_split_panes_opt(worksheet: [*c]lxw_worksheet, vertical: f64, horizontal: f64, top_row: lxw_row_t, left_col: lxw_col_t) void;

// other functions
// pub extern fn worksheet_merge_range(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, string: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn mergeRange(
    self: WorkSheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    string: [*c]const u8,
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
    ignore_blank: u8 = 0,
    show_input: u8 = 0,
    show_error: u8 = 0,
    error_type: ValidationErrorTypes = .stop,
    dropdown: u8 = 0,
    value_number: f64 = 0,
    value_formula: [*c]const u8 = null,
    value_list: StringArray = &.{},
    value_datetime: c.lxw_datetime = .{},
    minimum_number: f64 = 0,
    minimum_formula: [*c]const u8 = null,
    minimum_datetime: c.lxw_datetime = .{},
    maximum_number: f64 = 0,
    maximum_formula: [*c]const u8 = null,
    maximum_datetime: c.lxw_datetime = .{},
    input_title: [*c]const u8 = null,
    input_message: [*c]const u8 = null,
    error_title: [*c]const u8 = null,
    error_message: [*c]const u8 = null,
};

// pub extern fn worksheet_data_validation_cell(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, validation: [*c]lxw_data_validation) lxw_error;
pub inline fn dataValidationCell(
    self: WorkSheet,
    row: u32,
    col: u16,
    validation: DataValidation,
) XlsxError!void {
    var data_validation: c.lxw_data_validation = .{
        .validate = @intFromEnum(validation.validate),
        .criteria = @intFromEnum(validation.criteria),
        .ignore_blank = validation.ignore_blank,
        .show_input = validation.show_input,
        .show_error = validation.show_error,
        .error_type = @intFromEnum(validation.error_type),
        .dropdown = validation.dropdown,
        .value_number = validation.value_number,
        .value_formula = validation.value_formula,
        .value_list = @ptrCast(@constCast(validation.value_list)),
        .value_datetime = validation.value_datetime,
        .minimum_number = validation.minimum_number,
        .minimum_formula = validation.minimum_formula,
        .minimum_datetime = validation.minimum_datetime,
        .maximum_number = validation.maximum_number,
        .maximum_formula = validation.maximum_formula,
        .maximum_datetime = validation.maximum_datetime,
        .input_title = validation.input_title,
        .input_message = validation.input_message,
        .error_title = validation.error_title,
        .error_message = validation.error_message,
    };
    try check(c.worksheet_data_validation_cell(
        self.worksheet_c,
        row,
        col,
        &data_validation,
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

pub const ConditionalFormat = extern struct {
    type: ConditionalFormatTypes = .none,
    criteria: ConditionalCriteria = .none,
    value: f64 = 0,
    value_string: [*c]const u8 = null,
    format_c: ?*c.lxw_format = null,
    min_value: f64 = 0,
    min_value_string: [*c]const u8 = null,
    min_rule_type: ConditionalFormatRuleTypes = .none,
    min_color: Format.DefinedColors = @enumFromInt(0),
    mid_value: f64 = 0,
    mid_value_string: [*c]const u8 = null,
    mid_rule_type: ConditionalFormatRuleTypes = .none,
    mid_color: Format.DefinedColors = @enumFromInt(0),
    max_value: f64 = 0,
    max_value_string: [*c]const u8 = null,
    max_rule_type: ConditionalFormatRuleTypes = .none,
    max_color: Format.DefinedColors = @enumFromInt(0),
    bar_color: Format.DefinedColors = @enumFromInt(0),
    bar_only: u8 = 0,
    data_bar_2010: u8 = 0,
    bar_solid: u8 = 0,
    bar_negative_color: Format.DefinedColors = @enumFromInt(0),
    bar_border_color: Format.DefinedColors = @enumFromInt(0),
    bar_negative_border_color: Format.DefinedColors = @enumFromInt(0),
    bar_negative_color_same: u8 = 0,
    bar_negative_border_color_same: u8 = 0,
    bar_no_border: u8 = 0,
    bar_direction: ConditionalFormatBarDirection = .context,
    bar_axis_position: ConditionalBarAxisPosition = .automatic,
    bar_axis_color: Format.DefinedColors = @enumFromInt(0),
    icon_style: ConditionalIconTypes = .icons_3_arrows_colored,
    reverse_icons: u8 = 0,
    icons_only: u8 = 0,
    multi_range: [*c]const u8 = null,
    stop_if_true: u8 = 0,

    pub const empty: ConditionalFormat = .{
        .type = .none,
        .criteria = .none,
        .value = 0,
        .value_string = null,
        .format_c = null,
        .min_value = 0,
        .min_value_string = null,
        .min_rule_type = .none,
        .min_color = @enumFromInt(0),
        .mid_value = 0,
        .mid_value_string = null,
        .mid_rule_type = .none,
        .mid_color = @enumFromInt(0),
        .max_value = 0,
        .max_value_string = null,
        .max_rule_type = .none,
        .max_color = @enumFromInt(0),
        .bar_color = @enumFromInt(0),
        .bar_only = 0,
        .data_bar_2010 = 0,
        .bar_solid = 0,
        .bar_negative_color = @enumFromInt(0),
        .bar_border_color = @enumFromInt(0),
        .bar_negative_border_color = @enumFromInt(0),
        .bar_negative_color_same = 0,
        .bar_negative_border_color_same = 0,
        .bar_no_border = 0,
        .bar_direction = .context,
        .bar_axis_position = .automatic,
        .bar_axis_color = @enumFromInt(0),
        .icon_style = .icons_3_arrows_colored,
        .reverse_icons = 0,
        .icons_only = 0,
        .multi_range = null,
        .stop_if_true = 0,
    };
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
        @ptrCast(@constCast(&conditional_format)),
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

pub const TableColumn = extern struct {
    header: [*c]const u8 = null,
    formula: [*c]const u8 = null,
    total_string: [*c]const u8 = null,
    total_function: TableTotalFunctions = .none,
    header_format_c: ?*c.lxw_format = null,
    format_c: ?*c.lxw_format = null,
    total_value: f64 = 0,

    pub const empty: TableColumn = .{
        .header = null,
        .formula = null,
        .total_string = null,
        .total_function = .none,
        .header_format_c = null,
        .format_c = null,
        .total_value = 0,
    };
};
pub const TableColumnArray = [:null]const ?*TableColumn;
pub const TableOptions = struct {
    name: [*c]const u8 = null,
    no_header_row: u8 = 0,
    no_autofilter: u8 = 0,
    no_banded_rows: u8 = 0,
    banded_columns: u8 = 0,
    first_column: u8 = 0,
    last_column: u8 = 0,
    style_type: TableStyleType = .default,
    style_type_number: u8 = 0,
    total_row: u8 = 0,
    columns: TableColumnArray = &.{},

    pub const empty: TableOptions = .{
        .name = null,
        .no_header_row = 0,
        .no_autofilter = 0,
        .no_banded_rows = 0,
        .banded_columns = 0,
        .first_column = 0,
        .last_column = 0,
        .style_type = .default,
        .style_type_number = 0,
        .total_row = 0,
        .columns = &.{},
    };
};

// pub extern fn worksheet_add_table(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, options: [*c]lxw_table_options) lxw_error;
pub inline fn addTable(
    self: WorkSheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    options: TableOptions,
) XlsxError!void {
    var table_options: c.lxw_table_options = .{
        .name = options.name,
        .no_header_row = options.no_header_row,
        .no_autofilter = options.no_autofilter,
        .no_banded_rows = options.no_banded_rows,
        .banded_columns = options.banded_columns,
        .first_column = options.first_column,
        .last_column = options.last_column,
        .style_type = @intFromEnum(options.style_type),
        .style_type_number = options.style_type_number,
        .total_row = options.total_row,
        .columns = @ptrCast(@constCast(options.columns)),
    };
    try check(c.worksheet_add_table(
        self.worksheet_c,
        first_row,
        first_col,
        last_row,
        last_col,
        &table_options,
    ));
}

pub const HeaderFooterOptions = c.lxw_header_footer_options;

// pub extern fn worksheet_set_header(worksheet: [*c]lxw_worksheet, string: [*c]const u8) lxw_error;
pub inline fn setHeader(self: WorkSheet, string: [*c]const u8) XlsxError!void {
    try check(c.worksheet_set_header(self.worksheet_c, string));
}

// pub extern fn worksheet_set_footer(worksheet: [*c]lxw_worksheet, string: [*c]const u8) lxw_error;
pub inline fn setFooter(self: WorkSheet, string: [*c]const u8) XlsxError!void {
    try check(c.worksheet_set_footer(self.worksheet_c, string));
}

// pub extern fn worksheet_set_header_opt(worksheet: [*c]lxw_worksheet, string: [*c]const u8, options: [*c]lxw_header_footer_options) lxw_error;
pub inline fn setHeaderOpt(
    self: WorkSheet,
    string: [*c]const u8,
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
    string: [*c]const u8,
    options: HeaderFooterOptions,
) XlsxError!void {
    try check(c.worksheet_set_footer_opt(
        self.worksheet_c,
        string,
        @constCast(&options),
    ));
}

// pub extern fn worksheet_data_validation_range(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, validation: [*c]lxw_data_validation) lxw_error;
// pub extern fn worksheet_insert_button(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, options: [*c]lxw_button_options) lxw_error;
// pub extern fn worksheet_activate(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_select(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_hide(worksheet: [*c]lxw_worksheet) void;
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
// pub extern fn worksheet_protect(worksheet: [*c]lxw_worksheet, password: [*c]const u8, options: [*c]lxw_protection) void;
// pub extern fn worksheet_outline_settings(worksheet: [*c]lxw_worksheet, visible: u8, symbols_below: u8, symbols_right: u8, auto_style: u8) void;
// pub extern fn worksheet_show_comments(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_ignore_errors(worksheet: [*c]lxw_worksheet, @"type": u8, range: [*c]const u8) lxw_error;

const c = @import("xlsxwriter_c");
const StringArray = @import("xlsxwriter").StringArray;
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;
const Format = @import("Format.zig");
const Chart = @import("Chart.zig");
const filter = @import("filter.zig");
