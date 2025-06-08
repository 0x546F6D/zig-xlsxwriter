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
// pub extern fn worksheet_set_row_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, height: f64, format: [*c]lxw_format, options: [*c]lxw_row_col_options) lxw_error;
pub inline fn setRow(
    self: WorkSheet,
    row: u32,
    height: f64,
    format: Format,
    options: ?RowColOptions,
) XlsxError!void {
    try check(if (options) |rc_options|
        c.worksheet_set_row_opt(
            self.worksheet_c,
            row,
            height,
            format.format_c,
            @constCast(&rc_options.toC()),
        )
    else
        c.worksheet_set_row(
            self.worksheet_c,
            row,
            height,
            format.format_c,
        ));
}

// pub extern fn worksheet_set_row_pixels(worksheet: [*c]lxw_worksheet, row: lxw_row_t, pixels: u32, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_set_row_pixels_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, pixels: u32, format: [*c]lxw_format, options: [*c]lxw_row_col_options) lxw_error;
pub inline fn setRowPixels(
    self: WorkSheet,
    row: u32,
    pixels: u32,
    format: Format,
    options: ?RowColOptions,
) XlsxError!void {
    try check(if (options) |rc_options|
        c.worksheet_set_row_pixels_opt(
            self.worksheet_c,
            row,
            pixels,
            format.format_c,
            @constCast(&rc_options.toC()),
        )
    else
        c.worksheet_set_row_pixels(
            self.worksheet_c,
            row,
            pixels,
            format.format_c,
        ));
}

// pub extern fn worksheet_set_column(worksheet: [*c]lxw_worksheet, first_col: lxw_col_t, last_col: lxw_col_t, width: f64, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_set_column_opt(worksheet: [*c]lxw_worksheet, first_col: lxw_col_t, last_col: lxw_col_t, width: f64, format: [*c]lxw_format, options: [*c]lxw_row_col_options) lxw_error;
pub inline fn setColumn(
    self: WorkSheet,
    cols: Cols,
    width: f64,
    format: Format,
    options: ?RowColOptions,
) XlsxError!void {
    try check(if (options) |rc_options|
        c.worksheet_set_column_opt(
            self.worksheet_c,
            cols.first,
            cols.last,
            width,
            format.format_c,
            @constCast(&rc_options.toC()),
        )
    else
        c.worksheet_set_column(
            self.worksheet_c,
            cols.first,
            cols.last,
            width,
            format.format_c,
        ));
}

// pub extern fn worksheet_set_column_pixels(worksheet: [*c]lxw_worksheet, first_col: lxw_col_t, last_col: lxw_col_t, pixels: u32, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_set_column_pixels_opt(worksheet: [*c]lxw_worksheet, first_col: lxw_col_t, last_col: lxw_col_t, pixels: u32, format: [*c]lxw_format, options: [*c]lxw_row_col_options) lxw_error;
pub inline fn setColumnPixels(
    self: WorkSheet,
    cols: Cols,
    pixels: u32,
    format: Format,
    options: ?RowColOptions,
) XlsxError!void {
    try check(if (options) |rc_options|
        c.worksheet_set_column_pixels_opt(
            self.worksheet_c,
            cols.first,
            cols.last,
            pixels,
            format.format_c,
            @constCast(&rc_options.toC()),
        )
    else
        c.worksheet_set_column_pixels(
            self.worksheet_c,
            cols.first,
            cols.last,
            pixels,
            format.format_c,
        ));
}

// pub extern fn worksheet_set_h_pagebreaks(worksheet: [*c]lxw_worksheet, breaks: [*c]lxw_row_t) lxw_error;
pub inline fn setHPageBreaks(
    self: WorkSheet,
    breaks: [:0]const u32,
) XlsxError!void {
    try check(c.worksheet_set_h_pagebreaks(
        self.worksheet_c,
        @constCast(breaks),
    ));
}

// pub extern fn worksheet_set_v_pagebreaks(worksheet: [*c]lxw_worksheet, breaks: [*c]lxw_col_t) lxw_error;
pub inline fn setVPageBreaks(
    self: WorkSheet,
    breaks: [:0]const u16,
) XlsxError!void {
    try check(c.worksheet_set_v_pagebreaks(
        self.worksheet_c,
        @constCast(breaks),
    ));
}

// pub extern fn worksheet_set_margins(worksheet: [*c]lxw_worksheet, left: f64, right: f64, top: f64, bottom: f64) void;
pub inline fn setMargins(
    self: WorkSheet,
    left: f64,
    right: f64,
    top: f64,
    bottom: f64,
) void {
    c.worksheet_set_margins(
        self.worksheet_c,
        left,
        right,
        top,
        bottom,
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
    range: Range,
) XlsxError!void {
    try check(c.worksheet_set_selection(
        self.worksheet_c,
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
    ));
}

// pub extern fn worksheet_set_top_left_cell(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t) void;
pub inline fn setTopLeftCell(
    self: WorkSheet,
    cell: Cell,
) void {
    c.worksheet_set_top_left_cell(self.worksheet_c, cell.row, cell.col);
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
pub inline fn setVbaName(self: WorkSheet, name: [:0]const u8) XlsxError!void {
    try check(c.worksheet_set_vba_name(self.worksheet_c, name));
}

// pub extern fn worksheet_set_comments_author(worksheet: [*c]lxw_worksheet, author: [*c]const u8) void;
pub inline fn setCommentsAuthor(self: WorkSheet, author: [:0]const u8) void {
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
    cell: Cell,
    number: f64,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_number(
        self.worksheet_c,
        cell.row,
        cell.col,
        number,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_string(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, string: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn writeString(
    self: WorkSheet,
    cell: Cell,
    string: [:0]const u8,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_string(
        self.worksheet_c,
        cell.row,
        cell.col,
        string,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_formula(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn writeFormula(
    self: WorkSheet,
    cell: Cell,
    formula: [:0]const u8,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_formula(
        self.worksheet_c,
        cell.row,
        cell.col,
        formula,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_formula_num(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: f64) lxw_error;
pub inline fn writeFormulaNum(
    self: WorkSheet,
    cell: Cell,
    formula: [:0]const u8,
    format: Format,
    result: f64,
) XlsxError!void {
    try check(c.worksheet_write_formula_num(
        self.worksheet_c,
        cell.row,
        cell.col,
        formula,
        format.format_c,
        result,
    ));
}

// pub extern fn worksheet_write_formula_str(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: [*c]const u8) lxw_error;
pub inline fn writeFormulaStr(
    self: WorkSheet,
    cell: Cell,
    formula: [:0]const u8,
    format: Format,
    result: [:0]const u8,
) XlsxError!void {
    try check(c.worksheet_write_formula_str(
        self.worksheet_c,
        cell.row,
        cell.col,
        formula,
        format.format_c,
        result,
    ));
}

// pub extern fn worksheet_write_array_formula(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn writeArrayFormula(
    self: WorkSheet,
    range: Range,
    formula: [:0]const u8,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_array_formula(
        self.worksheet_c,
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
        formula,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_array_formula_num(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: f64) lxw_error;
pub inline fn writeArrayFormulaNum(
    self: WorkSheet,
    range: Range,
    formula: [:0]const u8,
    format: Format,
    result: f64,
) XlsxError!void {
    try check(c.worksheet_write_array_formula_num(
        self.worksheet_c,
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
        formula,
        format.format_c,
        result,
    ));
}

// pub extern fn worksheet_write_dynamic_formula(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn writeDynamicFormula(
    self: WorkSheet,
    cell: Cell,
    formula: [:0]const u8,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_dynamic_formula(
        self.worksheet_c,
        cell.row,
        cell.col,
        formula,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_dynamic_formula_num(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: f64) lxw_error;
pub inline fn writeDynamicFormulaNum(
    self: WorkSheet,
    cell: Cell,
    formula: [:0]const u8,
    format: Format,
    result: f64,
) XlsxError!void {
    try check(c.worksheet_write_dynamic_formula_num(
        self.worksheet_c,
        cell.row,
        cell.col,
        formula,
        format.format_c,
        result,
    ));
}

// pub extern fn worksheet_write_dynamic_array_formula(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn writeDynamicArrayFormula(
    self: WorkSheet,
    range: Range,
    formula: [:0]const u8,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_dynamic_array_formula(
        self.worksheet_c,
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
        formula,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_dynamic_array_formula_num(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: f64) lxw_error;
pub inline fn writeDynamicArrayFormulaNum(
    self: WorkSheet,
    range: Range,
    formula: [:0]const u8,
    format: Format,
    result: f64,
) XlsxError!void {
    try check(c.worksheet_write_dynamic_array_formula_num(
        self.worksheet_c,
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
        formula,
        format.format_c,
        result,
    ));
}

// pub extern fn worksheet_write_boolean(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, value: c_int, format: [*c]lxw_format) lxw_error;
pub inline fn writeBoolean(
    self: WorkSheet,
    cell: Cell,
    value: bool,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_boolean(
        self.worksheet_c,
        cell.row,
        cell.col,
        @intFromBool(value),
        format.format_c,
    ));
}

// pub extern fn worksheet_write_blank(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, format: [*c]lxw_format) lxw_error;
pub inline fn writeBlank(
    self: WorkSheet,
    cell: Cell,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_blank(
        self.worksheet_c,
        cell.row,
        cell.col,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_datetime(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, datetime: [*c]lxw_datetime, format: [*c]lxw_format) lxw_error;
pub inline fn writeDateTime(
    self: WorkSheet,
    cell: Cell,
    datetime: c.lxw_datetime,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_datetime(
        self.worksheet_c,
        cell.row,
        cell.col,
        @constCast(&datetime),
        format.format_c,
    ));
}

// pub extern fn worksheet_write_unixtime(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, unixtime: i64, format: [*c]lxw_format) lxw_error;
pub inline fn writeUnixTime(
    self: WorkSheet,
    cell: Cell,
    unixtime: i64,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_unixtime(
        self.worksheet_c,
        cell.row,
        cell.col,
        unixtime,
        format.format_c,
    ));
}

pub const UrlOptions = struct {
    string: ?[*:0]const u8,
    tooltip: ?[*:0]const u8,
};
// pub extern fn worksheet_write_url(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, url: [*c]const u8, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_write_url_opt(worksheet: [*c]lxw_worksheet, row_num: lxw_row_t, col_num: lxw_col_t, url: [*c]const u8, format: [*c]lxw_format, string: [*c]const u8, tooltip: [*c]const u8) lxw_error;
pub inline fn writeUrl(
    self: WorkSheet,
    cell: Cell,
    url: [:0]const u8,
    format: Format,
    options: ?UrlOptions,
) XlsxError!void {
    try check(if (options) |url_options|
        c.worksheet_write_url_opt(
            self.worksheet_c,
            cell.row,
            cell.col,
            url,
            format.format_c,
            url_options.string,
            url_options.tooltip,
        )
    else
        c.worksheet_write_url(
            self.worksheet_c,
            cell.row,
            cell.col,
            url,
            format.format_c,
        ));
}

pub const RichStringTuple = struct {
    format: Format = .default,
    string: [:0]const u8,

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
    cell: Cell,
    rich_string: []const RichStringTuple,
    format: Format,
) !void {
    const allocator: std.mem.Allocator = if (self.alloc) |allocator|
        allocator
    else
        return XlsxError.WriteRichStringNoAlloc;

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
        cell.row,
        cell.col,
        @ptrCast(rich_string_c),
        format.format_c,
    ));
}

pub const RichStringTupleNoAlloc = extern struct {
    format_c: ?*c.lxw_format = null,
    string: ?[*:0]const u8 = null,
};

pub const RichStringNoAllocArray: type = [:null]const ?*const RichStringTupleNoAlloc;
// pub extern fn worksheet_write_rich_string(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, rich_string: [*c][*c]lxw_rich_string_t uple, format: [*c]lxw_format) lxw_error;
pub inline fn writeRichStringNoAlloc(
    self: WorkSheet,
    cell: Cell,
    rich_string: RichStringNoAllocArray,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_rich_string(
        self.worksheet_c,
        cell.row,
        cell.col,
        @ptrCast(@constCast(rich_string)),
        format.format_c,
    ));
}

// pub extern fn worksheet_set_background(worksheet: [*c]lxw_worksheet, filename: [*c]const u8) lxw_error;
pub inline fn setBackGround(
    self: WorkSheet,
    filename: [:0]const u8,
) XlsxError!void {
    try check(c.worksheet_set_background(
        self.worksheet_c,
        filename,
    ));
}

// pub extern fn worksheet_set_background_buffer(worksheet: [*c]lxw_worksheet, image_buffer: [*c]const u8, image_size: usize) lxw_error;
pub inline fn setBackGroundBuffer(
    self: WorkSheet,
    image_buffer: [:0]const u8,
) XlsxError!void {
    try check(c.worksheet_set_background_buffer(
        self.worksheet_c,
        image_buffer,
        image_buffer.len,
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

// pub extern fn worksheet_insert_chart(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, chart: [*c]lxw_chart) lxw_error;
// pub extern fn worksheet_insert_chart_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, chart: [*c]lxw_chart, user_options: [*c]lxw_chart_options) lxw_error;
pub inline fn insertChart(
    self: WorkSheet,
    cell: Cell,
    chart: Chart,
    user_options: ?ChartOptions,
) XlsxError!void {
    try check(if (user_options) |chart_options|
        c.worksheet_insert_chart_opt(
            self.worksheet_c,
            cell.row,
            cell.col,
            chart.chart_c,
            @constCast(&chart_options.toC()),
        )
    else
        c.worksheet_insert_chart(
            self.worksheet_c,
            cell.row,
            cell.col,
            chart.chart_c,
        ));
}

// panes functions
// pub extern fn worksheet_freeze_panes(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t) void;
pub inline fn freezePanes(
    self: WorkSheet,
    cell: Cell,
) void {
    c.worksheet_freeze_panes(
        self.worksheet_c,
        cell.row,
        cell.col,
    );
}

// pub extern fn worksheet_freeze_panes_opt(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, top_row: lxw_row_t, left_col: lxw_col_t, @"type": u8) void;
pub inline fn freezePanesOpt(
    self: WorkSheet,
    range: Range,
    @"type": bool,
) void {
    c.worksheet_freeze_panes_opt(
        self.worksheet_c,
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
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
    cell: Cell,
) void {
    c.worksheet_split_panes_opt(
        self.worksheet_c,
        vertical,
        horizontal,
        cell.row,
        cell.col,
    );
}

// other functions
// pub extern fn worksheet_merge_range(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, string: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn mergeRange(
    self: WorkSheet,
    range: Range,
    string: ?[*:0]const u8,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_merge_range(
        self.worksheet_c,
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
        string,
        format.format_c,
    ));
}

pub const HeaderFooterOptions = c.lxw_header_footer_options;

// pub extern fn worksheet_set_header(worksheet: [*c]lxw_worksheet, string: [*c]const u8) lxw_error;
// pub extern fn worksheet_set_header_opt(worksheet: [*c]lxw_worksheet, string: [*c]const u8, options: [*c]lxw_header_footer_options) lxw_error;
pub inline fn setHeader(
    self: WorkSheet,
    string: [:0]const u8,
    options: ?HeaderFooterOptions,
) XlsxError!void {
    try check(if (options) |hf_options|
        c.worksheet_set_header_opt(
            self.worksheet_c,
            string,
            @constCast(&hf_options),
        )
    else
        c.worksheet_set_header(
            self.worksheet_c,
            string,
        ));
}

// pub extern fn worksheet_set_footer(worksheet: [*c]lxw_worksheet, string: [*c]const u8) lxw_error;
// pub extern fn worksheet_set_footer_opt(worksheet: [*c]lxw_worksheet, string: [*c]const u8, options: [*c]lxw_header_footer_options) lxw_error;
pub inline fn setFooter(
    self: WorkSheet,
    string: [:0]const u8,
    options: ?HeaderFooterOptions,
) XlsxError!void {
    try check(if (options) |hf_options|
        c.worksheet_set_footer_opt(
            self.worksheet_c,
            string,
            @constCast(&hf_options),
        )
    else
        c.worksheet_set_footer(
            self.worksheet_c,
            string,
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
    password: ?[*:0]const u8,
    options: Protection,
) void {
    c.worksheet_protect(
        self.worksheet_c,
        password,
        @constCast(&options.toC()),
    );
}

pub const ButtonOptions = extern struct {
    caption: ?[*:0]const u8 = null,
    macro: ?[*:0]const u8 = null,
    description: ?[*:0]const u8 = null,
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
    cell: Cell,
    options: ButtonOptions,
) XlsxError!void {
    try check(c.worksheet_insert_button(
        self.worksheet_c,
        cell.row,
        cell.col,
        @ptrCast(@constCast(&options)),
    ));
}

pub const IgnoreErrors = enum(u8) {
    number_stored_as_text = c.LXW_IGNORE_NUMBER_STORED_AS_TEXT,
    eval_error = c.LXW_IGNORE_EVAL_ERROR,
    formula_differs = c.LXW_IGNORE_FORMULA_DIFFERS,
    formula_range = c.LXW_IGNORE_FORMULA_RANGE,
    formula_unlocked = c.LXW_IGNORE_FORMULA_UNLOCKED,
    empty_cell_reference = c.LXW_IGNORE_EMPTY_CELL_REFERENCE,
    list_data_validation = c.LXW_IGNORE_LIST_DATA_VALIDATION,
    calculated_column = c.LXW_IGNORE_CALCULATED_COLUMN,
    two_digit_text_year = c.LXW_IGNORE_TWO_DIGIT_TEXT_YEAR,
    last_option = c.LXW_IGNORE_LAST_OPTION,
};

// pub extern fn worksheet_ignore_errors(worksheet: [*c]lxw_worksheet, @"type": u8, range: [*c]const u8) lxw_error;
pub inline fn ignoreErrors(
    self: WorkSheet,
    @"type": IgnoreErrors,
    range: [:0]const u8,
) XlsxError!void {
    try check(c.worksheet_ignore_errors(
        self.worksheet_c,
        @intFromEnum(@"type"),
        range,
    ));
}

// pub extern fn worksheet_activate(worksheet: [*c]lxw_worksheet) void;
pub inline fn activate(self: WorkSheet) XlsxError!void {
    c.worksheet_activate(self.worksheet_c);
}

// pub extern fn worksheet_select(worksheet: [*c]lxw_worksheet) void;
pub inline fn select(self: WorkSheet) XlsxError!void {
    c.worksheet_select(self.worksheet_c);
}

// pub extern fn worksheet_set_first_sheet(worksheet: [*c]lxw_worksheet) void;
pub inline fn setFirstSheet(self: WorkSheet) XlsxError!void {
    c.worksheet_set_first_sheet(self.worksheet_c);
}

// pub extern fn worksheet_print_across(worksheet: [*c]lxw_worksheet) void;
pub inline fn printAcross(self: WorkSheet) XlsxError!void {
    c.worksheet_print_across(self.worksheet_c);
}

pub const GridLines = enum(u8) {
    hide_all = c.LXW_HIDE_ALL_GRIDLINES,
    show_screen = c.LXW_SHOW_SCREEN_GRIDLINES,
    show_print = c.LXW_SHOW_PRINT_GRIDLINES,
    show_all = c.LXW_SHOW_ALL_GRIDLINES,
};
// pub extern fn worksheet_gridlines(worksheet: [*c]lxw_worksheet, option: u8) void;
pub inline fn gridlines(self: WorkSheet, option: GridLines) XlsxError!void {
    c.worksheet_gridlines(self.worksheet_c, @intFromEnum(option));
}

// pub extern fn worksheet_center_horizontally(worksheet: [*c]lxw_worksheet) void;
pub inline fn centerHorizontally(self: WorkSheet) XlsxError!void {
    c.worksheet_center_horizontally(self.worksheet_c);
}

// pub extern fn worksheet_center_vertically(worksheet: [*c]lxw_worksheet) void;
pub inline fn centerVertically(self: WorkSheet) XlsxError!void {
    c.worksheet_center_vertically(self.worksheet_c);
}

// pub extern fn worksheet_print_row_col_headers(worksheet: [*c]lxw_worksheet) void;
pub inline fn printRowColHeaders(self: WorkSheet) XlsxError!void {
    c.worksheet_print_row_col_headers(self.worksheet_c);
}

// pub extern fn worksheet_repeat_rows(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, last_row: lxw_row_t) lxw_error;
pub inline fn repeatRows(
    self: WorkSheet,
    rows: Rows,
) XlsxError!void {
    try check(c.worksheet_repeat_rows(
        self.worksheet_c,
        rows.first,
        rows.last,
    ));
}

// pub extern fn worksheet_repeat_columns(worksheet: [*c]lxw_worksheet, first_col: lxw_col_t, last_col: lxw_col_t) lxw_error;
pub inline fn repeatCols(
    self: WorkSheet,
    cols: Cols,
) XlsxError!void {
    try check(c.worksheet_repeat_columns(
        self.worksheet_c,
        cols.first,
        cols.last,
    ));
}

// pub extern fn worksheet_print_area(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t) lxw_error;
pub inline fn printArea(
    self: WorkSheet,
    range: Range,
) XlsxError!void {
    try check(c.worksheet_print_area(
        self.worksheet_c,
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
    ));
}

// pub extern fn worksheet_fit_to_pages(worksheet: [*c]lxw_worksheet, width: u16, height: u16) void;
pub inline fn fitToPages(
    self: WorkSheet,
    width: u16,
    height: u16,
) XlsxError!void {
    try check(c.worksheet_fit_to_pages(
        self.worksheet_c,
        width,
        height,
    ));
}
// pub extern fn worksheet_print_black_and_white(worksheet: [*c]lxw_worksheet) void;
pub inline fn printBlackAndWhite(self: WorkSheet) XlsxError!void {
    c.worksheet_print_black_and_white(self.worksheet_c);
}

// pub extern fn worksheet_right_to_left(worksheet: [*c]lxw_worksheet) void;
pub inline fn rightToLeft(self: WorkSheet) XlsxError!void {
    c.worksheet_right_to_left(self.worksheet_c);
}

// pub extern fn worksheet_hide_zero(worksheet: [*c]lxw_worksheet) void;
pub inline fn hideZero(self: WorkSheet) XlsxError!void {
    c.worksheet_hide_zero(self.worksheet_c);
}

// pub extern fn worksheet_outline_settings(worksheet: [*c]lxw_worksheet, visible: u8, symbols_below: u8, symbols_right: u8, auto_style: u8) void;
pub inline fn outlineSettings(
    self: WorkSheet,
    visible: bool,
    symbols_below: bool,
    symbols_right: bool,
    auto_style: bool,
) XlsxError!void {
    try check(c.worksheet_outline_settings(
        self.worksheet_c,
        @intFromBool(visible),
        @intFromBool(symbols_below),
        @intFromBool(symbols_right),
        @intFromBool(auto_style),
    ));
}

// filter functions
const filter = @import("worksheet/filter.zig");
pub const autoFilter = filter.autoFilter;
pub const filterColumn = filter.filterColumn;
pub const filterColumn2 = filter.filterColumn2;
pub const filterList = filter.filterList;

// table functions
const table = @import("worksheet/table.zig");
pub const addTable = table.addTable;
pub const addTableNoAlloc = table.addTableNoAlloc;

// image functions
const image = @import("worksheet/image.zig");
pub const insertImage = image.insertImage;
pub const insertImageBuffer = image.insertImageBuffer;
pub const embedImage = image.embedImage;
pub const embedImageBuffer = image.embedImageBuffer;

// conditional functions
const conditional = @import("worksheet/conditional.zig");
pub const conditionalFormatRange = conditional.formatRange;
pub const conditionalFormatCell = conditional.formatCell;

// validation functions
const validation = @import("worksheet/validation.zig");
pub const dataValidationCell = validation.dataValidationCell;
pub const dataValidationRange = validation.dataValidationRange;

// comment functions
const comment = @import("worksheet/comment.zig");
pub const writeComment = comment.writeComment;
pub const showComments = comment.showComments;

const std = @import("std");
const c = @import("lxw");
const xlsxwriter = @import("xlsxwriter.zig");
const CStringArray = xlsxwriter.CStringArray;
const Bool = xlsxwriter.Boolean;
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;
const utility = @import("utility.zig");
const Cell = utility.Cell;
const Rows = utility.Rows;
const Cols = utility.Cols;
const Range = utility.Range;
const Format = @import("Format.zig");
const Chart = @import("Chart.zig");
const ChartOptions = Chart.Options;
