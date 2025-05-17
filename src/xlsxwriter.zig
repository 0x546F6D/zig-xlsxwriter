pub const c = @import("xlsxwriter_c");

pub const XlsxError = @import("errors.zig").XlsxError;
pub const checkResult = @import("errors.zig").checkResult;
pub const formatErr = @import("errors.zig").formatErr;

pub const WorkBook = @import("WorkBook.zig");
pub const WorkSheet = @import("WorkSheet.zig");

pub const Format = @import("Format.zig");
pub const DefinedColor = Format.DefinedColor;

pub const Chart = @import("Chart.zig");
pub const ChartType = Chart.Type;
pub const ChartFont = Chart.Font;
pub const ChartSeries = Chart.Series;

pub const explicit_false = c.LXW_EXPLICIT_FALSE;
pub const datetime = c.lxw_datetime;
pub const range = c.RANGE;
pub const cell = c.CELL;
pub const cols = c.COLS;
