const lxw = @import("lxw");

pub fn main() void {

    // Create a new workbook and add a worksheet.
    const workbook: ?*lxw.lxw_workbook = lxw.workbook_new("zig-array_formula.xlsx");
    const worksheet: ?*lxw.lxw_worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Write some data for the formulas.
    _ = lxw.worksheet_write_number(worksheet, 0, 1, 500, null);
    _ = lxw.worksheet_write_number(worksheet, 1, 1, 10, null);
    _ = lxw.worksheet_write_number(worksheet, 4, 1, 1, null);
    _ = lxw.worksheet_write_number(worksheet, 5, 1, 2, null);
    _ = lxw.worksheet_write_number(worksheet, 6, 1, 3, null);

    _ = lxw.worksheet_write_number(worksheet, 0, 2, 300, null);
    _ = lxw.worksheet_write_number(worksheet, 1, 2, 15, null);
    _ = lxw.worksheet_write_number(worksheet, 4, 2, 20234, null);
    _ = lxw.worksheet_write_number(worksheet, 5, 2, 21003, null);
    _ = lxw.worksheet_write_number(worksheet, 6, 2, 10000, null);

    // Write an array formula that returns a single value.
    _ = lxw.worksheet_write_array_formula(worksheet, 0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}", null);

    // Similar to above but using the RANGE macro.
    _ = lxw.worksheet_write_array_formula(worksheet, 1, lxw.RANGE("A2:A2"), 0, 0, "{=SUM(B1:C1*B2:C2)}", null);

    // Write an array formula that returns a range of values.
    _ = lxw.worksheet_write_array_formula(worksheet, 4, 0, 6, 0, "{=TREND(C5:C7,B5:B7)}", null);

    _ = lxw.workbook_close(workbook);
}
