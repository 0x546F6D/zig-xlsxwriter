//
// Example of how to create defined names using libxlsxwriter. This method is
// used to define a user friendly name to represent a value, a single cell or
// a range of cells in a workbook.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-defined_name.xlsx");

    // Add two worksheets
    const worksheet1 = lxw.workbook_add_worksheet(workbook, null);
    const worksheet2 = lxw.workbook_add_worksheet(workbook, null);

    // Define some global/workbook names
    _ = lxw.workbook_define_name(workbook, "Exchange_rate", "=0.96");
    _ = lxw.workbook_define_name(workbook, "Sales", "=Sheet1!$G$1:$H$10");

    // Define a local/worksheet name. This overrides the global "Sales" name
    // with a local defined name.
    _ = lxw.workbook_define_name(workbook, "Sheet2!Sales", "=Sheet2!$G$1:$G$10");

    // Write some text to the worksheets and one of the defined names in a formula
    // Process worksheet1
    _ = lxw.worksheet_set_column(
        worksheet1,
        0,
        0,
        45,
        null,
    );

    _ = lxw.worksheet_write_string(
        worksheet1,
        0,
        0,
        "This worksheet contains some defined names.",
        null,
    );

    _ = lxw.worksheet_write_string(
        worksheet1,
        1,
        0,
        "See Formulas -> Name Manager above.",
        null,
    );

    _ = lxw.worksheet_write_string(
        worksheet1,
        2,
        0,
        "Example formula in cell B3 ->",
        null,
    );

    _ = lxw.worksheet_write_formula(
        worksheet1,
        2,
        1,
        "=Exchange_rate",
        null,
    );

    // Process worksheet2
    _ = lxw.worksheet_set_column(
        worksheet2,
        0,
        0,
        45,
        null,
    );

    _ = lxw.worksheet_write_string(
        worksheet2,
        0,
        0,
        "This worksheet contains some defined names.",
        null,
    );

    _ = lxw.worksheet_write_string(
        worksheet2,
        1,
        0,
        "See Formulas -> Name Manager above.",
        null,
    );

    _ = lxw.worksheet_write_string(
        worksheet2,
        2,
        0,
        "Example formula in cell B3 ->",
        null,
    );

    _ = lxw.worksheet_write_formula(
        worksheet2,
        2,
        1,
        "=Exchange_rate",
        null,
    );

    _ = lxw.workbook_close(workbook);
}
