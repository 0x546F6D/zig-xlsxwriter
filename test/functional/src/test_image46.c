/*****************************************************************************
 * Test cases for libxlsxwriter.
 *
 * Test to compare output against Excel files.
 *
 * SPDX-License-Identifier: BSD-2-Clause
 * Copyright 2014-2025, John McNamara, jmcnamara@cpan.org.
 *
 */

#include "xlsxwriter.h"

int main() {

    lxw_workbook  *workbook  = workbook_new("test_image46.xlsx");
    lxw_worksheet *worksheet = workbook_add_worksheet(workbook, NULL);

    lxw_image_options image_options = {.y_offset = 4, .object_position = LXW_OBJECT_MOVE_AND_SIZE_AFTER};
    worksheet_insert_image_opt(worksheet, CELL("E9"), "images/red.png", &image_options);

    lxw_row_col_options row_options = {.hidden = LXW_TRUE};
    worksheet_set_row_opt(worksheet, 8, 30, NULL, &row_options);

    return workbook_close(workbook);
}
