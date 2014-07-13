#!/usr/bin/env python2

"""
Generates a wholesale order form in spreadsheet form.
"""

import ConfigParser
import argparse
import cctools
import datetime
import itertools
import logging
import math
import notify_send_handler
import openpyxl  # sudo apt-get install python-openpyxl
import os

# Cell style constants.
ALIGNMENT_HORIZONTAL_LEFT = openpyxl.style.Alignment.HORIZONTAL_LEFT
ALIGNMENT_HORIZONTAL_RIGHT = openpyxl.style.Alignment.HORIZONTAL_RIGHT
ALIGNMENT_VERTICAL_TOP = openpyxl.style.Alignment.VERTICAL_TOP
NUMBER_FORMAT_USD = openpyxl.style.NumberFormat.FORMAT_CURRENCY_USD_SIMPLE


def has_merge_cells(worksheet):
    """Determine if Worksheet.merge_cells method exists."""
    # merge_cells not supported by openpyxl-1.5.6 (Ubuntu 12.04)
    return hasattr(worksheet, "merge_cells")


def set_cell(
    worksheet,
    row,
    col,
    value,
    bold=None,
    alignment_horizontal=None,
    alignment_vertical=None
):
    """Set cell value and style."""
    cell = worksheet.cell(row=row, column=col)
    cell.value = value
    if bold is not None:
        cell.style.font.bold = bold
    if alignment_horizontal is not None:
        cell.style.alignment.horizontal = alignment_horizontal
    if alignment_vertical is not None:
        cell.style.alignment.vertical = alignment_vertical
    return cell


def col_letter(col):
    """Return column letter for given column."""
    return chr(ord("A") + col)


def row_number(row):
    """Return row number for given row."""
    return row + 1


def add_title(args, config, worksheet):
    """Add worksheet title."""
    row = 0
    doc_title = config.get("wholesale_order", "title")
    style = set_cell(worksheet, row, 0, doc_title, bold=True).style
    style.font.size = 20
    if has_merge_cells(worksheet):
        worksheet.merge_cells(
            start_row=0,
            start_column=0,
            end_row=0,
            end_column=1
        )
    worksheet.row_dimensions[1].height = 25
    row += 1

    now = datetime.datetime.now()
    cell_text = now.strftime("Date: %Y-%m-%d")
    set_cell(worksheet, row, 0, cell_text)
    row += 1

    valid_date = now + datetime.timedelta(days=args.valid_ndays)
    cell_text = valid_date.strftime("Valid until: %Y-%m-%d")
    set_cell(worksheet, row, 0, cell_text)
    row += 1

    cell_text = "Prices are {:.0%} of retail".format(args.wholesale_fraction)
    set_cell(worksheet, row, 0, cell_text)
    row += 1

    return row


def add_ship_to(worksheet, row):
    """Add Ship To block."""
    start_col = 0
    end_col = 1
    if has_merge_cells(worksheet):
        worksheet.merge_cells(
            start_row=row,
            start_column=start_col,
            end_row=row,
            end_column=end_col
        )
    set_cell(
        worksheet,
        row,
        start_col,
        "Ship To:",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_LEFT
    )
    row += 1

    nrows = 3
    for i in range(0, nrows):
        style = worksheet.cell(row=row, column=start_col).style
        style.alignment.horizontal = ALIGNMENT_HORIZONTAL_LEFT
        if has_merge_cells(worksheet):
            worksheet.merge_cells(
                start_row=row,
                start_column=start_col,
                end_row=row,
                end_column=end_col
            )
        for col in range(start_col, end_col + 1):
            style = worksheet.cell(row=row, column=col).style
            borders = style.borders
            if i == 0:
                borders.top.border_style = openpyxl.style.Border.BORDER_THIN
            if i == nrows - 1:
                borders.bottom.border_style = openpyxl.style.Border.BORDER_THIN
            if col == start_col:
                borders.left.border_style = openpyxl.style.Border.BORDER_THIN
            if col == end_col:
                borders.right.border_style = openpyxl.style.Border.BORDER_THIN
        row += 1

    return row


def set_label_dollar_value(
    worksheet,
    row,
    col_label_start,
    col_label_end,
    col_total,
    label,
    value
):
    """Add label: value."""
    if has_merge_cells(worksheet):
        worksheet.merge_cells(
            start_row=row,
            start_column=col_label_start,
            end_row=row,
            end_column=col_label_end
        )
        col_label = col_label_start
    else:
        col_label = col_label_end
    label_style = set_cell(worksheet, row, col_label, label).style
    label_style.alignment.horizontal = ALIGNMENT_HORIZONTAL_RIGHT

    value_style = set_cell(worksheet, row, col_total, value).style
    value_style.number_format.format_code = NUMBER_FORMAT_USD
    return(label_style, value_style)


def add_products(args, worksheet, row, cc_browser, products):
    """Add row for each product."""
    col_category = 0
    col_description = 1
    col_price = 2
    col_qty = 3
    col_total = 4
    col_sku = 5
    col_size = 6

    # Add header row.
    set_cell(
        worksheet,
        row,
        col_category,
        "Category",
        bold=True
    )
    set_cell(worksheet, row, col_description, "Description", bold=True)
    set_cell(
        worksheet,
        row,
        col_price,
        "Price",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    set_cell(
        worksheet,
        row,
        col_qty,
        "Qty",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    set_cell(
        worksheet,
        row,
        col_total,
        "Total",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    set_cell(
        worksheet,
        row,
        col_sku,
        "SKU",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    set_cell(
        worksheet,
        row,
        col_size,
        "Size",
        bold=True
    )
    row += 1

    # Remove excluded SKUs.
    if args.exclude_skus:
        products = [
            x for x in products if str(x["SKU"]) not in args.exclude_skus
        ]

    # Sort products by category, product_name.
    products = sorted(products, key=cc_browser.sort_key_by_category_and_name)

    # Group products by category.
    first_product_row = row
    for _, product_group in itertools.groupby(
        products,
        key=cc_browser.sort_key_by_category
    ):
        # Add product rows.
        for product in product_group:
            if product["Discontinued Item"] == "Y":
                continue
            description = "{}: {}".format(
                product["Product Name"],
                cctools.html_to_plain_text(product["Teaser"])
            )
            set_cell(worksheet, row, col_category, product["Category"])
            set_cell(worksheet, row, col_description, description)
            online_price = product["Price"]
            rounded_price = math.floor(float(online_price) + 0.5)
            wholesale_price = rounded_price * args.wholesale_fraction
            style = set_cell(worksheet, row, col_price, wholesale_price).style
            style.number_format.format_code = NUMBER_FORMAT_USD
            total_formula = "=IF({}{}=\"\", \"\", {}{} * {}{})".format(
                col_letter(col_qty),
                row_number(row),
                col_letter(col_price),
                row_number(row),
                col_letter(col_qty),
                row_number(row)
            )
            style = set_cell(worksheet, row, col_total, total_formula).style
            style.number_format.format_code = NUMBER_FORMAT_USD
            set_cell(worksheet, row, col_sku, product["SKU"])
            set_cell(worksheet, row, col_size, product["Size"])
            row += 1
    last_product_row = row - 1

    # Set column widths.
    worksheet.column_dimensions[col_letter(col_category)].width = 20
    worksheet.column_dimensions[col_letter(col_description)].width = 65
    worksheet.column_dimensions[col_letter(col_price)].width = 7
    worksheet.column_dimensions[col_letter(col_qty)].width = 5
    worksheet.column_dimensions[col_letter(col_total)].width = 10
    worksheet.column_dimensions[col_letter(col_sku)].width = 6
    worksheet.column_dimensions[col_letter(col_size)].width = 24

    # Blank row.
    row += 1

    col_label_start = col_total - 2
    col_label_end = col_total - 1

    # Subtotal.
    subtotal_formula = "=SUM({}{}:{}{})".format(
        col_letter(col_total),
        row_number(first_product_row),
        col_letter(col_total),
        row_number(last_product_row)
    )
    set_label_dollar_value(
        worksheet,
        row,
        col_label_start,
        col_label_end,
        col_total,
        "Subtotal:",
        subtotal_formula
    )
    subtotal_row = row
    row += 1

    # Shipping.
    set_label_dollar_value(
        worksheet,
        row,
        col_label_start,
        col_label_end,
        col_total,
        "Shipping:",
        0.0
    )
    row += 1

    # Adjustment.
    set_label_dollar_value(
        worksheet,
        row,
        col_label_start,
        col_label_end,
        col_total,
        "Adjustment:",
        0.0
    )
    row += 1

    # Total.
    total_formula = "=SUM({}{}:{}{})".format(
        col_letter(col_total),
        row_number(subtotal_row),
        col_letter(col_total),
        row_number(row - 1)
    )
    (label_style, value_style) = set_label_dollar_value(
        worksheet,
        row,
        col_label_start,
        col_label_end,
        col_total,
        "Total:",
        total_formula
    )
    label_style.font.bold = True
    value_style.font.bold = True


def add_order_form(args, config, cc_browser, products, worksheet):
    """Create the Wholesale Order Form worksheet."""

    # Prepare worksheet.
    worksheet.title = "Wholesale Order Form"

    # Add title.
    row = add_title(args, config, worksheet)

    # Blank row.
    row += 1

    # Ship To block.
    row = add_ship_to(worksheet, row)

    # Blank row.
    row += 1

    # Add products.
    add_products(
        args,
        worksheet,
        row,
        cc_browser,
        products
    )


def generate_xlsx(args, config, cc_browser, products):
    """Generate the XLS file."""

    # Construct a document.
    workbook = openpyxl.workbook.Workbook()

    # Create PO-Invoice worksheet.
    add_order_form(
        args,
        config,
        cc_browser,
        products,
        workbook.worksheets[0]
    )

    # Write to file.
    workbook.save(args.xlsx_filename)


def main():
    """main"""
    default_config = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "cctools.cfg"
    )
    now = datetime.datetime.now()
    default_xlsx_filename = now.strftime("%Y-%m-%d-WholesaleOrderForm.xlsx")

    arg_parser = argparse.ArgumentParser(
        description="Generates a wholesale order form."
    )
    arg_parser.add_argument(
        "--config",
        action="store",
        dest="config",
        metavar="FILE",
        default=default_config,
        help="configuration filename (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--wholesale-fraction",
        metavar="FRAC",
        default=0.5,
        help="wholesale price fraction (default=%(default).2f)"
    )
    arg_parser.add_argument(
        "--valid-ndays",
        metavar="N",
        type=int,
        default=30,
        help="number of days prices are valid (default=%(default)i)"
    )
    arg_parser.add_argument(
        "--outfile",
        action="store",
        dest="xlsx_filename",
        metavar="FILE",
        default=default_xlsx_filename,
        help="output XLSX filename (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--exclude-sku",
        action="append",
        dest="exclude_skus",
        metavar="SKU",
        help="exclude SKU from output"
    )
    arg_parser.add_argument(
        "--verbose",
        action="store_true",
        default=False,
        help="display progress messages"
    )

    # Parse command line arguments.
    args = arg_parser.parse_args()

    # Configure logging.
    logging.basicConfig(
        level=logging.INFO if args.verbose else logging.WARNING
    )
    logger = logging.getLogger()

    # Also log using notify-send if it is available.
    if notify_send_handler.NotifySendHandler.is_available():
        logger.addHandler(
            notify_send_handler.NotifySendHandler(
                os.path.splitext(os.path.basename(__file__))[0]
            )
        )

    # Read config file.
    config = ConfigParser.RawConfigParser()
    config.readfp(open(args.config))

    # Create a connection to CoreCommerce.
    cc_browser = cctools.CCBrowser(
        config.get("website", "host"),
        config.get("website", "site"),
        config.get("website", "username"),
        config.get("website", "password")
    )

    # Fetch products list.
    products = cc_browser.get_products()

    # Generate spreadsheet.
    logger.debug("Generating {}".format(args.xlsx_filename))
    generate_xlsx(args, config, cc_browser, products)

    logger.debug("Generation complete")
    return 0


if __name__ == "__main__":
    main()
