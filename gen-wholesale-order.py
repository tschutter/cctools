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
import openpyxl  # sudo pip install openpyxl
import os

NUMBER_FORMAT_USD = "$#,##0.00;-$#,##0.00"


def set_cell(
    worksheet,
    row,
    col,
    value,
    font_bold=False,
    font_size=11,
    alignment_horizontal="general",
    alignment_vertical="bottom",
    number_format="General"
):
    """Set cell value and style."""
    cell = worksheet.cell(row=row, column=col)
    cell.value = value
    cell.font = openpyxl.styles.Font(bold=font_bold, size=font_size)
    cell.alignment = openpyxl.styles.Alignment(
        horizontal=alignment_horizontal,
        vertical=alignment_vertical
    )
    if number_format != "General":
        cell.number_format = number_format


def col_letter(col):
    """Return column letter for given column."""
    return chr(ord("A") + col - 1)


def add_title(args, config, worksheet):
    """Add worksheet title."""
    row = 1
    doc_title = config.get("wholesale_order", "title")
    set_cell(worksheet, row, 1, doc_title, font_bold=True)
    worksheet.merge_cells(
        start_row=0,
        start_column=1,
        end_row=0,
        end_column=2
    )
    worksheet.row_dimensions[1].height = 25
    row += 1

    now = datetime.datetime.now()
    cell_text = now.strftime("Date: %Y-%m-%d")
    set_cell(worksheet, row, 1, cell_text)
    row += 1

    valid_date = now + datetime.timedelta(days=args.valid_ndays)
    cell_text = valid_date.strftime("Valid until: %Y-%m-%d")
    set_cell(worksheet, row, 1, cell_text)
    row += 1

    cell_text = "Prices are {:.0%} of retail".format(args.wholesale_fraction)
    set_cell(worksheet, row, 1, cell_text)
    row += 1

    return row


def add_ship_to(worksheet, row):
    """Add Ship To block."""
    start_col = 1
    end_col = 2
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
        font_bold=True,
        alignment_horizontal="left"
    )
    row += 1

    nrows = 3
    for i in range(0, nrows):
        cell = worksheet.cell(row=row, column=start_col)
        cell.alignment = openpyxl.styles.Alignment(horizontal="left")
        worksheet.merge_cells(
            start_row=row,
            start_column=start_col,
            end_row=row,
            end_column=end_col
        )
        for col in range(start_col, end_col + 1):
            default_side = openpyxl.styles.Side(
                border_style=None,
                color='FF000000'
            )
            if i == 0:
                top = openpyxl.styles.Side("thin")  # BORDER_THIN
            else:
                top = default_side
            if i == nrows - 1:
                bottom = openpyxl.styles.Side("thin")
            else:
                bottom = default_side
            if col == start_col:
                left = openpyxl.styles.Side("thin")
            else:
                left = default_side
            if col == end_col:
                right = openpyxl.styles.Side("thin")
            else:
                right = default_side
            cell = worksheet.cell(row=row, column=col)
            cell.border = openpyxl.styles.Border(
                left=left,
                right=right,
                top=top,
                bottom=bottom
            )
        row += 1

    return row


def set_label_dollar_value(
    worksheet,
    row,
    col_label_start,
    col_label_end,
    col_total,
    label,
    value,
    font_bold=False
):
    """Add label: value."""
    worksheet.merge_cells(
        start_row=row,
        start_column=col_label_start,
        end_row=row,
        end_column=col_label_end
    )
    set_cell(
        worksheet,
        row,
        col_label_start,
        label,
        font_bold=font_bold,
        alignment_horizontal="right"
    )

    set_cell(
        worksheet,
        row,
        col_total,
        value,
        font_bold=font_bold,
        number_format=NUMBER_FORMAT_USD
    )


def add_products(args, worksheet, row, cc_browser, products):
    """Add row for each product."""
    col_category = 1
    col_description = 2
    col_price = 3
    col_qty = 4
    col_total = 5
    col_sku = 6
    col_size = 7

    # Add header row.
    set_cell(
        worksheet,
        row,
        col_category,
        "Category",
        font_bold=True
    )
    set_cell(worksheet, row, col_description, "Description", font_bold=True)
    set_cell(
        worksheet,
        row,
        col_price,
        "Price",
        font_bold=True,
        alignment_horizontal="right"
    )
    set_cell(
        worksheet,
        row,
        col_qty,
        "Qty",
        font_bold=True,
        alignment_horizontal="right"
    )
    set_cell(
        worksheet,
        row,
        col_total,
        "Total",
        font_bold=True,
        alignment_horizontal="right"
    )
    set_cell(
        worksheet,
        row,
        col_sku,
        "SKU",
        font_bold=True,
        alignment_horizontal="right"
    )
    set_cell(
        worksheet,
        row,
        col_size,
        "Size",
        font_bold=True
    )
    row += 1

    # Remove excluded SKUs.
    if args.exclude_skus:
        products = [
            x for x in products if str(x["SKU"]) not in args.exclude_skus
        ]

    # Sort products by category, product_name.
    products = sorted(products, key=cc_browser.product_key_by_cat_and_name)

    # Group products by category.
    first_product_row = row
    for _, product_group in itertools.groupby(
        products,
        key=cc_browser.product_key_by_category
    ):
        # Add product rows.
        for product in product_group:
            if product["Available"] != "Y":
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
            set_cell(
                worksheet,
                row,
                col_price,
                wholesale_price,
                number_format=NUMBER_FORMAT_USD
            )
            total_formula = "=IF({}{}=\"\", \"\", {}{} * {}{})".format(
                col_letter(col_qty),
                row,
                col_letter(col_price),
                row,
                col_letter(col_qty),
                row
            )
            set_cell(
                worksheet,
                row,
                col_total,
                total_formula,
                number_format=NUMBER_FORMAT_USD
            )
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
        first_product_row,
        col_letter(col_total),
        last_product_row
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
        subtotal_row,
        col_letter(col_total),
        row - 1
    )
    set_label_dollar_value(
        worksheet,
        row,
        col_label_start,
        col_label_end,
        col_total,
        "Total:",
        total_formula,
        font_bold=True
    )


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
