#!/usr/bin/env python2

"""
Generates a wholesale line sheet in spreadsheet form.
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


def get_optional_option(config, section, option):
    """Return an option from a ConfigParser, or None if it doesn't exist."""
    if config.has_option(section, option):
        return config.get(section, option)
    return None


def add_title(args, config, worksheet):
    """Add worksheet title."""
    row = 0
    doc_title = config.get("wholesale_line_sheet", "title")
    style = set_cell(worksheet, row, 0, doc_title, bold=True).style
    style.font.size = 20
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

    email = get_optional_option(config, "wholesale_line_sheet", "email")
    if email:
        cell_text = "Email: {}".format(email)
        set_cell(worksheet, row, 0, cell_text)
        row += 1

    website = get_optional_option(config, "wholesale_line_sheet", "website")
    if website:
        cell_text = "Website: {}".format(website)
        set_cell(worksheet, row, 0, cell_text)
        row += 1

    address = get_optional_option(config, "wholesale_line_sheet", "address")
    if address:
        cell_text = "Address: {}".format(address)
        set_cell(worksheet, row, 0, cell_text)
        row += 1

    policy = get_optional_option(config, "wholesale_line_sheet", "policy")
    if policy:
        cell_text = "Policy: {}".format(policy)
        set_cell(worksheet, row, 0, cell_text)
        row += 1

    if has_merge_cells(worksheet):
        for merge_row in range(row):
            worksheet.merge_cells(
                start_row=merge_row,
                start_column=0,
                end_row=merge_row,
                end_column=1
            )

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
    col_srp = 3
    col_size = 4
    col_sku = 5

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
        col_srp,
        "SRP",
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
    set_cell(
        worksheet,
        row,
        col_sku,
        "SKU",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
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
    for _, product_group in itertools.groupby(
        products,
        key=cc_browser.sort_key_by_category
    ):
        # Add product rows.
        for product in product_group:
            if product["Discontinued Item"] == "Y":
                continue

            set_cell(worksheet, row, col_category, product["Category"])

            description = "{}: {}".format(
                product["Product Name"],
                cctools.html_to_plain_text(product["Teaser"])
            )
            set_cell(worksheet, row, col_description, description)

            online_price = product["Price"]
            rounded_price = math.floor(float(online_price) + 0.5)
            wholesale_price = rounded_price * args.wholesale_fraction
            style = set_cell(worksheet, row, col_price, wholesale_price).style
            style.number_format.format_code = NUMBER_FORMAT_USD

            style = set_cell(worksheet, row, col_srp, online_price).style
            style.number_format.format_code = NUMBER_FORMAT_USD

            set_cell(worksheet, row, col_size, product["Size"])

            set_cell(worksheet, row, col_sku, product["SKU"])

            row += 1

    # Set column widths.
    worksheet.column_dimensions[col_letter(col_category)].width = 21
    worksheet.column_dimensions[col_letter(col_description)].width = 70
    worksheet.column_dimensions[col_letter(col_price)].width = 7
    worksheet.column_dimensions[col_letter(col_srp)].width = 7
    worksheet.column_dimensions[col_letter(col_size)].width = 24
    worksheet.column_dimensions[col_letter(col_sku)].width = 6


def add_line_sheet(args, config, cc_browser, products, worksheet):
    """Create the Wholesale Line Sheet worksheet."""

    # Prepare worksheet.
    worksheet.title = "Wholesale Line Sheet"

    # Add title.
    row = add_title(args, config, worksheet)

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

    # Create Line Sheet worksheet.
    add_line_sheet(
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
    default_xlsx_filename = now.strftime("%Y-%m-%d-WholesaleLineSheet.xlsx")

    arg_parser = argparse.ArgumentParser(
        description="Generates a wholesale line sheet."
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
