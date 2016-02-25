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


def get_optional_option(config, section, option):
    """Return an option from a ConfigParser, or None if it doesn't exist."""
    if config.has_option(section, option):
        return config.get(section, option)
    return None


def add_title(args, config, worksheet):
    """Add worksheet title."""
    row = 1

    doc_title = config.get("wholesale_line_sheet", "title")
    set_cell(worksheet, row, 1, doc_title, font_bold=True, font_size=20)
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

    email = get_optional_option(config, "wholesale_line_sheet", "email")
    if email:
        cell_text = "Email: {}".format(email)
        set_cell(worksheet, row, 1, cell_text)
        row += 1

    phone = get_optional_option(config, "wholesale_line_sheet", "phone")
    if phone:
        cell_text = "Phone: {}".format(phone)
        set_cell(worksheet, row, 1, cell_text)
        row += 1

    website = get_optional_option(config, "wholesale_line_sheet", "website")
    if website:
        cell_text = "Website: {}".format(website)
        set_cell(worksheet, row, 1, cell_text)
        row += 1

    address = get_optional_option(config, "wholesale_line_sheet", "address")
    if address:
        cell_text = "Address: {}".format(address)
        set_cell(worksheet, row, 1, cell_text)
        row += 1

    policy = get_optional_option(config, "wholesale_line_sheet", "policy")
    if policy:
        cell_text = "Policy: {}".format(policy)
        set_cell(worksheet, row, 1, cell_text)
        row += 1

    for merge_row in range(row):
        worksheet.merge_cells(
            start_row=merge_row,
            start_column=1,
            end_row=merge_row,
            end_column=2
        )

    return row


def add_products(args, worksheet, row, cc_browser, products):
    """Add row for each product."""
    col_category = 1
    col_description = 2
    col_price = 3
    col_msrp = 4
    col_size = 5
    col_sku = 6

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
        col_msrp,
        "MSRP",
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
    set_cell(
        worksheet,
        row,
        col_sku,
        "SKU",
        font_bold=True,
        alignment_horizontal="right"
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
    for _, product_group in itertools.groupby(
        products,
        key=cc_browser.product_key_by_category
    ):
        # Add product rows.
        for product in product_group:
            if product["Available"] != "Y":
                continue

            set_cell(worksheet, row, col_category, product["Category"])

            description = "{}: {}".format(
                product["Product Name"],
                cctools.html_to_plain_text(product["Teaser"])
            )
            set_cell(worksheet, row, col_description, description)

            online_price = float(product["Price"])
            if online_price > 1.0:
                rounded_price = math.floor(online_price + 0.5)
            else:
                rounded_price = online_price
            wholesale_price = rounded_price * args.wholesale_fraction

            set_cell(
                worksheet,
                row,
                col_price,
                wholesale_price,
                number_format=NUMBER_FORMAT_USD
            )

            set_cell(
                worksheet,
                row,
                col_msrp,
                online_price,
                number_format=NUMBER_FORMAT_USD
            )

            set_cell(worksheet, row, col_size, product["Size"])

            set_cell(worksheet, row, col_sku, product["SKU"])

            row += 1

    # Set column widths.
    worksheet.column_dimensions[col_letter(col_category)].width = 21
    worksheet.column_dimensions[col_letter(col_description)].width = 70
    worksheet.column_dimensions[col_letter(col_price)].width = 9
    worksheet.column_dimensions[col_letter(col_msrp)].width = 9
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
