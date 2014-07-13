#!/usr/bin/env python2

"""
Generate an inventory report.
"""

import ConfigParser
import argparse
import cctools
import datetime
import logging
import notify_send_handler
import openpyxl  # sudo apt-get install python-openpyxl
import os


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


def generate_xlsx(args, config, cc_browser, inventory):
    """Generate the XLS file."""

    # Construct a document.
    workbook = openpyxl.workbook.Workbook()

    # Create Inventory worksheet.
    worksheet = workbook.worksheets[0]
    worksheet.title = "Inventory"

    # Create header row.
    set_cell(worksheet, 0, 0, "SKU", bold=True)
    worksheet.column_dimensions["A"].width = 14
    set_cell(
        worksheet,
        0,
        1,
        "Level",
        bold=True,
        alignment_horizontal=openpyxl.style.Alignment.HORIZONTAL_RIGHT
    )
    worksheet.column_dimensions["B"].width = 6
    set_cell(worksheet, 0, 2, "Product Name", bold=True)
    worksheet.column_dimensions["C"].width = 50
    set_cell(worksheet, 0, 3, "Enabled", bold=True)
    worksheet.column_dimensions["D"].width = 8
    set_cell(worksheet, 0, 4, "Main Photo", bold=True)
    worksheet.column_dimensions["E"].width = 11

    # Create data rows.
    for itemid, (sku, level, name, enabled, main_photo) in enumerate(
        inventory
    ):
        row = itemid + 1
        style = set_cell(worksheet, row, 0, sku).style
        style.alignment.horizontal =\
            openpyxl.style.Alignment.HORIZONTAL_LEFT
        style.number_format.format_code =\
            openpyxl.style.NumberFormat.FORMAT_TEXT
        set_cell(worksheet, row, 1, level)
        set_cell(worksheet, row, 2, name)
        set_cell(worksheet, row, 3, enabled)
        set_cell(worksheet, row, 4, main_photo)

    # Write to file.
    workbook.save(args.xlsx_filename)


def main():
    """main"""
    default_config = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "cctools.cfg"
    )
    now = datetime.datetime.now()
    default_xlsx_filename = now.strftime("%Y-%m-%d-OnlineInventory.xlsx")

    arg_parser = argparse.ArgumentParser(
        description="Generates an inventory report."
    )
    arg_parser.add_argument(
        "--config",
        metavar="FILE",
        default=default_config,
        help="configuration filename (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--outfile",
        dest="xlsx_filename",
        metavar="FILE",
        default=default_xlsx_filename,
        help="output XLSX filename (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--sort",
        dest="sort",
        choices=["SKU", "CAT/PROD"],
        default="SKU",
        help="sort order (default=%(default)s)"
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

    # Get list of products.
    products = cc_browser.get_products()

    # Sort products.
    if args.sort == "SKU":
        key = cc_browser.sort_key_by_sku
    else:
        key = cc_browser.sort_key_by_category_and_name
    products = sorted(products, key=key)

    # Get list of personalizations.
    personalizations = cc_browser.get_personalizations()
    personalizations = sorted(
        personalizations,
        key=cc_browser.personalization_sort_key
    )

    inventory = list()
    for product in products:
        if product["Available"] == "N":
            continue
        product_sku = product["SKU"]
        product_name = product["Product Name"]
        product_level = product["Inventory Level"]
        if product["Track Inventory"] == "By Product":
            enabled = product["Available"]
            main_photo = product["Main Photo (Image)"]
            inventory.append(
                (product_sku, product_level, product_name, enabled, main_photo)
            )
        else:
            for personalization in personalizations:
                if product_sku == personalization["Product SKU"]:
                    pers_sku = personalization["SKU"]
                    if pers_sku == "":
                        sku = product_sku
                    else:
                        sku = "{}-{}".format(product_sku, pers_sku)
                    pers_level = personalization["Inventory Level"]
                    answer = personalization["Question|Answer"]
                    answer = answer.replace("|", "=")
                    name = "{} ({})".format(product_name, answer)
                    enabled = personalization["Answer Enabled"]
                    main_photo = personalization["Main Photo"]
                    inventory.append(
                        (sku, pers_level, name, enabled, main_photo)
                    )

    # for sku, level, name in inventory:
    #     print("{:9} {:4} {}".format(sku, level, name))

    logger.debug("Generating {}".format(args.xlsx_filename))
    generate_xlsx(args, config, cc_browser, inventory)

    logger.debug("Generation complete")
    return 0

if __name__ == "__main__":
    main()
