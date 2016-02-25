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
import openpyxl  # sudo pip install openpyxl
import os


def set_cell(
    worksheet,
    row,
    col,
    value,
    font_bold=False,
    font_size=11,
    alignment_horizontal="general",
    alignment_vertical="bottom"
):
    """Set cell value and style."""
    cell = worksheet.cell(row=row, column=col)
    cell.value = value
    cell.font = openpyxl.styles.Font(bold=font_bold, size=font_size)
    cell.alignment = openpyxl.styles.Alignment(
        horizontal=alignment_horizontal,
        vertical=alignment_vertical
    )


def generate_xlsx(args, inventory):
    """Generate the XLS file."""

    # Construct a document.
    workbook = openpyxl.workbook.Workbook()

    # Create Inventory worksheet.
    worksheet = workbook.worksheets[0]
    worksheet.title = "Inventory"

    # Create header row.
    set_cell(worksheet, 1, 1, "SKU", font_bold=True)
    worksheet.column_dimensions["A"].width = 14
    set_cell(
        worksheet,
        1,
        2,
        "Level",
        font_bold=True,
        alignment_horizontal="right"
    )
    worksheet.column_dimensions["B"].width = 6
    set_cell(worksheet, 1, 3, "Product Name", font_bold=True)
    worksheet.column_dimensions["C"].width = 50
    set_cell(worksheet, 1, 4, "Enabled", font_bold=True)
    worksheet.column_dimensions["D"].width = 8
    set_cell(worksheet, 1, 5, "Main Photo", font_bold=True)
    worksheet.column_dimensions["E"].width = 11

    # Create data rows.
    for itemid, (sku, level, name, enabled, main_photo) in enumerate(
        inventory
    ):
        row = itemid + 2
        set_cell(worksheet, row, 1, sku, alignment_horizontal="left")
        set_cell(worksheet, row, 2, int(level))
        set_cell(worksheet, row, 3, name)
        set_cell(worksheet, row, 4, enabled)
        set_cell(worksheet, row, 5, main_photo)

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
        key = cc_browser.product_key_by_sku
    else:
        key = cc_browser.product_key_by_cat_and_name
    products = sorted(products, key=key)

    # Get list of variants.
    variants = cc_browser.get_variants()
    variants = sorted(variants, key=cc_browser.variant_key)

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
            for variant in variants:
                if product_sku == variant["Product SKU"]:
                    pers_sku = variant["SKU"]
                    if pers_sku == "":
                        sku = product_sku
                    else:
                        sku = "{}-{}".format(product_sku, pers_sku)
                    pers_level = variant["Inventory Level"]
                    answer = variant["Question|Answer"]
                    answer = answer.replace("|", "=")
                    name = "{} ({})".format(product_name, answer)
                    enabled = variant["Answer Enabled"]
                    main_photo = variant["Main Photo"]
                    inventory.append(
                        (sku, pers_level, name, enabled, main_photo)
                    )

    # for sku, level, name in inventory:
    #     print("{:9} {:4} {}".format(sku, level, name))

    logger.debug("Generating {}".format(args.xlsx_filename))
    generate_xlsx(args, inventory)

    logger.debug("Generation complete")
    return 0

if __name__ == "__main__":
    main()
