#!/usr/bin/env python2

"""
Generates a FedEx Product Dictionary.
"""

import ConfigParser
import argparse
import cctools
import itertools
import openpyxl  # sudo apt-get install python-openpyxl
import os
import sys
import datetime


def add_product_dict(args, config, cc_browser, products, worksheet):
    """Create the Product Dictionary worksheet."""

    # Prepare worksheet.
    worksheet.title = "Product Dictionary"

    # Add products.
    # Remove excluded SKUs.
    if args.exclude_skus:
        products = [
            x for x in products if str(x["SKU"]) not in args.exclude_skus
        ]

    # Add product rows, grouped by category.
    row = 0
    for _, product_group in itertools.groupby(
        products,
        key=cc_browser.sort_key_by_category
    ):
        # Add product rows.
        for product in product_group:
            if product["Discontinued Item"] == "Y":
                continue
            description = "%s: %s" % (
                product["Product Name"],
                cctools.html_to_plain_text(product["Teaser"])
            )
            worksheet.cell(row=row, column=0).value = product["SKU"]
            worksheet.cell(row=row, column=1).value = description
            worksheet.cell(row=row, column=2).value = product["HTSUS No"]
            row += 1

    # Set column widths.
    worksheet.column_dimensions["A"].width = 6
    worksheet.column_dimensions["B"].width = 95
    worksheet.column_dimensions["C"].width = 13


def generate_xlsx(args, config, cc_browser, products):
    """Generate the XLS file."""

    # Construct a document.
    workbook = openpyxl.workbook.Workbook()

    # Create Product Dictionary worksheet.
    add_product_dict(
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
    defaultConfig = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "cctools.cfg"
    )
    now = datetime.datetime.now()
    default_xlsx_filename = now.strftime("%Y-%m-%d-PurchaseOrder.xlsx")

    arg_parser = argparse.ArgumentParser(
        description="Generates a Purchase Order / Commercial Invoice."
    )
    arg_parser.add_argument(
        "--config",
        action="store",
        dest="config",
        metavar="FILE",
        default=defaultConfig,
        help="configuration filename (default=%(default)s)"
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

    # Read config file.
    config = ConfigParser.RawConfigParser()
    config.readfp(open(args.config))

    # Create a connection to CoreCommerce.
    cc_browser = cctools.CCBrowser(
        config.get("website", "host"),
        config.get("website", "site"),
        config.get("website", "username"),
        config.get("website", "password"),
        verbose=args.verbose
    )

    # Fetch products list.
    products = cc_browser.get_products()

    # Generate spreadsheet.
    if args.verbose:
        sys.stderr.write(
            "Generating %s\n" % os.path.abspath(args.xlsx_filename)
        )
    generate_xlsx(args, config, cc_browser, products)

    if args.verbose:
        sys.stderr.write("Generation complete\n")
    return 0


if __name__ == "__main__":
    main()
