#!/usr/bin/env python

"""
Generate an inventory report.
"""

import ConfigParser
import cctools
import datetime
import openpyxl  # sudo apt-get install python-openpyxl
import optparse
import os
import sys

def set_cell(
    worksheet,
    row,
    col,
    value,
    bold = None,
    alignment_horizontal = None,
    alignment_vertical = None
):
    """Set cell value and style."""
    cell = worksheet.cell(row=row, column=col)
    cell.value = value
    if bold != None:
        cell.style.font.bold = bold
    if alignment_horizontal != None:
        cell.style.alignment.horizontal = alignment_horizontal
    if alignment_vertical != None:
        cell.style.alignment.vertical = alignment_vertical
    return cell


def generate_xlsx(options, config, cc_browser, inventory):
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

    # Create data rows.
    for itemid, (sku, level, name, enabled) in enumerate(inventory):
        row = itemid + 1
        style = set_cell(worksheet, row, 0, sku).style
        style.alignment.horizontal =\
            openpyxl.style.Alignment.HORIZONTAL_LEFT
        style.number_format.format_code =\
            openpyxl.style.NumberFormat.FORMAT_TEXT
        set_cell(worksheet, row, 1, level)
        set_cell(worksheet, row, 2, name)
        set_cell(worksheet, row, 3, enabled)

    # Write to file.
    workbook.save(options.xlsx_filename)


def main():
    """main"""
    default_config = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "cctools.cfg"
    )
    now = datetime.datetime.now()
    default_xlsx_filename = now.strftime("%Y-%m-%d-OnlineInventory.xlsx")

    option_parser = optparse.OptionParser(
        usage="usage: %prog [options] action\n" +
        "  Actions:\n" +
        "    products - list products\n"
        "    categories - list categories\n"
        "    personalizations - list personalizations"
    )
    option_parser.add_option(
        "--config",
        action="store",
        metavar="FILE",
        default=default_config,
        help="configuration filename (default=%default)"
    )
    option_parser.add_option(
        "--outfile",
        action="store",
        dest="xlsx_filename",
        metavar="FILE",
        default=default_xlsx_filename,
        help="output XLSX filename (default=%default)"
    )
    option_parser.add_option(
        "--verbose",
        action="store_true",
        default=False,
        help="display progress messages"
    )

    # Parse command line arguments.
    (options, args) = option_parser.parse_args()
    if len(args) != 0:
        option_parser.error("invalid argument")

    # Read config file.
    config = ConfigParser.RawConfigParser()
    config.readfp(open(options.config))

    # Create a connection to CoreCommerce.
    cc_browser = cctools.CCBrowser(
        config.get("website", "host"),
        config.get("website", "site"),
        config.get("website", "username"),
        config.get("website", "password"),
        verbose=options.verbose
    )

    # Get list of products by category, product_name.
    products = cc_browser.get_products()
    products = sorted(products, key=cc_browser.sort_key_by_category_and_name)

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
            inventory.append(
                (product_sku, product_level, product_name, enabled)
            )
        else:
            for personalization in personalizations:
                if product_sku == personalization["Product SKU"]:
                    pers_sku = personalization["SKU"]
                    if pers_sku == "":
                        sku = product_sku
                    else:
                        sku = "%s-%s" % (product_sku, pers_sku)
                    pers_level = personalization["Inventory Level"]
                    answer = personalization["Question|Answer"]
                    answer = answer.replace("|", "=")
                    name = "%s (%s)" % (product_name, answer)
                    enabled = personalization["Answer Enabled"]
                    inventory.append((sku, pers_level, name, enabled))

    #for sku, level, name in inventory:
    #    print "%-9s %4s %-45s" % (sku, level, name)

    if options.verbose:
        sys.stderr.write("Generating %s\n" % options.xlsx_filename)
    generate_xlsx(options, config, cc_browser, inventory)

    return 0

if __name__ == "__main__":
    main()
