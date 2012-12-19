#!/usr/bin/env python

"""
Generates a Purchase Order / Commercial Invoice.
"""

# TODO:
#  fetch htsus

import ConfigParser
import csv
import itertools
import optparse
import random
import sys
import time
try:
    import openpyxl
except ImportError:
    print "ERROR: Python module 'openpyxl' not installed."
    print "       Run 'sudo apt-get install python-openpyxl' to install."
    sys.exit(1)

# Today's date in ISO format.
TODAY = time.strftime("%Y-%m-%d", time.localtime(time.time()))

# Cell style constants.
ALIGNMENT_HORIZONTAL_RIGHT = openpyxl.style.Alignment.HORIZONTAL_RIGHT
ALIGNMENT_VERTICAL_TOP = openpyxl.style.Alignment.VERTICAL_TOP


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


def col_letter(col):
    """Return letter for given column."""
    return chr(ord("A") + col)


def add_title(worksheet):
    """Add worksheet title."""
    style = set_cell(worksheet, 0, 0, "Purchase Order", bold=True).style
    style.font.size = 20
    #worksheet.merge_cells(start_row=0, start_col=0, end_row=0, end_col=2)


def add_header(worksheet, config, row):
    """Add PO/Invoice header."""
    col_value_name = 2
    col_value = col_value_name + 1

    set_cell(
        worksheet,
        row,
        col_value_name,
        "Date: ",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    set_cell(worksheet, row, col_value, TODAY)
    row += 1

    set_cell(
        worksheet,
        row,
        col_value_name,
        "Country of Origin: ",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    set_cell(
        worksheet,
        row,
        col_value,
        config.get("invoice", "country_of_origin")
    )

    row += 1

    set_cell(
        worksheet,
        row,
        col_value_name,
        "Manufacturer ID: ",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    set_cell(
        worksheet,
        row,
        col_value,
        config.get("invoice", "manufacturer_id")
    )
    row += 1

    set_cell(
        worksheet,
        row,
        col_value_name,
        "Ultimate Consignee: ",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT,
        alignment_vertical=ALIGNMENT_VERTICAL_TOP
    )
    set_cell(
        worksheet,
        row,
        col_value,
        config.get("invoice", "consignee").replace("\\n", "\n")
    )
    row += 1

    # Unit of measurement.
    set_cell(
        worksheet,
        row,
        col_value_name,
        "Unit of Measurement: ",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    set_cell(
        worksheet,
        row,
        col_value,
        config.get("invoice", "unit_of_measurement")
    )
    row += 1

    # Terms of sale.
    set_cell(
        worksheet,
        row,
        col_value_name,
        "Currency: ",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    set_cell(
        worksheet,
        row,
        col_value,
        config.get("invoice", "currency")
    )
    row += 1

    # Transport and delivery.
    set_cell(
        worksheet,
        row,
        col_value_name,
        "Transport and Delivery: ",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    set_cell(
        worksheet,
        row,
        col_value,
        config.get("invoice", "transport_and_delivery")
    )
    row += 1

    # Terms of sale.
    set_cell(
        worksheet,
        row,
        col_value_name,
        "Terms of Sale: ",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    set_cell(
        worksheet,
        row,
        col_value,
        config.get("invoice", "terms_of_sale")
    )
    row += 1

    return row


def add_products(worksheet, row, products_by_category):
    """Add row for each product."""
    col_line_no = 0
    col_sku = 1
    col_description = 2
    col_price = 3
    col_qty = 4
    col_total = 5
    col_htsus_no = 6
    col_instructions = 7

    # Add header row.
    set_cell(
        worksheet,
        row,
        col_line_no,
        "Line No",
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
        col_htsus_no,
        "HTSUS No",
        bold=True
    )
    set_cell(
        worksheet,
        row,
        col_instructions,
        "Item Special Instructions",
        bold=True
    )
    row += 1

    # Add product rows.
    first_product_row = row
    lineno = 1
    for category_group in products_by_category:
        (category_sort_id, category), products = category_group
        set_cell(worksheet, row, col_description, category, bold=True)
        row += 1
        for category, sku, cost, description, htsus_no in products:
            set_cell(worksheet, row, col_line_no, lineno)
            set_cell(worksheet, row, col_sku, sku)
            set_cell(worksheet, row, col_description, description)
            style = set_cell(worksheet, row, col_price, cost).style
            style.number_format.format_code = "0.00"
            if random.randint(1, 10) < 4:
                set_cell(worksheet, row, col_qty, random.randint(2, 8) * 10)
            total_formula = "=IF(%s%i=\"\", \"\", %s%i * %s%i)" % (
                col_letter(col_qty),
                row + 1,
                col_letter(col_price),
                row + 1,
                col_letter(col_qty),
                row + 1
            )
            style = set_cell(worksheet, row, col_total, total_formula).style
            style.number_format.format_code = "#,###.00"
            set_cell(worksheet, row, col_htsus_no, htsus_no)
            row += 1
            lineno += 1
    last_product_row = row - 1

    # Set column widths.
    worksheet.column_dimensions[col_letter(col_line_no)].width = 7
    worksheet.column_dimensions[col_letter(col_sku)].width = 6
    worksheet.column_dimensions[col_letter(col_description)].width = 60
    worksheet.column_dimensions[col_letter(col_price)].width = 5
    worksheet.column_dimensions[col_letter(col_qty)].width = 5
    worksheet.column_dimensions[col_letter(col_total)].width = 8
    worksheet.column_dimensions[col_letter(col_htsus_no)].width = 12
    worksheet.column_dimensions[col_letter(col_instructions)].width = 30

    return row, col_total, first_product_row, last_product_row


def add_totals(
    worksheet,
    config,
    row,
    col_total,
    first_product_row,
    last_product_row
):
    """Add subtotals and totals."""
    col_value_name = col_total - 1

    # Subtotal.
    sub_total_row = row
    set_cell(
        worksheet,
        row,
        col_value_name,
        "Subtotal",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    sub_total_formula = "=SUM(%s%i:%s%i)" % (
        col_letter(col_total),
        first_product_row + 1,
        col_letter(col_total),
        last_product_row + 1
    )
    style = set_cell(worksheet, row, col_total, sub_total_formula).style
    style.number_format.format_code = "#,###.00"
    row += 1

    # Discount.
    discount_row = row
    percent_discount = config.getfloat("invoice", "percent_discount")
    set_cell(
        worksheet,
        row,
        col_value_name,
        "%s%% Discount" % percent_discount,
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    discount_formula = "=%s%i * %f" % (
        col_letter(col_total),
        sub_total_row + 1,
        -percent_discount / 100.0
    )
    style = set_cell(worksheet, row, col_total, discount_formula).style
    style.number_format.format_code = "#,###.00"
    row += 1

    # Total.
    set_cell(
        worksheet,
        row,
        col_value_name,
        "Total",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    total_formula = "=%s%i + %s%i" % (
        col_letter(col_total),
        sub_total_row + 1,
        col_letter(col_total),
        discount_row + 1
    )
    style = set_cell(worksheet, row, col_total, total_formula).style
    style.number_format.format_code = "#,###.00"
    row += 1

    return(row)


def add_special_instructions(worksheet, row):
    """Add special instructions."""
    col_value_name = 2

    # Special instructions.
    set_cell(
        worksheet,
        row,
        col_value_name,
        "Special Instructions:",
        bold=True
    )


def add_invoice(worksheet, config, products_by_category):
    """Create the PO-Invoice worksheet."""

    # Prepare worksheet.
    worksheet.title = "PO-Invoice"

    # Add title.
    add_title(worksheet)

    # Add header.
    row = 2
    row = add_header(worksheet, config, row)

    # Blank row.
    row += 1

    # Add products.
    row, col_total, first_product_row, last_product_row = add_products(
        worksheet,
        row,
        products_by_category
    )

    # Blank row.
    row += 1

    # Add subtotals and total.
    row = add_totals(
        worksheet,
        config,
        row,
        col_total,
        first_product_row,
        last_product_row
    )

    # Add special instructions.
    add_special_instructions(worksheet, row + 1)


def add_instructions(worksheet):
    """Create the Instructions worksheet."""

    # Prepare worksheet.
    worksheet.title = "Instructions"
    col_line_no = 0
    col_instruction = 1
    row = 0

    # Add instructions.
    instructions = [
        "Change cell A1 from \"Purchase Order\" to \"Commercial Invoice\".",
        "Change quantities to match actual shipped amounts.",
        "If adding a product row, fix Subtotal formula to include new row.",
        "Save to a different file.  Keep original PO separate from Invoice."
    ]
    for line_no, instruction in enumerate(instructions):
        set_cell(
            worksheet,
            row,
            col_line_no,
            "%i. " % (line_no + 1),
            alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
        )
        set_cell(worksheet, row, col_instruction, instruction)
        row += 1

    # Set column widths.
    worksheet.column_dimensions[col_letter(col_line_no)].width = 3
    worksheet.column_dimensions[col_letter(col_instruction)].width = 70


def generate_xlsx(config, products_by_category, xlsx_filename):
    """Generate the XLS file."""

    # Construct a document.
    workbook = openpyxl.workbook.Workbook()

    # Create PO-Invoice worksheet.
    add_invoice(workbook.worksheets[0], config, products_by_category)

    # Create Instructions worksheet.
    add_instructions(workbook.create_sheet())

    # Write to file.
    workbook.save(xlsx_filename)


CATEGORY_SORT_KEY = dict()

def add_category_to_sort_key(config, category):
    global CATEGORY_SORT_KEY
    if category not in CATEGORY_SORT_KEY:
        prefixes = config.get("DEFAULT", "category_sort").split(",")
        for index, prefix in enumerate(prefixes):
            if category.startswith(prefix):
                CATEGORY_SORT_KEY[category] = index
                return
        print "ERROR: Category '%s' not found in category_sort" % category
        sys.exit(1)


def sort_key(tupl):
    """Return a sort key of a (category, sku, cost, description, ...) tuple."""
    category = tupl[0]
    description = tupl[3]
    return "%i%s" % (CATEGORY_SORT_KEY[category], description)


def load_products_by_category(config, csv_filename):
    """Load product data."""

    # Read the input file, extracting just the fields we need.
    is_header = True
    data = list()
    for fields in csv.reader(open(csv_filename)):
        if is_header:
            category_field = fields.index("Category")
            sku_field = fields.index("SKU")
            cost_field = fields.index("Cost")
            name_field = fields.index("Product Name")
            teaser_field = fields.index("Teaser")
            discontinued_field = fields.index("Discontinued Item")
            is_header = False
        elif fields[discontinued_field] == "N":
            category = fields[category_field]
            sku = fields[sku_field]
            cost = fields[cost_field]
            description = "%s: %s" % (
                fields[name_field],
                fields[teaser_field].replace("&quot;", "\"")
            )
            htsus_no = "7117.90.9000"
            data.append((category, sku, cost, description, htsus_no))
            add_category_to_sort_key(config, category)

    # Sort by category.
    data = sorted(data, key=sort_key)

    # Group by category.
    products_by_category = list()
    for key, group in itertools.groupby(data, lambda x: x[0]):
        products = list(group)
        category = products[0][0]
        category_group = ((key, category), products)
        products_by_category.append(category_group)

    return products_by_category


def main():
    """main"""
    option_parser = optparse.OptionParser(
        usage="usage: %prog [options]\n" +
        "  Generates a Purchase Order / Commercial Invoice."
    )
    option_parser.add_option(
        "--config",
        action="store",
        dest="config",
        metavar="FILE",
        default="cctools.cfg",
        help="configuration filename (default=%default)"
    )
    option_parser.add_option(
        "--prodfile",
        action="store",
        dest="csv_filename",
        metavar="FILE",
        default="products.csv",
        help="input product list filename (default=%default)"
    )
    option_parser.add_option(
        "--outfile",
        action="store",
        dest="xlsx_filename",
        metavar="FILE",
        default=TODAY + "-Invoice.xlsx",
        help="output XLSX filename (default=%default)"
    )

    (options, args) = option_parser.parse_args()
    if len(args) != 0:
        option_parser.error("invalid argument")

    # Read config file.
    config = ConfigParser.RawConfigParser()
    config.readfp(open(options.config))

    products_by_category = load_products_by_category(
        config,
        options.csv_filename
    )

    generate_xlsx(config, products_by_category, options.xlsx_filename)

    return 0


if __name__ == "__main__":
    main()
