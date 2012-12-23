#!/usr/bin/env python

"""
Generates a Purchase Order / Commercial Invoice.
"""

import ConfigParser
import cctools
import itertools
import openpyxl  # sudo apt-get install python-openpyxl
import optparse
import sys
import datetime

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
    """Return column letter for given column."""
    return chr(ord("A") + col)


def row_number(row):
    """Return row number for given row."""
    return row + 1


def add_title(worksheet):
    """Add worksheet title."""
    style = set_cell(worksheet, 0, 0, "Purchase Order", bold=True).style
    style.font.size = 20
    # merge_cells not supported by openpyxl-1.5.6 (Ubuntu 12.04)
    #worksheet.merge_cells(start_row=0, start_column=0, end_row=0, end_column=2)
    worksheet.row_dimensions[1].height = 25


def add_header(options, config, worksheet, row):
    """Add PO/Invoice header."""
    col_value_name = 2
    col_value = col_value_name + 1

    # Prefixing a value with a single quote forces it to be considered text.
    # NOTE: Not supported by localc-3.5.

    set_cell(
        worksheet,
        row,
        col_value_name,
        "PO/Invoice #: ",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    set_cell(worksheet, row, col_value, "'" + options.number)
    row += 1

    set_cell(
        worksheet,
        row,
        col_value_name,
        "Date: ",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    date_str = datetime.date.today().strftime("'%B %d, %Y")
    set_cell(worksheet, row, col_value, date_str)
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
    for key, value in config.items("invoice"):
        if key.startswith("consignee"):
            set_cell(worksheet, row, col_value, value)
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


def add_products(worksheet, row, cc_browser, products):
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

    # Sort products by category, product_name.
    products = sorted(products, key=cc_browser.sort_key_by_category_and_name)

    # Group products by category.
    first_product_row = row
    lineno = 1
    for _, product_group in itertools.groupby(
        products,
        key=cc_browser.sort_key_by_category
    ):
        # Leave a row for the category name.
        category = "unknown"
        category_row = row
        row += 1
        # Add product rows.
        for product in product_group:
            category = product["Category"]
            if product["Discontinued Item"] == "Y":
                continue
            description = "  %s: %s" % (
                product["Product Name"],
                cctools.html_to_plain_text(product["Teaser"])
            )
            htsus_no = product["UPC"]
            set_cell(worksheet, row, col_line_no, lineno)
            set_cell(worksheet, row, col_sku, product["SKU"])
            set_cell(worksheet, row, col_description, description)
            style = set_cell(worksheet, row, col_price, product["Cost"]).style
            style.number_format.format_code = "0.00"
            total_formula = "=IF(%s%i=\"\", \"\", %s%i * %s%i)" % (
                col_letter(col_qty),
                row_number(row),
                col_letter(col_price),
                row_number(row),
                col_letter(col_qty),
                row_number(row)
            )
            style = set_cell(worksheet, row, col_total, total_formula).style
            style.number_format.format_code = "#,###.00"
            set_cell(worksheet, row, col_htsus_no, htsus_no)
            row += 1
            lineno += 1
        # Go back and insert category name.
        set_cell(worksheet, category_row, col_description, category, bold=True)
    last_product_row = row - 1

    # Set column widths.
    worksheet.column_dimensions[col_letter(col_line_no)].width = 8
    worksheet.column_dimensions[col_letter(col_sku)].width = 6
    worksheet.column_dimensions[col_letter(col_description)].width = 60
    worksheet.column_dimensions[col_letter(col_price)].width = 5
    worksheet.column_dimensions[col_letter(col_qty)].width = 5
    worksheet.column_dimensions[col_letter(col_total)].width = 8
    worksheet.column_dimensions[col_letter(col_htsus_no)].width = 12
    worksheet.column_dimensions[col_letter(col_instructions)].width = 30

    return row, col_qty, col_total, first_product_row, last_product_row


def add_totals(
    worksheet,
    config,
    row,
    col_qty,
    col_total,
    first_product_row,
    last_product_row
):
    """Add subtotals and totals."""

    # Total quantity.
    col_value_name = col_qty - 1
    set_cell(
        worksheet,
        row,
        col_value_name,
        "Total Qty:",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    total_qty_formula = "=SUM(%s%i:%s%i)" % (
        col_letter(col_qty),
        row_number(first_product_row),
        col_letter(col_qty),
        row_number(last_product_row)
    )
    set_cell(worksheet, row, col_qty, total_qty_formula)
    row += 2

    # Subtotal.
    col_value_name = col_total - 1
    sub_total_row = row
    set_cell(
        worksheet,
        row,
        col_value_name,
        "Subtotal:",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    sub_total_formula = "=SUM(%s%i:%s%i)" % (
        col_letter(col_total),
        row_number(first_product_row),
        col_letter(col_total),
        row_number(last_product_row)
    )
    style = set_cell(worksheet, row, col_total, sub_total_formula).style
    style.number_format.format_code = "#,###.00"
    row += 1

    # Discount.
    percent_discount = config.getfloat("invoice", "percent_discount")
    set_cell(
        worksheet,
        row,
        col_value_name,
        "%s%% Discount:" % percent_discount,
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    discount_formula = "=%s%i * %f" % (
        col_letter(col_total),
        row_number(sub_total_row),
        -percent_discount / 100.0
    )
    style = set_cell(worksheet, row, col_total, discount_formula).style
    style.number_format.format_code = "#,###.00"
    last_adjustment_row = row
    row += 1

    # Adjustments.
    for key, value in config.items("invoice"):
        if key.startswith("adjustment"):
            set_cell(
                worksheet,
                row,
                col_value_name,
                value + ":",
                bold=True,
                alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
            )
            last_adjustment_row = row
            row += 1

    # Total.
    set_cell(
        worksheet,
        row,
        col_value_name,
        "Total:",
        bold=True,
        alignment_horizontal=ALIGNMENT_HORIZONTAL_RIGHT
    )
    total_formula = "=SUM(%s%i:%s%i)" % (
        col_letter(col_total),
        row_number(sub_total_row),
        col_letter(col_total),
        row_number(last_adjustment_row)
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
        "General Special Instructions:",
        bold=True
    )


def add_invoice(options, config, cc_browser, products, worksheet):
    """Create the PO-Invoice worksheet."""

    # Prepare worksheet.
    worksheet.title = "PO-Invoice"

    # Add title.
    add_title(worksheet)

    # Add header.
    row = 2
    row = add_header(options, config, worksheet, row)

    # Blank row.
    row += 1

    # Add products.
    row, col_qty, col_total, first_product_row, last_product_row = add_products(
        worksheet,
        row,
        cc_browser,
        products
    )

    # Blank row.
    row += 1

    # Add subtotals and total.
    row = add_totals(
        worksheet,
        config,
        row,
        col_qty,
        col_total,
        first_product_row,
        last_product_row
    )

    # Add special instructions.
    add_special_instructions(worksheet, row_number(row))


def add_instructions(config, worksheet):
    """Create the Instructions worksheet."""

    # Prepare worksheet.
    worksheet.title = "Instructions"
    col_instruction = 0
    row = 0

    # Add instructions.
    for key, value in config.items("invoice"):
        if key.startswith("instruction"):
            if value == "[EMPTY]":
                row += 1
                continue
            bold = False
            if value.startswith("[BOLD]"):
                bold = True
                value = value[6:]
            cell = set_cell(worksheet, row, col_instruction, value)
            if bold:
                cell.style.font.bold = bold
            row += 1

    # Set column widths.
    worksheet.column_dimensions[col_letter(col_instruction)].width = 70


def generate_xlsx(options, config, cc_browser, products):
    """Generate the XLS file."""

    # Construct a document.
    workbook = openpyxl.workbook.Workbook()

    # Create PO-Invoice worksheet.
    add_invoice(options, config, cc_browser, products, workbook.worksheets[0])

    # Create Instructions worksheet.
    add_instructions(config, workbook.create_sheet())

    # Write to file.
    workbook.save(options.xlsx_filename)


def main():
    """main"""
    now = datetime.datetime.now()
    default_number = now.strftime("%y%m%d00")
    default_xlsx_filename = now.strftime("%Y-%m-%d-PurchaseOrder.xlsx")

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
        "--number",
        action="store",
        dest="number",
        metavar="NUM",
        default=default_number,
        help="PO/Invoice number (default=%default)"
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

    # Fetch products list.
    products = cc_browser.get_products()

    # Generate spreadsheet.
    if options.verbose:
        sys.stderr.write("Generating %s\n" % options.xlsx_filename)
    generate_xlsx(options, config, cc_browser, products)

    if options.verbose:
        sys.stderr.write("Generation complete\n")
    return 0


if __name__ == "__main__":
    main()
