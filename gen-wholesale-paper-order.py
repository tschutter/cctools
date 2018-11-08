#!/usr/bin/env python2

"""
Generates a printable wholesale order form from CoreCommerce data
in PDF form.

TOASK:
  * Handling of variants with additional price

TODO:
  * [X] TypeError: cannot concatenate 'str' and 'float' objects
  * [X] fix price calc
  * [X] qty
  * [X] total
  * [ ] subtotal
  * [ ] shipping
  * [ ] adjustment
  * [ ] total (grand)

  * [ ] email
  * [ ] phone
  * [ ] item no

  * --add-size?  what is the real field name?
  * page orientation
"""

import argparse
import datetime
import itertools
import logging
import os

import ConfigParser
import calc_price
import cctools
import notify_send_handler
import reportlab.lib  # sudo apt-get install python-reportlab
import reportlab.platypus

# Best reportlab reference is the ReportLab User's Guide.

# Convenience constants.
INCH = reportlab.lib.units.inch


def on_page(canvas, doc):
    """Add page header and footer.  Called for each page."""

    # Save current state.
    canvas.saveState()

    # Get page geometry.
    page_width = doc.pagesize[0]
    page_height = doc.pagesize[1]

    # Draw header.
    canvas.setFont("Helvetica-Bold", 16)
    canvas.drawCentredString(
        page_width / 2.0,
        page_height - 0.5 * INCH,
        doc.my_title
    )

    # Draw left footer.
    canvas.setFont("Helvetica", 9)
    canvas.drawString(
        0.5 * INCH,
        0.5 * INCH,
        doc.left_footer
    )

    # Draw right footer.
    today_str = datetime.date.today().isoformat()
    canvas.drawRightString(
        page_width - 0.5 * INCH,
        0.5 * INCH,
        "Revised: {}".format(today_str)
    )

    # Restore current state.
    canvas.restoreState()


def create_col_frames(doc, ncols):
    """Create frames for each column."""
    page_width = doc.pagesize[0]
    page_height = doc.pagesize[1]
    lr_margin = 0.5 * INCH
    col_spacing = 0.5 * INCH
    tb_margin = 0.7 * INCH
    frame_width = (
        page_width - lr_margin * 2 - col_spacing * (ncols - 1)
    ) / ncols
    frame_height = page_height - tb_margin * 2
    frames = [
        reportlab.platypus.Frame(
            id="col{}".format(col),
            x1=lr_margin + (frame_width + col_spacing) * col,
            y1=tb_margin,
            width=frame_width,
            height=frame_height,
            leftPadding=0,
            bottomPadding=0,
            rightPadding=0,
            topPadding=0
        )
        for col in range(ncols)
    ]
    return frames


def generate_pdf(
    args,
    config,
    cc_browser,
    products
):
    """Generate a PDF given a list of products by category."""

    # Define some constants.
    color_black = reportlab.lib.colors.black

    # Construct a document.
    title = config.get("wholesale_paper_order", "title")
    doc = reportlab.platypus.BaseDocTemplate(
        args.pdf_file,
        pagesize=reportlab.lib.pagesizes.letter,
        title=title,
        # showBoundary=True  # debug
    )

    # Set some values that can be used in callbacks.
    doc.my_title = title
    doc.left_footer = config.get("wholesale_paper_order", "left_footer")

    # Construct a frame for each column.
    frames = create_col_frames(doc, args.ncols)

    # Construct a template and add it to the document.
    doc.addPageTemplates(
        reportlab.platypus.PageTemplate(
            id="mytemplate",
            frames=frames,
            onPage=on_page
        )
    )

    # Construct a story and add it to the document.
    greybar_color = reportlab.lib.colors.Whiter(
        reportlab.lib.colors.lightgrey,
        0.03
    )
    table_indent = 0.1 * INCH
    table_width = frames[0].width - table_indent
    price_width = 0.4 * INCH
    qty_width = 0.4 * INCH
    total_width = 0.5 * INCH
    name_width = table_width - price_width - qty_width - total_width
    col_widths = [name_width, price_width, qty_width, total_width]
    story = []

    # Sort products by category, product_name.
    if args.categories:
        cc_browser.set_category_sort_order(args.categories)
    products = sorted(products, key=cc_browser.product_key_by_cat_and_name)

    # Setup styles.
    body_fontsize = float(config.get("wholesale_paper_order", "body_fontsize"))
    row_padding = float(config.get("wholesale_paper_order", "row_padding"))
    base_styles = [
        ("FONTSIZE", (0, 0), (-1, -1), body_fontsize),
        ("FONT", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("ALIGN", (2, 0), (3, -1), "CENTER"),
        ("FONT", (1, 1), (1, -1), "Courier-Bold"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), row_padding),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3 + row_padding)
    ]

    # Group products by category.
    for _, product_group in itertools.groupby(
        products,
        key=cc_browser.product_key_by_category
    ):
        # Make a new styles instance just for this table.
        styles = list(base_styles)

        # Assemble table data for the product_group.
        table_data = list()
        row = 0

        for product in product_group:
            category = product["Category"]

            if row == 0:
                row_data = (category, "", "Qty", "Total")
                table_data.append(row_data)
                row += 1

            # Product name column.
            product_name = product["Product Name"]
            if args.add_sku:
                product_name = "{} - {}".format(product["SKU"], product_name)

            # Price column formatted as a whole number string like "$1,234.99".
            price = calc_price.calc_wholesale_price(
                product["Price"],
                args.wholesale_fraction
            )
            price = "${:,.2f}".format(price)

            row_data = (product_name, price, "", "")

            table_data.append(row_data)
            row += 1
        if len(table_data) == 0:
            continue

        # Grey background color to highlight alternate products.
        if args.greybar_interval > 1:
            for row in range(1, len(table_data), args.greybar_interval):
                styles.append(
                    ("BACKGROUND", (0, row), (-1, row), greybar_color)
                )

        # Draw a grid.
        styles.extend(
            [
                ('INNERGRID', (2, 1), (-1, -1), 0.25, color_black),
                ('BOX', (2, 1), (-1, -1), 0.25, color_black)
            ]
        )

        # Create the table.
        table = reportlab.platypus.Table(
            data=table_data,
            colWidths=col_widths,
            style=styles
        )

        story.append(table)
        story.append(reportlab.platypus.Spacer(width=0, height=0.04 * INCH))

    # Create subtotal table.
    table_data = list()
    table_data.append(("Order Date", "", "Subtotal", ""))
    table_data.append(("Contact", "", "Shipping", ""))
    table_data.append(("Tax ID", "", "Adjustment", ""))
    table_data.append(("", "", "Total", ""))
    styles = [
        ("FONTSIZE", (0, 0), (-1, -1), body_fontsize),
        ("FONT", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "RIGHT"),
        ('INNERGRID', (1, 0), (1, -1), 0.25, color_black),
        ('BOX', (1, 0), (1, -1), 0.25, color_black),
        ('INNERGRID', (3, 0), (3, -1), 0.25, color_black),
        ('BOX', (3, 0), (3, -1), 0.25, color_black)
    ]
    first_width = 0.6 * INCH
    third_width = 0.8 * INCH
    fourth_width = 0.8 * INCH
    second_width = table_width - first_width - third_width - fourth_width
    col_widths = [first_width, second_width, third_width, fourth_width]
    table = reportlab.platypus.Table(
        data=table_data,
        colWidths=col_widths,
        style=styles
    )
    story.append(reportlab.platypus.Spacer(width=0, height=0.1 * INCH))
    story.append(table)

    # Create customer info table.
    table_data = list()
    table_data.append(("Store", ""))
    table_data.append(("Address", ""))
    table_data.append(("", ""))
    table_data.append(("Phone", ""))
    table_data.append(("Email", ""))
    table_data.append(("Tax ID", ""))
    styles = [
        ("FONTSIZE", (0, 0), (-1, -1), body_fontsize),
        ("FONT", (0, 0), (0, -1), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (0, -1), "RIGHT"),
        ('INNERGRID', (1, 0), (-1, -1), 0.25, color_black),
        ('BOX', (1, 0), (-1, -1), 0.25, color_black)
    ]
    name_width = 0.6 * INCH
    value_width = table_width - name_width
    col_widths = [name_width, value_width]
    table = reportlab.platypus.Table(
        data=table_data,
        colWidths=col_widths,
        style=styles
    )
    story.append(reportlab.platypus.Spacer(width=0, height=0.1 * INCH))
    story.append(table)

    doc.build(story)


def get_products(args, cc_browser):
    """Get product list from CoreCommerce and filter it."""

    # Fetch products list.
    products = cc_browser.get_products()

    # Remove bad products.
    products = [
        p for p in products if
        p["Category"] != "" and p["Product Name"] != ""
    ]

    # Remove products that are not available online and discontinued.
    products = [
        p for p in products if
        p["Available"] == "Y" or p["Discontinued Item"] == "N"
    ]

    # Remove products that are not requested.
    if args.categories:
        cc_browser.set_category_sort_order(args.categories)
        products = [p for p in products if p["Category"] in args.categories]
    elif args.epclude_categories:
        products = [
            p for p in products if p["Category"] not in args.exclude_categories
        ]

    # Remove excluded SKUs.
    if args.exclude_skus:
        products = [
            p for p in products if str(p["SKU"]) not in args.exclude_skus
        ]

    return products


def main():
    """main"""
    # Construct default filename of configuration file.
    default_config = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "cctools.cfg"
    )
    now = datetime.datetime.now()
    default_pdf_filename = now.strftime("%Y-%m-%d-WholesaleOrderForm.pdf")

    arg_parser = argparse.ArgumentParser(
        description=(
            "Generates a printable wholesale order form "
            "from CoreCommerce data."
        )
    )
    arg_parser.add_argument(
        "--config",
        metavar="FILE",
        default=default_config,
        help="configuration filename (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--category",
        action="append",
        dest="categories",
        metavar="CAT",
        help="include category in output"
    )
    arg_parser.add_argument(
        "--exclude-category",
        action="append",
        dest="exclude_categories",
        metavar="CAT",
        help="exclude category from output"
    )
    arg_parser.add_argument(
        "--exclude-sku",
        action="append",
        dest="exclude_skus",
        metavar="SKU",
        help="exclude SKU from output"
    )
    arg_parser.add_argument(
        "--pdf-file",
        metavar="FILE",
        default=default_pdf_filename,
        help="output PDF filename (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--add-sku",
        dest="add_sku",
        action="store_true",
        default=False,
        help="append SKU to product name"
    )
    arg_parser.add_argument(
        "--ncols",
        type=int,
        metavar="N",
        default=2,
        help="number of report columns (default=2)"
    )
    arg_parser.add_argument(
        "--greybar-interval",
        type=int,
        metavar="N",
        default=2,
        help="greybar interval (default=%(default)i)"
    )
    arg_parser.add_argument(
        "--wholesale-fraction",
        metavar="FRAC",
        default=0.5,
        help="wholesale price fraction (default=%(default).2f)"
    )
    arg_parser.add_argument(
        "--verbose",
        action="store_true",
        default=False,
        help="display progress messages"
    )

    # Parse command line arguments.
    args = arg_parser.parse_args()
    if args.categories and args.exclude_categories:
        arg_parser.error("--category and --exclude-category specified")

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
    config = ConfigParser.SafeConfigParser({
        "body_fontsize": "12",
        "row_padding": "0",
        "title": "Wholesale Order"
    })
    config.readfp(open(args.config))

    # Create a connection to CoreCommerce.
    cc_browser = cctools.CCBrowser(
        config.get("website", "base_url"),
        config.get("website", "username"),
        config.get("website", "password")
    )

    # Get product list.
    products = get_products(args, cc_browser)

    # Generate PDF file.
    logger.debug("Generating {}\n".format(os.path.abspath(args.pdf_file)))
    generate_pdf(
        args,
        config,
        cc_browser,
        products
    )

    logger.debug("Generation complete")
    return 0

if __name__ == "__main__":
    main()
