#!/usr/bin/env python2

"""
Generates a price list from CoreCommerce data in PDF form.  Prices
are adjusted to include sales tax and are rounded to even dollar
amounts.  The intent is to use the price list at fairs and shows to
avoid the handling of change.  It also makes accounting easier because
you can deal with round numbers.

TODO: page orientation
"""

import ConfigParser
import argparse
import cctools
import datetime
import itertools
import math
import os
import reportlab.lib  # sudo apt-get install python-reportlab
import reportlab.platypus
import sys

# Best reportlab reference is the ReportLab User's Guide.

# Convenience constants.
INCH = reportlab.lib.units.inch

# Currency related constants (pre_symbol, post_symbol, rounder).
CURRENCY_INFO = {
    "ugx": ("", " USh", 1000.0),
    "usd": ("$", "", 1.0)
}


def calc_price_inc_tax(args, price, price_multiplier):
    """Calculate a price including tax."""
    price_inc_tax = float(price) * price_multiplier
    pre_symbol, post_symbol, rounder = CURRENCY_INFO[args.currency]
    whole_price_inc_tax = max(
        rounder,
        math.floor((price_inc_tax + 0.5) / rounder) * rounder
    )
    return "{}{:,.0f}{}".format(pre_symbol, whole_price_inc_tax, post_symbol)


def on_page(canvas, doc):
    """Add page header and footer.  Called for each page."""
    page_width = doc.pagesize[0]
    page_height = doc.pagesize[1]
    canvas.saveState()
    canvas.setFont("Helvetica-Bold", 16)
    canvas.drawCentredString(
        page_width / 2.0,
        page_height - 0.5 * INCH,
        doc.my_title
    )
    canvas.setFont("Helvetica", 9)
    canvas.drawString(
        0.5 * INCH,
        0.5 * INCH,
        "PricePreTax = PriceIncTax / (1 + TaxPercent / 100)"
    )
    today_str = datetime.date.today().isoformat()
    canvas.drawRightString(
        page_width - 0.5 * INCH,
        0.5 * INCH,
        "Revised: %s" % today_str
    )
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
            id="col%i" % col,
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
    products,
    price_multiplier
):
    """Generate a PDF given a list of products by category."""
    # Construct a document.
    doc_title = config.get("price_list", "title")
    doc = reportlab.platypus.BaseDocTemplate(
        args.pdf_file,
        pagesize=reportlab.lib.pagesizes.letter,
        title=doc_title,
        # showBoundary=True  # debug
    )

    # Construct a frame for each column.
    frames = create_col_frames(doc, args.ncols)

    # Construct a template and add it to the document.
    doc.my_title = doc_title
    doc.addPageTemplates(
        reportlab.platypus.PageTemplate(
            id="mytemplate",
            frames=frames,
            onPage=on_page
        )
    )

    # Construct a story and add it to the document.
    category_style = reportlab.lib.styles.ParagraphStyle(
        name="category",
        spaceBefore=0.15 * INCH,
        spaceAfter=0.05 * INCH,
        fontName="Helvetica-Bold",
        fontSize=12
    )
    greybar_color = reportlab.lib.colors.Whiter(
        reportlab.lib.colors.lightgrey,
        0.5
    )
    table_indent = 0.1 * INCH
    table_width = frames[0].width - table_indent
    price_width = 0.4 * INCH
    col_widths = [table_width - price_width, price_width]
    story = []

    # Sort products by category, product_name.
    if args.categories:
        cc_browser.set_category_sort_order(args.categories)
    products = sorted(products, key=cc_browser.sort_key_by_category_and_name)

    # Group products by category.
    body_fontsize = float(config.get("price_list", "body_fontsize"))
    for _, product_group in itertools.groupby(
        products,
        key=cc_browser.sort_key_by_category
    ):
        # TableStyle cell formatting commands.
        styles = [
            ("FONTSIZE", (0, 0), (-1, -1), body_fontsize),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("FONT", (1, 0), (1, -1), "Courier-Bold"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            # ("GRID", (0, 0), (-1, -1), 1.0, reportlab.lib.colors.black)
        ]
        table_data = list()
        for product in product_group:
            category = product["Category"]
            product_name = cctools.plain_text_to_html(product["Product Name"])
            price = calc_price_inc_tax(
                args,
                product["Price"],
                price_multiplier
            )
            if args.display_sku:
                row = ("%s (%s)" % (product_name, product["SKU"]), price)
            else:
                row = (product_name, price)
            table_data.append(row)
        if len(table_data) == 0:
            continue
        if args.greybar_interval > 1:
            for row in range(0, len(table_data), args.greybar_interval):
                styles.append(
                    ("BACKGROUND", (0, row), (1, row), greybar_color)
                )
        table = reportlab.platypus.Table(
            data=table_data,
            colWidths=col_widths,
            style=styles
        )
        story.append(
            reportlab.platypus.KeepTogether([
                reportlab.platypus.Paragraph(
                    cctools.plain_text_to_html(category),
                    category_style
                ),
                reportlab.platypus.Indenter(left=table_indent),
                table,
                reportlab.platypus.Indenter(left=-table_indent)
            ])
        )
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
    default_config = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "cctools.cfg"
    )

    arg_parser = argparse.ArgumentParser(
        description="Generates a price list from CoreCommerce data."
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
        dest="pdf_file",
        metavar="PDF_FILE",
        default="PriceListRetailTaxInc.pdf",
        help="output PDF filename (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--display-sku",
        action="store_true",
        dest="display_sku",
        default=False,
        help="display SKU with product name"
    )
    arg_parser.add_argument(
        "--ncols",
        type=int,
        dest="ncols",
        metavar="N",
        default=2,
        help="number of report columns (default=%(default)i)"
    )
    arg_parser.add_argument(
        "--greybar-interval",
        type=int,
        dest="greybar_interval",
        metavar="N",
        default=2,
        help="greybar interval (default=%(default)i)"
    )
    arg_parser.add_argument(
        "--tax",
        type=float,
        dest="tax_percent",
        metavar="PCT",
        default=8.4,
        help="tax rate in percent (default=%(default).2f)"
    )
    arg_parser.add_argument(
        "--currency",
        dest="currency",
        choices=["usd", "ugx"],
        default="usd",
        help="currency (default=%(default)s)"
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

    # Read config file.
    config = ConfigParser.SafeConfigParser({
        "body_fontsize": "12"
    })
    config.readfp(open(args.config))

    # Determine price multiplier.
    price_multiplier = 1.0 + args.tax_percent / 100.0
    if args.currency == "ugx":
        price_multiplier *= float(config.get("price_list", "ugx_exchange"))

    # Create a connection to CoreCommerce.
    cc_browser = cctools.CCBrowser(
        config.get("website", "host"),
        config.get("website", "site"),
        config.get("website", "username"),
        config.get("website", "password"),
        verbose=args.verbose
    )

    # Get product list.
    products = get_products(args, cc_browser)

    # Generate PDF file.
    if args.verbose:
        sys.stderr.write("Generating %s\n" % os.path.abspath(args.pdf_file))
    generate_pdf(
        args,
        config,
        cc_browser,
        products,
        price_multiplier
    )

    if args.verbose:
        sys.stderr.write("Generation complete\n")
    return 0

if __name__ == "__main__":
    main()
