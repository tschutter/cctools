#!/usr/bin/env python

"""
Generates a price list from CoreCommerce data in PDF form.  Prices
are adjusted to include sales tax and are rounded to even dollar
amounts.  The intent is to use the price list at fairs and shows to
avoid the handling of change.  It also makes accounting easier because
you can deal with round numbers.

TODO: page orientation
"""

import ConfigParser
import cctools
import datetime
import itertools
import math
import optparse
import os
import reportlab.lib  # sudo apt-get install python-reportlab
import reportlab.platypus
import sys

# Best reportlab reference is the ReportLab User's Guide.

# Convenience constants.
INCH = reportlab.lib.units.inch

# Currency related constants (pre_symbol, post_symbol, rounder).
CURRENCY_INFO = {
    "UGX": ("", " USh", 1000.0),
    "USD": ("$", "", 1.0)
}

def calc_price_inc_tax(options, price, price_multiplier):
    """Calculate a price including tax."""
    price_inc_tax = float(price) * price_multiplier
    pre_symbol, post_symbol, rounder = CURRENCY_INFO[options.currency]
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
    options,
    config,
    cc_browser,
    products,
    price_multiplier
):
    """Generate a PDF given a list of products by category."""
    # Construct a document.
    doc_title = config.get("price_list", "title")
    doc = reportlab.platypus.BaseDocTemplate(
        options.pdf_file,
        pagesize=reportlab.lib.pagesizes.letter,
        title=doc_title,
        #showBoundary=True  # debug
    )

    # Construct a frame for each column.
    frames = create_col_frames(doc, options.ncols)

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

    # Remove products that are not requested.
    if options.categories:
        cc_browser.set_category_sort_order(options.categories)
        products = filter(
            lambda x: x["Category"] in options.categories,
            products
        )
    elif options.exclude_categories:
        products = filter(
            lambda x: x["Category"] not in options.exclude_categories,
            products
        )

    # Remove excluded SKUs.
    if options.exclude_skus:
        products = filter(
            lambda x: str(x["SKU"]) not in options.exclude_skus,
            products
        )

    # Removed discontinued products.
    products = filter(lambda x: x["Discontinued Item"] != "Y", products)

    # Sort products by category, product_name.
    products = sorted(products, key=cc_browser.sort_key_by_category_and_name)

    # Group products by category.
    for _, product_group in itertools.groupby(
        products,
        key=cc_browser.sort_key_by_category
    ):
        # TableStyle cell formatting commands.
        styles = [
            ("FONTSIZE", (0, 0), (-1, -1), 12),
            ("ALIGN", (0, 0), (0, -1), "LEFT"),
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("FONT", (1, 0), (1, -1), "Courier-Bold"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            #("GRID", (0, 0), (-1, -1), 1.0, reportlab.lib.colors.black)
        ]
        table_data = list()
        for product in product_group:
            category = product["Category"]
            product_name = cctools.plain_text_to_html(product["Product Name"])
            price = calc_price_inc_tax(
                options,
                product["Price"],
                price_multiplier
            )
            if options.display_sku:
                row = ("%s (%s)" % (product_name, product["SKU"]), price)
            else:
                row = (product_name, price)
            table_data.append(row)
        if len(table_data) == 0:
            continue
        if options.greybar_interval > 1:
            for row in range(0, len(table_data), options.greybar_interval):
                styles.append(("BACKGROUND", (0, row), (1, row), greybar_color))
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


def main():
    """main"""
    default_config = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "cctools.cfg"
    )

    option_parser = optparse.OptionParser(
        usage="usage: %prog [options]\n" +
        "  Generates a price list from CoreCommerce data in PDF form.\n" +
        "  Items that are discontinued are not included."
    )
    option_parser.add_option(
        "--config",
        action="store",
        metavar="FILE",
        default=default_config,
        help="configuration filename (default=%default)"
    )
    option_parser.add_option(
        "--category",
        action="append",
        dest="categories",
        metavar="CAT",
        help="include category in output"
    )
    option_parser.add_option(
        "--exclude-category",
        action="append",
        dest="exclude_categories",
        metavar="CAT",
        help="exclude category from output"
    )
    option_parser.add_option(
        "--exclude-sku",
        action="append",
        dest="exclude_skus",
        metavar="SKU",
        help="exclude SKU from output"
    )
    option_parser.add_option(
        "--pdf-file",
        action="store",
        dest="pdf_file",
        metavar="PDF_FILE",
        default="PriceListRetailTaxInc.pdf",
        help="output PDF filename (default=%default)"
    )
    option_parser.add_option(
        "--display-sku",
        action="store_true",
        dest="display_sku",
        default=False,
        help="display SKU with product name"
    )
    option_parser.add_option(
        "--ncols",
        action="store",
        type="int",
        dest="ncols",
        metavar="N",
        default=2,
        help="number of report columns (default=%default)"
    )
    option_parser.add_option(
        "--greybar-interval",
        action="store",
        type="int",
        dest="greybar_interval",
        metavar="N",
        default=2,
        help="greybar interval (default=%default)"
    )
    option_parser.add_option(
        "--tax",
        action="store",
        type="float",
        dest="tax_percent",
        metavar="PCT",
        default=8.4,
        help="tax rate in percent (default=%default)"
    )
    option_parser.add_option(
        "--currency",
        action="store",
        dest="currency",
        metavar="USD|UGX",
        default="USD",
        help="currency (default=%default)"
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
    if options.categories and options.exclude_categories:
        option_parser.error("--category and --exclude-category specified")

    # Read config file.
    config = ConfigParser.RawConfigParser()
    config.readfp(open(options.config))

    # Determine price multiplier.
    price_multiplier = 1.0 + options.tax_percent / 100.0
    if options.currency == "UGX":
        price_multiplier *= float(config.get("price_list", "ugx_exchange"))

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

    # Clean bad products.
    orig_products = products[:]
    products = list()
    for product in orig_products:
        try:
            if (
                product["Category"] != "" and
                product["Product Name"] != "" and
                float(product["Price"]) > 0.0
            ):
                products.append(product)
        except ValueError:
            pass

    # Generate PDF file.
    if options.verbose:
        sys.stderr.write("Generating %s\n" % os.path.abspath(options.pdf_file))
    generate_pdf(
        options,
        config,
        cc_browser,
        products,
        price_multiplier
    )

    if options.verbose:
        sys.stderr.write("Generation complete\n")
    return 0

if __name__ == "__main__":
    main()
