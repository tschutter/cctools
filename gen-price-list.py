#!/usr/bin/env python2

"""
Generates a price list from CoreCommerce data in PDF form.

Prices are adjusted to include any discount and sales tax.  They are
then rounded to the nearest dollar.  The intent is to use the price
list at fairs and shows to avoid the handling of change.  It also
makes accounting easier because you can deal with round numbers.

The average sales tax rate is used rather than a specific rate so that
the same price list can be used in multiple jurisdictions.  Because
the result is rounded to the nearest dollar, the exact sales tax rate
is not significant.

TODO:
  fix blank first page if --add-teaser
  --add-size?  what is the real field name?
  page orientation
"""

import ConfigParser
import argparse
import calc_price
import cctools
import datetime
import itertools
import logging
import notify_send_handler
import os
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
        "Price includes {:g}% discount and sales tax".format(
            doc.discount_percent
        )
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
    doc.discount_percent = args.discount_percent
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
    products = sorted(products, key=cc_browser.product_key_by_cat_and_name)

    # Setup styles.
    body_fontsize = float(config.get("price_list", "body_fontsize"))
    row_padding =  float(config.get("price_list", "row_padding"))
    base_styles = [
        ("FONTSIZE", (0, 0), (-1, -1), body_fontsize),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("FONT", (1, 0), (1, -1), "Courier-Bold"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), row_padding),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3 + row_padding),
        # ("GRID", (0, 0), (-1, -1), 1.0, reportlab.lib.colors.black)
    ]

    # Group products by category.
    first = True
    for _, product_group in itertools.groupby(
        products,
        key=cc_browser.product_key_by_category
    ):
        # Make a new styles instance just for this table.
        styles = list(base_styles)
        if args.add_teaser:
            rows_per_product = 2
        else:
            rows_per_product = 1

        # Assemble table data for the product_group.
        table_data = list()
        row = 0
        for product in product_group:
            category = product["Category"]
            product_name = product["Product Name"]
            price = calc_price.calc_event_price(
                product["Price"],
                args.discount_percent,
                args.avg_tax_percent
            )
            # Format price as whole number string like "$1,234".
            price = "${:,.0f}".format(price)
            if args.add_sku:
                row_data = (
                    "{} ({})".format(product_name, product["SKU"]),
                    price
                )
            else:
                row_data = (product_name, price)
            table_data.append(row_data)
            row += 1
            if args.add_teaser:
                # Strip HTML formatting.
                product_teaser = cctools.html_to_plain_text(product["Teaser"])
                row_data = (product_teaser,)
                table_data.append(row_data)
                # Indent the teaser.
                styles.append(("LEFTPADDING", (0, row), (0, row), 18))
                row += 1
        if len(table_data) == 0:
            continue

        # Prevent page splitting in the middle of a product.
        if rows_per_product > 1:
            for row in range(0, len(table_data), rows_per_product):
                styles.append(
                    ("NOSPLIT", (0, row), (0, row + rows_per_product - 1))
                )

        # Grey background color to highlight alternate products.
        if args.greybar_interval > 1:
            interval = args.greybar_interval * rows_per_product
            for start_row in range(0, len(table_data), interval):
                for sub_row in range(0, rows_per_product):
                    row = start_row + sub_row
                    styles.append(
                        ("BACKGROUND", (0, row), (-1, row), greybar_color)
                    )

        # Create the table.
        table = reportlab.platypus.Table(
            data=table_data,
            colWidths=col_widths,
            style=styles
        )

        # Create a group of the category name and the products.
        category_group = [
            reportlab.platypus.Paragraph(category, category_style),
            reportlab.platypus.Indenter(left=table_indent),
            table,
            reportlab.platypus.Indenter(left=-table_indent)
        ]
        # There is a bug in platypus where if a KeepTogether is at
        # the top of a page and the KeepTogether is longer than a
        # page, then a blank page will be emitted.  We can't always
        # know when this will happen, but we know for certain it will
        # happen for the first category group.
        if first:
            for flowable in category_group:
                story.append(flowable)
            first = False
        else:
            story.append(reportlab.platypus.KeepTogether(category_group))

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
        metavar="PDF_FILE",
        default="PriceListRetailTaxInc.pdf",
        help="output PDF filename (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--add-sku",
        action="store_true",
        default=False,
        help="append SKU to product name"
    )
    arg_parser.add_argument(
        "--add-teaser",
        action="store_true",
        default=False,
        help="add teaser line"
    )
    arg_parser.add_argument(
        "--ncols",
        type=int,
        metavar="N",
        default=None,
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
        "--discount",
        type=float,
        dest="discount_percent",
        metavar="PCT",
        default=30,
        help="discount in percent (default=%(default).0f)"
    )
    arg_parser.add_argument(
        "--avg-tax",
        type=float,
        dest="avg_tax_percent",
        metavar="PCT",
        default=8.3,
        help="average sales tax rate in percent (default=%(default).2f)"
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
    if args.ncols is None:
        if args.add_teaser:
            args.ncols = 1
        else:
            args.ncols = 2

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
        "row_padding": "0"
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
