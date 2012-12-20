#!/usr/bin/env python

"""
Generates a price list from CoreCommerce data in PDF form.  Prices
are adjusted to include sales tax and are rounded to even dollar
amounts.  The intent is to use the price list at fairs and shows to
avoid the handling of change.  It also makes accounting easier because
you can deal with round numbers.

sudo apt-get install python-reportlab

Best reportlab reference is the ReportLab User's Guide.

TODO: do what if "Discontinued Item" == ""
TODO: page orientation
"""

import ConfigParser
import cctools
import datetime
import itertools
import math
import optparse
import reportlab.lib
import reportlab.platypus
import sys

INCH = reportlab.lib.units.inch

TITLE = "CoHU PRICE LIST (RETAIL, TAX INCLUDED)"

def clean_text(text):
    """Cleanup HTML in text."""
    text = text.replace("&", "&amp;")
    return text


def calc_price_inc_tax(price, tax_fraction):
    """Calculate a price including tax."""
    price_inc_tax = float(price) * (1.0 + tax_fraction)
    whole_price_inc_tax = max(1.0, math.floor(price_inc_tax + 0.5))
    return "$%.0f" % whole_price_inc_tax


def on_page(canvas, doc):
    """Add page header and footer.  Called for each page."""
    page_width = doc.pagesize[0]
    page_height = doc.pagesize[1]
    canvas.saveState()
    canvas.setFont("Helvetica-Bold", 16)
    canvas.drawCentredString(page_width / 2.0, page_height - 0.5 * INCH, TITLE)
    canvas.setFont("Helvetica", 9)
    today_str = datetime.date.today().isoformat()
    canvas.drawCentredString(
        page_width / 2.0,
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
    cc_browser,
    products,
    tax_fraction,
    ncols,
    greybar_interval,
    pdf_file
):
    """Generate a PDF given a list of products by category."""
    # Construct a document.
    doc = reportlab.platypus.BaseDocTemplate(
        pdf_file,
        pagesize=reportlab.lib.pagesizes.letter,
        title=TITLE,
        #showBoundary=True
    )

    # Construct a frame for each column.
    frames = create_col_frames(doc, ncols)

    # Construct a template and add it to the document.
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
            if product["Discontinued Item"] == "N":
                row = (
                    clean_text(product["Product Name"]),
                    calc_price_inc_tax(product["Price"], tax_fraction)
                )
                table_data.append(row)
        if len(table_data) == 0:
            continue
        if greybar_interval > 1:
            for row in range(0, len(table_data), greybar_interval):
                styles.append(("BACKGROUND", (0, row), (1, row), greybar_color))
        table = reportlab.platypus.Table(
            data=table_data,
            colWidths=col_widths,
            style=styles
        )
        story.append(
            reportlab.platypus.KeepTogether([
                reportlab.platypus.Paragraph(
                    clean_text(category),
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
    option_parser = optparse.OptionParser(
        usage="usage: %prog [options]\n" +
        "  Generates a price list from CoreCommerce data in PDF form."
    )
    option_parser.add_option(
        "--config",
        action="store",
        metavar="FILE",
        default="cctools.cfg",
        help="configuration filename (default=%default)"
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
        config.get("website", "password")
    )

    # Download products list.
    products = list(cc_browser.get_products())

    # Generate PDF file.
    if options.verbose:
        sys.stderr.write("Generating %s\n" % options.pdf_file)
    generate_pdf(
        cc_browser,
        products,
        options.tax_percent / 100.0,
        options.ncols,
        options.greybar_interval,
        options.pdf_file
    )

    if options.verbose:
        sys.stderr.write("Generation complete\n")
    return 0

if __name__ == "__main__":
    main()
