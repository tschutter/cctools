#!/usr/bin/env python

"""
Generates a price list from CoreCommerce data in PDF form.  Prices
are adjusted to include sales tax and are rounded to even dollar
amounts.  The intent is to use the price list at fairs and shows to
avoid the handling of change.  It also makes accounting easier because
you can deal with round numbers.

sudo apt-get install python-reportlab

Best reportlab reference is the ReportLab User's Guide.

TODO
fetch csv
page orientation
"""

import csv
import datetime
import itertools
import math
import optparse
import reportlab.lib
import reportlab.platypus

INCH = reportlab.lib.units.inch

TITLE = "CoHU PRICE LIST (RETAIL, TAX INCLUDED)"

def on_page(canvas, doc):
    """Add page header and footer.  Called for each page."""
    page_width = doc.pagesize[0]
    page_height = doc.pagesize[1]
    canvas.saveState()
    canvas.setFont('Helvetica-Bold', 16)
    canvas.drawCentredString(page_width / 2.0, page_height - 0.5 * INCH, TITLE)
    canvas.setFont('Helvetica', 9)
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
    col_spacing = 0.2 * INCH
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


def generate_pdf(products_by_category, ncols, greybar_interval, pdf_filename):
    """Generate a PDF given a list of products by category."""
    # Construct a document.
    doc = reportlab.platypus.BaseDocTemplate(
        pdf_filename,
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
    for category in products_by_category:
        category_name, products = category
        # TableStyle cell formatting commands.
        styles = [
            ('FONTSIZE', (0, 0), (-1, -1), 12),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ('FONT', (1, 0), (1, -1), 'Courier-Bold'),
            ('LEFTPADDING', (0, 0), (-1, -1), 3),
            ('RIGHTPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            #('GRID', (0, 0), (-1, -1), 1.0, reportlab.lib.colors.black)
        ]
        if greybar_interval > 1:
            for row in range(0, len(products), greybar_interval):
                styles.append(('BACKGROUND', (0, row), (1, row), greybar_color))
        table = reportlab.platypus.Table(
            data=[(product[1], product[2]) for product in products],
            colWidths=col_widths,
            style=styles
        )
        story.append(
            reportlab.platypus.KeepTogether([
                reportlab.platypus.Paragraph(category_name, category_style),
                reportlab.platypus.Indenter(left=table_indent),
                table,
                reportlab.platypus.Indenter(left=-table_indent)
            ])
        )
    doc.build(story)


def clean_text(text):
    """Cleanup HTML in text."""
    text = text.replace("&", "&amp;")
    return text


def calc_price_inc_tax(price, tax_fraction):
    """Calculate a price including tax."""
    price_inc_tax = float(price) * (1.0 + tax_fraction)
    whole_price_inc_tax = max(1.0, math.floor(price_inc_tax + 0.5))
    return "$%.0f" % whole_price_inc_tax


def load_products_by_category(csv_filename, tax_fraction):
    """Load product data."""

    # Read the input file, extracting just the fields we need.
    is_header = True
    data = list()
    for fields in csv.reader(open(csv_filename)):
        if is_header:
            category_field = fields.index("Category")
            name_field = fields.index("Product Name")
            price_field = fields.index("Price")
            discontinued_field = fields.index("Discontinued Item")
            is_header = False
        elif fields[discontinued_field] == "N":
            data.append(
                (
                    clean_text(fields[category_field]),
                    clean_text(fields[name_field]),
                    calc_price_inc_tax(fields[price_field], tax_fraction)
                )
            )

    # Sort by category.
    data = sorted(data, key=lambda x: x[0])

    # Group by category.
    products_by_category = list()
    for key, group in itertools.groupby(data, lambda x: x[0]):
        category = (key, list(group))
        products_by_category.append(category)

    return products_by_category


def main():
    """main"""
    option_parser = optparse.OptionParser(
        usage="usage: %prog [options]\n" +
        "  Generates a price list from CoreCommerce data in PDF form."
    )
    option_parser.add_option(
        "--infile",
        action="store",
        dest="csv_filename",
        metavar="CSV_FILE",
        default="products.csv",
        help="input product list filename (default=%default)"
    )
    option_parser.add_option(
        "--outfile",
        action="store",
        dest="pdf_filename",
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

    (options, args) = option_parser.parse_args()
    if len(args) != 0:
        option_parser.error("invalid argument")

    products_by_category = load_products_by_category(
        options.csv_filename,
        options.tax_percent / 100.0
    )

    generate_pdf(
        products_by_category,
        options.ncols,
        options.greybar_interval,
        options.pdf_filename
    )


if __name__ == "__main__":
    main()
