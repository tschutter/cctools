#!/usr/bin/env python2

"""
Generates an Art Mart Inventory Sheet.
"""

import ConfigParser
import argparse
import cctools
import csv
import datetime
import logging
import math
import notify_send_handler
import os
import reportlab.lib  # sudo apt-get install python-reportlab
import reportlab.platypus

# Best reportlab reference is the ReportLab User's Guide.

# Convenience constants.
INCH = reportlab.lib.units.inch
BLACK = reportlab.lib.colors.black


class NumberedCanvas(reportlab.pdfgen.canvas.Canvas):
    """http://code.activestate.com/recipes/576832/"""
    def __init__(self, *args, **kwargs):
        reportlab.pdfgen.canvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        """add page info to each page (page x of y)"""
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            reportlab.pdfgen.canvas.Canvas.showPage(self)
        reportlab.pdfgen.canvas.Canvas.save(self)

    def draw_page_number(self, page_count):
        """Add 'Page x of y' to canvas."""
        self.setFont("Helvetica", 9)
        label_str = "Page {} of {}".format(self._pageNumber, page_count)
        self.drawRightString(8.25 * INCH, 0.50 * INCH, label_str)


def on_page(canvas, doc):
    """Add page header and footer.  Called for each page."""
    today_str = datetime.date.today().isoformat()
    page_width = doc.pagesize[0]
    page_height = doc.pagesize[1]
    canvas.saveState()

    # Title.
    canvas.setFont('Helvetica-Bold', 18)
    canvas.drawCentredString(
        3.65 * INCH,
        page_height - 0.75 * INCH,
        "Art Mart Inventory Sheet"
    )
    canvas.drawCentredString(
        3.65 * INCH,
        page_height - 1.0 * INCH,
        "Check In/Out"
    )

    # Upper-right block.
    canvas.setFont('Helvetica', 10)
    canvas.drawString(
        6.65 * INCH,
        page_height - 0.60 * INCH,
        "Office Use Only"
    )
    canvas.setFont('Helvetica', 9)
    canvas.drawRightString(7.15 * INCH, page_height - 0.80 * INCH, "Recd. Via")
    canvas.drawRightString(7.15 * INCH, page_height - 1.00 * INCH, "Date Recd")
    canvas.drawString(7.25 * INCH, page_height - 0.80 * INCH, "_________")
    canvas.drawString(7.25 * INCH, page_height - 1.00 * INCH, "_________")

    # Second block.
    canvas.setFont('Helvetica-Bold', 12)
    canvas.drawRightString(1.65 * INCH, page_height - 1.45 * INCH, "Artist #")
    canvas.drawRightString(1.65 * INCH, page_height - 1.75 * INCH, "Store")
    canvas.drawRightString(3.15 * INCH, page_height - 1.45 * INCH, "Artist")
    canvas.drawRightString(3.15 * INCH, page_height - 1.75 * INCH, "Email")
    canvas.drawRightString(6.75 * INCH, page_height - 1.45 * INCH, "Date")
    canvas.drawRightString(6.75 * INCH, page_height - 1.75 * INCH, "Phone")
    canvas.setFont('Helvetica', 12)
    email = "linda@circleofhandsuganda.com"
    canvas.drawString(1.75 * INCH, page_height - 1.45 * INCH, "2587")
    canvas.drawString(1.75 * INCH, page_height - 1.75 * INCH, "Boulder")
    canvas.drawString(3.25 * INCH, page_height - 1.45 * INCH, "Linda Schutter")
    canvas.drawString(3.25 * INCH, page_height - 1.75 * INCH, email)
    canvas.drawString(6.85 * INCH, page_height - 1.45 * INCH, today_str)
    canvas.drawString(6.85 * INCH, page_height - 1.75 * INCH, "720.318.4099")

    # Bottom block.
    canvas.setFont('Helvetica', 9)
    canvas.drawCentredString(
        page_width / 2.0,
        1.35 * INCH,
        "All checkin and out sheets should be signed below after being" +
        " verified by the Artist and an Art Mart Employee"
    )
    canvas.drawRightString(1.45 * INCH, 0.98 * INCH, "Signed")
    sig_line = "____________________________________________"
    canvas.drawString(1.55 * INCH, 0.98 * INCH, sig_line)
    canvas.drawString(4.80 * INCH, 0.98 * INCH, sig_line)
    canvas.drawCentredString(3.00 * INCH, 0.80 * INCH, "Art Mart Employee")
    canvas.drawCentredString(
        6.10 * INCH,
        0.80 * INCH,
        "Artist or Representative"
    )

    canvas.restoreState()


def generate_pdf(products, quantities, pdf_filename):
    """Generate the PDF file."""
    # Construct a document.
    doc = reportlab.platypus.BaseDocTemplate(
        pdf_filename,
        pagesize=reportlab.lib.pagesizes.letter,
        title="Art Mart Inventory Sheet Check In/Out",
        # showBoundary=True
    )

    # Construct a frame.
    page_width = doc.pagesize[0]
    page_height = doc.pagesize[1]
    frames = [
        reportlab.platypus.Frame(
            id="table_frame",
            x1=0.25 * INCH,
            y1=1.60 * INCH,
            width=page_width - 0.45 * INCH,
            height=page_height - 3.60 * INCH,
            leftPadding=0,
            bottomPadding=0,
            rightPadding=0,
            topPadding=0
        )
    ]

    # Construct a template and add it to the document.
    doc.addPageTemplates(
        reportlab.platypus.PageTemplate(
            id="mytemplate",
            frames=frames,
            onPage=on_page
        )
    )

    # Construct a story and add it to the document.
    col_widths = [
        0.40 * INCH,  # Existing Qty
        1.15 * INCH,  # Barcode
        0.40 * INCH,  # Qty Added
        0.60 * INCH,  # Price
        5.00 * INCH,  # Description
        0.40 * INCH,  # Total Qty
    ]
    col_header_style = reportlab.lib.styles.ParagraphStyle(
        name="col_header",
        fontName="Helvetica-Bold",
        fontSize=10
    )
    col_center_style = reportlab.lib.styles.ParagraphStyle(
        name="col_center",
        fontName="Helvetica-Bold",
        fontSize=10,
        alignment=reportlab.lib.enums.TA_CENTER
    )
    col_align_right_style = reportlab.lib.styles.ParagraphStyle(
        name="col_align_right",
        fontName="Helvetica-Bold",
        fontSize=10,
        alignment=reportlab.lib.enums.TA_RIGHT
    )
    # TableStyle cell formatting commands.
    styles = [
        # Whole table.
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 1), (-1, -1), 0.26 * INCH),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 0.10 * INCH),
        # Existing Qty
        ('LEFTPADDING', (0, 0), (0, -1), 0),
        # Barcode
        ('ALIGN', (1, 0), (1, 0), 'CENTER'),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('FONT', (1, 0), (1, -1), 'Courier-Bold'),
        # Qty Added
        ('LEFTPADDING', (2, 0), (2, 0), 0),
        ('ALIGN', (2, 0), (2, 0), 'LEFT'),
        ('ALIGN', (2, 1), (2, -1), 'RIGHT'),
        ('FONT', (2, 0), (2, -1), 'Courier-Bold'),
        # Price
        ('ALIGN', (3, 0), (3, -1), 'RIGHT'),
        ('FONT', (3, 0), (3, -1), 'Courier-Bold'),
        # Description
        ('ALIGN', (4, 0), (4, -1), 'LEFT'),
        ('FONT', (4, 0), (4, -1), 'Helvetica'),
        # Total Qty
        ('LEFTPADDING', (5, 0), (5, -1), 0),
        ('ALIGN', (5, 0), (5, 0), 'LEFT'),
        # Other
        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold'),
        # ('GRID', (0, 0), (-1, -1), 1.0, BLACK)
    ]
    header_row = (
        reportlab.platypus.Paragraph("Existing<br/>Qty", col_header_style),
        reportlab.platypus.Paragraph("Barcode<br/>&nbsp;", col_center_style),
        reportlab.platypus.Paragraph("Qty<br/>Added", col_header_style),
        reportlab.platypus.Paragraph(
            "Price<br/>&nbsp;",
            col_align_right_style
        ),
        reportlab.platypus.Paragraph(
            "Description<br/>&nbsp;",
            col_header_style
        ),
        reportlab.platypus.Paragraph("Total<br/>Qty", col_header_style)
    )
    table_data = [header_row]
    for product in products:
        sku = product["SKU"]
        if sku in quantities:
            quantity = quantities[sku]
            # Round price to nearest dollar.
            price = product["Price"]
            price = "${.0f}".format(math.trunc(float(price) + 0.5))
            description = "{}: {}".format(
                product["Product Name"],
                cctools.html_to_plain_text(product["Teaser"])
            )
            description = description[:68]
            table_data.append(
                ("_____", sku, quantity, price, description, "_____")
            )
    table = reportlab.platypus.Table(
        data=table_data,
        colWidths=col_widths,
        style=styles,
        repeatRows=1
    )
    table.hAlign = "LEFT"
    story = [table]
    doc.build(story, canvasmaker=NumberedCanvas)


def load_quantities(quant_filename):
    """Load quantities data."""

    # Read the input file, extracting just the fields we need.
    is_header = True
    quantities = dict()
    for fields in csv.reader(open(quant_filename)):
        if is_header:
            sku_field = fields.index("SKU")
            quantity_field = fields.index("Quantity")
            is_header = False
        else:
            quantity = int(fields[quantity_field])
            if quantity != 0:
                quantities[fields[sku_field]] = str(quantity)

    return quantities


def write_quantities(quant_filename, products):
    """Write empty quantities file."""

    with open(quant_filename, "w") as quant_file:
        quant_file.write("Quantity,SKU,Description(ignored)\n")
        for product in products:
            sku = product["SKU"]
            description = "{}: {}".format(
                product["Product Name"],
                cctools.html_to_plain_text(product["Teaser"])
            )
            quant_file.write(",".join(["0", sku, '"%s"' % description]) + "\n")


def main():
    """main"""
    default_config = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "cctools.cfg"
    )

    today = datetime.date.today()
    default_pdf_filename = "{:4}-{:02}-{:02}-ArtMartCheckInOut.pdf".format(
        today.year,
        today.month,
        today.day
    )

    arg_parser = argparse.ArgumentParser(
        description="Generates an Art Mart Inventory Sheet in PDF form."
    )
    arg_parser.add_argument(
        "--config",
        dest="config",
        metavar="FILE",
        default=default_config,
        help="configuration filename (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--quantfile",
        dest="quant_filename",
        metavar="CSV_FILE",
        default="ArtMartQuantities.csv",
        help="input product quantities filename (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--outfile",
        dest="pdf_filename",
        metavar="PDF_FILE",
        default=default_pdf_filename,
        help="output PDF filename (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--write-quant",
        action="store_true",
        dest="write_quant",
        default=False,
        help="write template quantity file instead of PDF"
    )
    arg_parser.add_argument(
        "--verbose",
        action="store_true",
        default=False,
        help="display progress messages"
    )

    # Parse command line arguments.
    args = arg_parser.parse_args()

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
    config = ConfigParser.RawConfigParser()
    config.readfp(open(args.config))

    # Create a connection to CoreCommerce.
    cc_browser = cctools.CCBrowser(
        config.get("website", "base_url"),
        config.get("website", "username"),
        config.get("website", "password")
    )

    # Fetch products list.
    products = cc_browser.get_products()

    # Sort products by category, product_name.
    products = sorted(
        products,
        key=cc_browser.product_key_by_cat_and_name
    )

    if args.write_quant:
        logger.debug("Generating {}".format(args.quant_filename))
        write_quantities(args.quant_filename, products)

    else:
        quantities = load_quantities(args.quant_filename)
        pdf_filename = args.pdf_filename
        logger.debug("Generating {}".format(pdf_filename))
        generate_pdf(products, quantities, pdf_filename)

    logger.debug("Generation complete")
    return 0

if __name__ == "__main__":
    main()
