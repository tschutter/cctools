#!/usr/bin/env python

"""
Generates an Art Mart Inventory Sheet.

sudo apt-get install python-reportlab

Best reportlab reference is the ReportLab User's Guide.
"""

import csv
import datetime
import optparse
import reportlab.lib
import reportlab.platypus

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
        self.setFont("Helvetica", 9)
        label_str = "Page %d of %d" % (self._pageNumber, page_count)
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
    canvas.drawString(6.65 * INCH, page_height - 0.60 * INCH, "Office Use Only")
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

    # Footer.
    canvas.setFont('Helvetica', 9)
    canvas.drawString(
        0.25 * INCH,
        0.50 * INCH,
        "Revised: %s" % today_str
    )

    canvas.restoreState()


def generate_pdf(products, quantities, pdf_filename):
    """Generate a PDF given a list of products by category."""
    # Construct a document.
    doc = reportlab.platypus.BaseDocTemplate(
        pdf_filename,
        pagesize=reportlab.lib.pagesizes.letter,
        title="Art Mart Inventory Sheet Check In/Out",
        #showBoundary=True
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
        #('GRID', (0, 0), (-1, -1), 1.0, BLACK)
    ]
    header_row = (
        reportlab.platypus.Paragraph("Existing<br/>Qty", col_header_style),
        reportlab.platypus.Paragraph("Barcode<br/>&nbsp;", col_center_style),
        reportlab.platypus.Paragraph("Qty<br/>Added", col_header_style),
        reportlab.platypus.Paragraph("Price<br/>&nbsp;", col_align_right_style),
        reportlab.platypus.Paragraph("Description<br/>&nbsp;", col_header_style),
        reportlab.platypus.Paragraph("Total<br/>Qty", col_header_style)
    )
    table_data = [header_row]
    for sku, price, description in products:
        if sku in quantities:
            quantity = quantities[sku]
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


def clean_text(text):
    """Cleanup HTML in text."""
    text = text.replace("&", "&amp;")
    return text


def load_products(csv_filename):
    """Load product data."""

    # Read the input file, extracting just the fields we need.
    is_header = True
    products = list()
    for fields in csv.reader(open(csv_filename)):
        if is_header:
            sku_field = fields.index("SKU")
            price_field = fields.index("Price")
            name_field = fields.index("Product Name")
            teaser_field = fields.index("Teaser")
            discontinued_field = fields.index("Discontinued Item")
            is_header = False
        elif fields[discontinued_field] == "N":
            sku = fields[sku_field]
            price = "$%.2f" % float(fields[price_field])
            description = "%s: %s" % (
                clean_text(fields[name_field]),
                clean_text(fields[teaser_field])
            )
            products.append((sku, price, description))

    # Sort by SKU.
    products = sorted(products, key=lambda x: x[0])

    return products


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
            quantities[fields[sku_field]] = fields[quantity_field]

    return quantities


def main():
    """main"""
    option_parser = optparse.OptionParser(
        usage="usage: %prog [options]\n" +
        "  Generates an Art Mart Inventory Sheet in PDF form."
    )
    option_parser.add_option(
        "--prodfile",
        action="store",
        dest="csv_filename",
        metavar="CSV_FILE",
        default="products.csv",
        help="input product list filename (default=%default)"
    )
    option_parser.add_option(
        "--quantfile",
        action="store",
        dest="quant_filename",
        metavar="CSV_FILE",
        default="quantities.csv",
        help="input product quantities filename (default=%default)"
    )
    option_parser.add_option(
        "--outfile",
        action="store",
        dest="pdf_filename",
        metavar="PDF_FILE",
        default="ArtMartCheckInOut.pdf",
        help="output PDF filename (default=%default)"
    )

    (options, args) = option_parser.parse_args()
    if len(args) != 0:
        option_parser.error("invalid argument")

    products = load_products(options.csv_filename)

    quantities = load_quantities(options.quant_filename)

    generate_pdf(products, quantities, options.pdf_filename)


if __name__ == "__main__":
    main()
