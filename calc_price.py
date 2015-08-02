#!/usr/bin/env python

"""
Converts an event price (tax-included) to a retail price (pre-tax)
or vice versa.

Event prices are calculated by applying a discount to the online
price, adding sales tax, and rounding to the nearest dollar.

Retail prices are calculated by removing sales tax, inverting the
discount, rounding to the nearest dime, and subtracting a penny.

To generate a table of event to retail prices::

  echo "event,retail";\
  i=1; while [ "$i" -le 10 ]; do\
    echo "$i,`./calc_price.py retail $i`";\
    i=`echo $i+1|bc`;\
  done
"""

from __future__ import print_function
import argparse
import math
import Tkinter
import tkMessageBox


def calc_event_price(price, discount_percent, sales_tax_percent):
    """Calculate the event price from the retail price."""

    # Apply discount.
    price = float(price) * (1.0 - float(discount_percent) / 100.0)

    # Impose sales tax.
    price = price * (1.0 + float(sales_tax_percent) / 100.0)

    # Round to the nearest dollar.
    price = max(1.0, math.floor(price + 0.5))

    # Return as a float.
    return price


def calc_retail_price(price, discount_percent, sales_tax_percent):
    """Calculate the retail price from the event price."""

    # Remove sales tax.
    price = float(price) / (1.0 + float(sales_tax_percent) / 100.0)

    # Unapply discount.
    price = price / (1.0 - float(discount_percent) / 100.0)

    # Round to the nearest dime and subtract a penny.
    price = max(0.01, math.floor(price * 10.0 + 0.5) / 10.0 - 0.01)

    # Return as a float.
    return price


class AppGUI(object):
    """Application graphical user interface."""

    # Width of price text entry boxes.
    PRICE_WIDTH = 10

    # pylint: disable=no-self-use
    def __init__(self, args, root):
        self.args = args
        self.root = root
        self.frame = Tkinter.Frame(root)
        self.frame.pack()

        # Retail price group.
        retail_group = Tkinter.LabelFrame(
            self.frame,
            text="Retail (website) price",
            padx=5,
            pady=5
        )
        retail_group.pack(padx=10, pady=10)
        self.retail_entry = Tkinter.Entry(
            retail_group,
            width=AppGUI.PRICE_WIDTH
        )
        self.retail_entry.bind('<Return>', self.calc_event_price)
        self.retail_entry.pack(side=Tkinter.LEFT)
        calc_event_button = Tkinter.Button(
            retail_group,
            text="Calc event",
            command=self.calc_event_price
        )
        calc_event_button.bind('<Return>', self.calc_event_price)
        calc_event_button.pack(side=Tkinter.LEFT)

        # Event price group.
        event_group = Tkinter.LabelFrame(
            self.frame,
            text="Event (discounted) price",
            padx=5,
            pady=5
        )
        event_group.pack(padx=10, pady=10)
        self.event_entry = Tkinter.Entry(
            event_group,
            width=AppGUI.PRICE_WIDTH
        )
        self.event_entry.bind('<Return>', self.calc_retail_price)
        self.event_entry.pack(side=Tkinter.LEFT)
        calc_retail_button = Tkinter.Button(
            event_group,
            text="Calc retail",
            command=self.calc_retail_price
        )
        calc_retail_button.bind('<Return>', self.calc_retail_price)
        calc_retail_button.pack(side=Tkinter.LEFT)

        # Quit button
        self.quit_button = Tkinter.Button(
            self.frame,
            text="Quit",
            command=self.frame.quit
        )
        self.quit_button.pack(side=Tkinter.LEFT, padx=5, pady=5)

        # Give focus to the retail_entry.
        self.retail_entry.focus_set()

    def get_price(self, entry, entry_name):
        """Get float value from text entry box."""
        price_str = entry.get()
        if price_str is None:
            self.error("No {} price specified".format(entry_name))
            return None
        try:
            price_float = float(price_str)
        except ValueError:
            self.error(
                "{} price of '{}' is not a valid number".format(
                    entry_name,
                    price_str
                )
            )
            return None
        return price_float

    def set_price(self, entry, price):
        """Set price in text entry box."""
        entry.delete(0, Tkinter.END)
        entry.insert(0, price)

        # Copy to clipboard.
        self.root.clipboard_clear()
        self.root.clipboard_append(price)

    def calc_retail_price(self, _=None):
        """
        Calculate retail price action.
        Second arg is event when called via <Return>.
        """
        event_price = self.get_price(self.event_entry, "Event")
        if event_price is None:
            return
        retail_price = calc_retail_price(
            event_price,
            self.args.discount_percent,
            self.args.avg_tax_percent
        )
        retail_price = "{:,.2f}".format(retail_price)
        self.set_price(self.retail_entry, retail_price)

    def calc_event_price(self, _=None):
        """
        Calculate event price action.
        Second arg is event when called via <Return>.
        """
        retail_price = self.get_price(self.retail_entry, "Retail")
        if retail_price is None:
            return
        event_price = calc_event_price(
            retail_price,
            self.args.discount_percent,
            self.args.avg_tax_percent
        )
        event_price = "{:,.2f}".format(event_price)
        self.set_price(self.event_entry, event_price)

    def error(self, msg):
        """Display an error message."""
        tkMessageBox.showerror("Error", msg)


def main():
    """main"""
    arg_parser = argparse.ArgumentParser(
        description="Calculate pre-tax price based upon tax-included price "
        "or vice-versa."
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

    subparsers = arg_parser.add_subparsers(help='operations')

    subparsers.add_parser("gui", help="display graphical interface")

    event_parser = subparsers.add_parser(
        "event",
        help="calculate event price"
    )
    event_parser.add_argument(
        "retail_price",
        type=float,
        help="price"
    )

    retail_parser = subparsers.add_parser(
        "retail",
        help="calculate retail price"
    )
    retail_parser.add_argument(
        "event_price",
        type=float,
        help="price"
    )

    # Parse command line arguments.
    args = arg_parser.parse_args()

    # Convert and print price.
    if "retail_price" in args:
        price = calc_event_price(
            args.retail_price,
            args.discount_percent,
            args.avg_tax_percent
        )
        print("{:,.0f}".format(price))
    elif "event_price" in args:
        price = calc_retail_price(
            args.event_price,
            args.discount_percent,
            args.avg_tax_percent
        )
        print("{:,.2f}".format(price))
    else:
        root = Tkinter.Tk()
        root.title("Calc Product Price")
        AppGUI(args, root)
        root.mainloop()

    return 0

if __name__ == "__main__":
    main()
