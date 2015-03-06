#!/usr/bin/env python

"""
Converts an event price to a retail price or vice versa.

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


def main():
    """main"""
    arg_parser = argparse.ArgumentParser(
        description="Calculate pre-tax price based upon tax-included price."
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
        "operation",
        choices=("event", "retail"),
        default="event",
        help="calculate event price or retail price"
    )
    arg_parser.add_argument(
        "price",
        metavar="PRICE",
        type=float,
        help="price"
    )

    # Parse command line arguments.
    args = arg_parser.parse_args()

    # Convert and print price.
    if args.operation == "event":
        price = calc_event_price(
            args.price,
            args.discount_percent,
            args.avg_tax_percent
        )
        print("{:,.0f}".format(price))
    else:
        price = calc_retail_price(
            args.price,
            args.discount_percent,
            args.avg_tax_percent
        )
        print("{:,.2f}".format(price))

    return 0

if __name__ == "__main__":
    main()
