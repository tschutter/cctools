#!/usr/bin/env python2

"""
Calculates an online price (pre-tax) based upon a tax-included price.
"""

from __future__ import print_function
import argparse
import math


def calc_pre_tax_price(tax_included_price, price_multiplier):
    """Calculate the pre-tax price."""
    pre_tax_price = tax_included_price / price_multiplier
    pre_tax_price = math.floor(pre_tax_price * 10.0 + 0.5) / 10.0 - 0.01
    return "{:,.2f}".format(pre_tax_price)


def main():
    """main"""
    arg_parser = argparse.ArgumentParser(
        description="Calculate pre-tax price based upon tax-included price."
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
        "tax_included_price",
        metavar="PRICE",
        type=float,
        help="price including tax"
    )

    # Parse command line arguments.
    args = arg_parser.parse_args()

    # Determine price multiplier.
    price_multiplier = 1.0 + args.tax_percent / 100.0

    # Calculate and print the pre-tax price.
    print(calc_pre_tax_price(args.tax_included_price, price_multiplier))

    return 0

if __name__ == "__main__":
    main()
