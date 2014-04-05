#!/usr/bin/env python2

"""
Calculates an online price (pre-tax) based upon a tax-included price.
"""

from __future__ import print_function
import math
import optparse

def calc_pre_tax_price(tax_included_price, price_multiplier):
    """Calculate the pre-tax price."""
    pre_tax_price = tax_included_price / price_multiplier
    pre_tax_price = math.floor(pre_tax_price * 10.0 + 0.5) / 10.0 - 0.01
    return "{:,.2f}".format(pre_tax_price)

def main():
    """main"""
    option_parser = optparse.OptionParser(
        usage="usage: %prog [options] tax-included-price\n" +
        "  Calculate an pre-tax price based upon a tax-included price."
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

    # Parse command line arguments.
    (options, args) = option_parser.parse_args()
    if len(args) == 0:
        option_parser.error("price not specified")
    elif len(args) > 1:
        option_parser.error("invalid argument")
    else:
        tax_included_price = float(args[0])

    # Determine price multiplier.
    price_multiplier = 1.0 + options.tax_percent / 100.0

    print(calc_pre_tax_price(tax_included_price, price_multiplier))

    return 0

if __name__ == "__main__":
    main()
