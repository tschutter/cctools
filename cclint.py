#!/usr/bin/env python

"""
Detects problems in data exported from CoreCommerce.
"""

import ConfigParser
import cctools
import optparse
import sys

def check_string(type_name, item_name, item, key, min_len):
    """Print message if item[key] is empty or shorter than min_len."""
    value_len = len(item[key])
    if value_len == 0:
        print "%s '%s': Value '%s' not defined" % (
            type_name,
            item_name,
            key
        )
    elif value_len < min_len:
        print "%s '%s': Value '%s' == '%s' is too short" % (
            type_name,
            item_name,
            key,
            item[key]
        )


def check_value_in_set(type_name, item_name, item, key, valid_values):
    """Print message if item[key] not in valid_values."""
    if not item[key] in valid_values:
        print "%s '%s': Invalid '%s' == '%s' not in %s" % (
            type_name,
            item_name,
            key,
            item[key],
            valid_values
        )


def check_products(products):
    """Check products list for problems."""
    for product in products:
        display_name = "%s %s" % (product["SKU"], product["Product Name"])

        check_string(
            "Product",
            display_name,
            product,
            "Teaser",
            10
        )

        check_value_in_set(
            "Product",
            display_name,
            product,
            "Discontinued Item",
            ("Y", "N")
        )


def main():
    """main"""
    option_parser = optparse.OptionParser(
        usage="usage: %prog [options]\n" +
        "  Detects problems in data exported from CoreCommerce."
    )
    option_parser.add_option(
        "--config",
        action="store",
        metavar="FILE",
        default="cctools.cfg",
        help="configuration filename (default=%default)"
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
        config.get("website", "password"),
        verbose=options.verbose,
        clean=False  # don't cleanup data
    )

    # Check products list.
    products = cc_browser.get_products()
    check_products(products)

    if options.verbose:
        sys.stderr.write("Checks complete\n")
    return 0

if __name__ == "__main__":
    main()
