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


def check_product(product):
    """Check product for problems."""
    display_name = "%s %s" % (product["SKU"], product["Product Name"])

    check_string("Product", display_name, product, "Teaser", 10)

    y_n = ("Y", "N")
    check_value_in_set("Product", display_name, product, "Available", y_n)

    check_value_in_set(
        "Product",
        display_name,
        product,
        "Discontinued Item",
        y_n
    )

    if product["Available"] == "Y" and product["Discontinued Item"] == "Y":
        print "Product '%s': Is Available and is a Discontinued Item" % (
            display_name
        )

    if product["UPC"] != "":
        print "Product '%s': UPC '%s' is not blank" % (
            display_name,
            product["UPC"]
        )

    if product["Category"] in ("Necklaces", "Bracelets"):
        if product["MPN"] != "7117.90.9000":
            print "Product '%s': MPN (HTSUS No) '%s' != '7117.90.9000'" % (
                display_name,
                product["MPN"]
            )
    else:
        if product["MPN"] == "":
            print "Product '%s': MPN (HTSUS No) not set" % (display_name)
        elif len(product["MPN"]) != 12:
            print "Product '%s': Invalid MPN (HTSUS No) '%s'" % (
                display_name,
                product["MPN"]
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
        "--no-clean",
        action="store_false",
        dest="clean",
        default=True,
        help="do not clean data before checking"
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
        clean=options.clean
    )

    # Check category list.
    categories = cc_browser.get_categories()
    #for category in categories:
    #    check_category(category)

    ## Build map of category ID to category name.
    #category_names = dict()
    #for category in categories:
    #    category_names[category["Category Id"]] = category["Category Name"]

    # Check products list.
    products = cc_browser.get_products()
    for product in products:
        check_product(product)

    if options.verbose:
        sys.stderr.write("Checks complete\n")
    return 0

if __name__ == "__main__":
    main()
