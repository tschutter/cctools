#!/usr/bin/env python

"""Detects problems in CoreCommerce product data.

Requires a [website] section in config file for login info.

The [lint_product] section of the config file contains product value
checks.  The key is the product value (column) to check and the value
is a regular expression which the product value must match.  You can
specify multiple checks for a product value by putting them on
multiple lines.  The regex can be followed by a Python "if"
expression.  For example:

    [lint_product]
    Teaser: \S.{9,59}  ; 10-59 chars
    Discontinued Item: Y|N  ; Y or N
    Available:
        Y|N
        N if item["Discontinued Item"] == "Y"
    # UPC must be blank
    UPC:
    Cost: [0-9]+\.[0-9]{2}  ; positive float
    Price: [0-9]+\.[0-9]{2}  ; positive float
    # SKU must be 5 digits.
    # SKU must start with 8 if the category is Cat or Dog.
    SKU:
        [1-9][0-9]{4}
        [8][0-9]{4} if item["Category"] in ("Cat", "Dog")

The [lint_category] section is similar and is used for category value
checks.
"""

import ConfigParser
import cctools
import optparse
import re
import sys

def parse_checks(config, section):
    """Parse value check definitions in config file."""
    checks = list()
    for name, value in config.items(section):
        for check in value.split("\n"):
            check = check.strip()
            if check != "":
                checks.append((name, check))
    return checks


def product_display_name(product):
    """Construct a display name for a product."""
    display_name = "%s %s" % (product["SKU"], product["Product Name"])
    return display_name.strip()


def check_skus(products):
    """Check SKUs for uniqueness."""
    skus = dict()
    for product in products:
        display_name = product_display_name(product)
        sku = product["SKU"]
        if sku != "":
            if sku in skus:
                print "%s '%s': SKU already used by '%s'" % (
                    "Product",
                    display_name,
                    skus[sku]
                )
            else:
                skus[sku] = display_name


def check_item(item_checks, item_type_name, item, item_name):
    """Check category or product for problems."""

    for key, check in item_checks:
        check_parts = check.split(None, 2)
        if len(check_parts) == 1:
            pass
        elif len(check_parts) == 3 and check_parts[1].lower() == "if":
            predicate = check_parts[2]
            if not eval(predicate, {"__builtins__": {}}, {"item": item}):
                continue
        else:
            print "ERROR: Unknown check syntax '%s'" % check
            sys.exit(1)
        if not re.match(check_parts[0], item[key]):
            print "%s '%s': Invalid '%s' of '%s' (does not match %s)" % (
                item_type_name,
                item_name,
                key,
                item[key],
                check_parts[0]
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
        "--cache-ttl",
        action="store",
        type=int,
        metavar="SEC",
        default=3600,
        help="cache TTL in seconds (default=%default)"
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
    config.optionxform = str  # preserve case of option names
    config.readfp(open(options.config))
    category_checks = parse_checks(config, "lint_category")
    product_checks = parse_checks(config, "lint_product")

    # Create a connection to CoreCommerce.
    cc_browser = cctools.CCBrowser(
        config.get("website", "host"),
        config.get("website", "site"),
        config.get("website", "username"),
        config.get("website", "password"),
        verbose=options.verbose,
        clean=options.clean,
        cache_ttl=options.cache_ttl
    )

    # Check category list.
    categories = cc_browser.get_categories()
    for category in categories:
        check_item(
            category_checks,
            "Category",
            category,
            category["Category Name"]
        )

    # Check products list.
    products = cc_browser.get_products()
    check_skus(products)
    for product in products:
        check_item(
            product_checks,
            "Product",
            product,
            product_display_name(product)
        )

    if options.verbose:
        sys.stderr.write("Checks complete\n")
    return 0

if __name__ == "__main__":
    main()
