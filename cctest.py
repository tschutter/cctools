#!/usr/bin/env python

"""
Test cctools.
"""

import ConfigParser
import cctools
import optparse
import os
import sys

def table_sort(table, field):
    return table


def table_print_divider(col_widths):
    print(
        "+" + "+".join(
            "-" * (col_width + 2)
            for col_width in col_widths
        ) + "+"
    )


def table_print_row(row, col_widths):
    print(
        "| " + " | ".join(
            "%-*s" % (col_widths[col], field)
            for col, field in enumerate(row)
        ) + " |"
    )


def table_print(table, fields):
    # Determine the column widths.
    col_widths = list()
    for field in fields:
        col_widths.append(len(field))
    for record in table:
        for col, field in enumerate(fields):
            col_widths[col] = max(col_widths[col], len(str(record[field])))

    # Pretty print the records to the output.
    table_print_divider(col_widths)
    table_print_row(fields, col_widths)
    table_print_divider(col_widths)
    for row, record in enumerate(table):
        table_print_row(
            [cctools.html_to_plain_text(record[field]) for field in fields],
            col_widths
        )
    table_print_divider(col_widths)


def main():
    """main"""
    default_config = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "cctools.cfg"
    )

    option_parser = optparse.OptionParser(
        usage="usage: %prog [options] action\n" +
        "  Actions:\n" +
        "    products - list products\n"
        "    categories - list categories\n"
        "    personalizations - list personalizations"
    )
    option_parser.add_option(
        "--config",
        action="store",
        metavar="FILE",
        default=default_config,
        help="configuration filename (default=%default)"
    )
    option_parser.add_option(
        "--cache-ttl",
        action="store",
        metavar="SEC",
        default=3600,
        help="cache TTL (default=%default)"
    )
    option_parser.add_option(
        "--verbose",
        action="store_true",
        default=False,
        help="display progress messages"
    )

    # Parse command line arguments.
    (options, args) = option_parser.parse_args()
    if len(args) != 1:
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
        cache_ttl=options.cache_ttl,
        verbose=options.verbose,
        #proxy="localhost:8080"
    )

    action = args[0]
    if (action == "products"):
        products = cc_browser.get_products()
        table_print(products, ["SKU", "Product Name", "Price", "Available", "Track Inventory", "Inventory Level", "Teaser"])
    elif (action == "categories"):
        categories = cc_browser.get_categories()
        categories = table_sort(categories, "Sort")
        table_print(categories, ["Category Id", "Category Name", "Sort"])
    elif (action == "personalizations"):
        personalizations = cc_browser.get_personalizations()
        personalizations = sorted(
            personalizations,
            key=cc_browser.personalization_sort_key
        )
        table_print(
            personalizations,
            [
                "Product SKU",
                "Product Name",
                "Question|Answer",
                "SKU",
                "Answer Sort Order",
                "Inventory Level",
                "Answer Input Type",
                "Main Photo"
            ]
        )
    else:
        option_parser.error("invalid action")

    return 0

if __name__ == "__main__":
    main()
