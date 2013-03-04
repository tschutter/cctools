#!/usr/bin/env python

"""
Test cctools.
"""

import ConfigParser
import cctools
import optparse
import os
import sys

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
        default=60,
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
        print cc_browser.get_products()
    elif (action == "categories"):
        print cc_browser.get_categories()
    elif (action == "personalizations"):
        print cc_browser.get_personalizations()
    else:
        option_parser.error("invalid action")

    return 0

if __name__ == "__main__":
    main()
