#!/usr/bin/env python2

r"""Detects problems in CoreCommerce product data.

Requires a [website] section in config file for login info.

The rules file (default=cclint.rules) contains product value checks in
JSON format.  The schema is:

    {
        constants: {
            CONST_NAME1: CONST_VALUE1,
            CONST_NAME2: [ "CV2_PART1", "CV2_PART2", ... ],
            ...
        },
        rules: [
            {
                "id": "unique ID for rule",
                "disabled": "message why rule disabled (key optional)",
                "itemtype": "category|product|variant",
                "test": "Python predicate, True means test passed",
                "message": "message output if test returned False"
            },
            ...
        ]
    }

The message is formatted using message.format(current_item), so it can
contain member variable references like "{SKU}".

TODO
----

* HTSUS No in message is not working
* change rules from array to dict; make id the key
* change error from string to a tuple
* display errors in a GUI table
* specify SKU uniqueness check as a rule
"""

from __future__ import print_function
import ConfigParser
import argparse
import cctools
import json
import logging
import notify_send_handler
import os
import re
import sys


def load_constants(logger, constants, eval_locals):
    """Load constants from rules file into eval_locals."""
    for varname in constants:
        value = constants[varname]
        if isinstance(value, list):
            value = "\n".join(value)
        statement = "{} = {}".format(varname, value)
        # pylint: disable=W0122
        exec(statement, {}, eval_locals)


def validate_rules(logger, rules):
    """Validate rules read from rules file."""
    rule_ids = []
    failed = False
    for ruleno, rule in enumerate(rules):
        rule_id = "#{}".format(ruleno + 1)
        if "id" not in rule:
            logger.error(
                "Rule {} does not specify an id.".format(rule_id)
            )
            failed = True
        elif len(rule["id"].strip()) == 0:
            logger.error(
                "ERROR: Rule {} has an invalid id of '{}'.".format(
                    rule_id,
                    rule["id"]
                )
            )
            failed = True
        elif rule["id"] in rule_ids:
            logger.error(
                "ERROR: Rule {} has a duplicate id of '{}'.".format(
                    rule_id,
                    rule["id"]
                )
            )
            failed = True
        else:
            rule_id = rule["id"].strip()
            rule_ids.append(rule_id)

        if "itemtype" not in rule:
            logger.error(
                "ERROR: Rule {} does not specify an itemtype.".format(rule_id)
            )
            failed = True
        elif rule["itemtype"] not in ["category", "product", "variant"]:
            logger.error(
                "ERROR: Rule {} has an invalid itemtype.".format(rule_id)
            )
            failed = True

        if "test" not in rule:
            logger.error(
                "ERROR: Rule {} does not specify a test.".format(rule_id)
            )
            failed = True
        else:
            if isinstance(rule["test"], list):
                rule["test"] = " ".join(rule["test"])

        if "message" not in rule:
            logger.error(
                "ERROR: Rule {} does not specify a message.".format(rule_id)
            )
            failed = True

    if failed:
        sys.exit(1)


def product_display_name(product):
    """Construct a display name for a product."""
    display_name = "{} {}".format(product["SKU"], product["Product Name"])
    return display_name.strip()


def variant_display_name(variant):
    """Construct a display name for a variant."""
    display_name = "{} {} {}".format(
        variant["Product SKU"],
        variant["Product Name"],
        variant["Question|Answer"]
    )
    return display_name.strip()


def check_skus(products):
    """Check SKUs for uniqueness."""
    errors = []
    skus = dict()
    for product in products:
        display_name = product_display_name(product)
        sku = product["SKU"]
        if sku != "":
            if sku in skus:
                errors.append(
                    "{} '{}': SKU already used by '{}'".format(
                        "Product",
                        display_name,
                        skus[sku]
                    )
                )
            else:
                skus[sku] = display_name
    return errors


def check_item(logger, eval_locals, rules, item, item_name):
    """Check item for problems."""

    errors = []
    eval_globals = {"__builtins__": {"len": len, "re": re}}
    eval_locals["item"] = item

    for rule in rules:
        try:
            success = eval(rule["test"], eval_globals, eval_locals)
        except SyntaxError as ex:
            # SyntaxError probably means that the test is a statement,
            # not an expression.
            logger.fatal("Syntax error in {} rule:\n{}".format(rule["id"], rule["test"]))
            sys.exit(1)
        if not success:
            try:
                # pylint: disable=W0142
                message = rule["message"].format(**item)
            except Exception:
                message = rule["message"]
            errors.append("{} '{}': {}".format(rule["id"], item_name, message))

    return errors


def run_checks(logger, eval_locals, cc_browser, rules):
    """Run all checks, returning a list of errors."""

    errors = []

    # Check category list.
    categories = cc_browser.get_categories()
    eval_locals["items"] = categories
    category_rules = [rule for rule in rules if rule["itemtype"] == "category"]
    for category in categories:
        errors.extend(
            check_item(
                logger,
                eval_locals,
                category_rules,
                category,
                category["Category Name"]
            )
        )

    # Check products list.
    products = cc_browser.get_products()
    eval_locals["items"] = products
    product_rules = [rule for rule in rules if rule["itemtype"] == "product"]
    errors.extend(check_skus(products))
    for product in products:
        for key in ["Teaser"]:
            product[key] = cctools.html_to_plain_text(product[key])
        errors.extend(
            check_item(
                logger,
                eval_locals,
                product_rules,
                product,
                product_display_name(product)
            )
        )

    # Check variants list.
    variants = cc_browser.get_variants()
    eval_locals["items"] = variants
    variant_rules = [rule for rule in rules if rule["itemtype"] == "variant"]
    for variant in variants:
        errors.extend(
            check_item(
                logger,
                eval_locals,
                variant_rules,
                variant,
                variant_display_name(variant)
            )
        )

    return errors


def main():
    """main"""
    default_config = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "cctools.cfg"
    )
    default_rules = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "cclint.rules"
    )

    arg_parser = argparse.ArgumentParser(
        description="Detects problems in data exported from CoreCommerce."
    )
    arg_parser.add_argument(
        "--config",
        metavar="FILE",
        default=default_config,
        help="configuration filename (default=%(default)s)"
    )
    arg_parser.add_argument(
        "--rules",
        action="append",
        metavar="FILE",
        help="lint rules filename (default={})".format(default_rules)
    )
    arg_parser.add_argument(
        "--no-clean",
        action="store_false",
        dest="clean",
        default=True,
        help="do not clean data before checking"
    )
    arg_parser.add_argument(
        "--refresh-cache",
        action="store_true",
        default=False,
        help="refresh cache from website"
    )
    arg_parser.add_argument(
        "--cache-ttl",
        type=int,
        metavar="SEC",
        default=3600,
        help="cache TTL in seconds (default=%(default)i)"
    )
    arg_parser.add_argument(
        "--rule-ids",
        metavar="ID1,ID2,...",
        help="rules to check (default=all)"
    )
    arg_parser.add_argument(
        "--verbose",
        action="store_true",
        default=False,
        help="display progress messages"
    )

    # Parse command line arguments.
    args = arg_parser.parse_args()
    if args.rules is None:
        args.rules = [default_rules]

    # Configure logging.
    logging.basicConfig(
        level=logging.INFO if args.verbose else logging.WARNING
    )
    logger = logging.getLogger()

    # Also log using notify-send if it is available.
    if notify_send_handler.NotifySendHandler.is_available():
        logger.addHandler(
            notify_send_handler.NotifySendHandler(
                os.path.splitext(os.path.basename(__file__))[0]
            )
        )

    # Read config file.
    config = ConfigParser.RawConfigParser()
    config.optionxform = str  # preserve case of option names
    config.readfp(open(args.config))

    # Read rules files.
    eval_locals = {}
    rules = []
    for rulesfile in args.rules:
        try:
            constants_and_rules = json.load(open(rulesfile))
        except Exception as ex:
            print("Error loading rules file {}:".format(rulesfile))
            print("  {}".format(ex.message))
            return 1

        #print(json.dumps(rules, indent=4, sort_keys=True))
        if "constants" in constants_and_rules:
            load_constants(
                logger,
                constants_and_rules["constants"],
                eval_locals
            )
        if "rules" in constants_and_rules:
            file_rules = [
                rule
                for rule in constants_and_rules["rules"]
                if "disabled" not in rule
            ]
            validate_rules(logger, file_rules)
            if args.rule_ids is not None:
                rule_ids = args.rule_ids.split(",")
                file_rules = [
                    rule
                    for rule in file_rules
                    if rule["id"] in args.rule_ids
                ]
            rules.extend(file_rules)

    # Create a connection to CoreCommerce.
    cc_browser = cctools.CCBrowser(
        config.get("website", "host"),
        config.get("website", "site"),
        config.get("website", "username"),
        config.get("website", "password"),
        clean=args.clean,
        cache_ttl=0 if args.refresh_cache else args.cache_ttl
    )

    # Run checks.
    errors = run_checks(logger, eval_locals, cc_browser, rules)

    # Display errors.
    for error in errors:
        print(error)

    logger.info("Checks complete")
    return 0


if __name__ == "__main__":
    main()
