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
            "uniqueid": {
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

* change rule file format from JSON to YAML
  * http://wikipedia.org/wiki/YAML
  * http://pyyaml.org/wiki/PyYAMLDocumentation
* fix width determination of item column
* specify SKU uniqueness check as a rule
"""

from __future__ import print_function
import ConfigParser
import Tkinter
import argparse
import cctools
import json
import os
import re
import sys
import tkFont
import tkMessageBox
import ttk
import webbrowser


def dupe_checking_hook(pairs):
    """
    An object_pairs_hook for json.load() that raises a KeyError on
    duplicate or blank keys.
    """
    result = dict()
    for key, val in pairs:
        if key.strip() == "":
            raise KeyError("Blank key specified")
        if key in result:
            raise KeyError("Duplicate key specified: %s" % key)
        result[key] = val
    return result


def load_constants(constants, eval_locals):
    """Load constants from rules file into eval_locals."""
    for varname in constants:
        value = constants[varname]
        if isinstance(value, list):
            value = "\n".join(value)
        statement = "{} = {}".format(varname, value)
        # pylint: disable=W0122
        exec(statement, {}, eval_locals)


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


def item_edit_url(config, itemtype, item):
    """Create an URL for editing an item."""
    base_url = "https://{}/~{}/admin/index.php".format(
        config.get("website", "host"),
        config.get("website", "site")
    )

    if itemtype == "category":
        url = "{}?m=edit_category&catID={}".format(
            base_url,
            item["Category Id"]
        )

    elif itemtype == "product":
        # CoreCommerce does not report the Product Id; it is guessed
        # by cctools.  have a pID.
        if "Product Id" in item and item["Product Id"] != "":
            url = "{}?m=edit_product&pID={}".format(
                base_url,
                item["Product Id"]
            )
        elif len(item["SKU"]) > 0:
            url = "{}?m=products_browse&search={}".format(
                base_url,
                item["SKU"]
            )
        else:
            url = "{}?m=products_browse&search={}".format(
                base_url,
                item["Product Name"]
            )

    elif itemtype == "variant":
        question_id = item["Question ID|Answer ID"].split("|")[0]
        url = "{}?m=edit_product_personalizations&pID={}&persId={}".format(
            base_url,
            item["Product Id"],
            question_id
        )
    else:
        url = base_url
    return url


def check_skus(config, products):
    """Check SKUs for uniqueness."""
    errors = []
    skus = dict()
    for product in products:
        display_name = product_display_name(product)
        sku = product["SKU"]
        if sku != "":
            if sku in skus:
                url = item_edit_url(config, "product", product)
                error = (
                    "product",
                    display_name,
                    "P00",
                    "SKU already used by '{}'".format(skus[sku]),
                    url
                )
                errors.append(error)
            else:
                skus[sku] = display_name
    return errors


class AppUI(object):
    """Base application user interface."""
    def __init__(self, args):
        self.args = args
        self.config = None
        self.eval_locals = {}
        self.rules = {}

    def validate_rules(self, rulesfile, rules):
        """Validate rules read from rules file."""
        failed = False
        for rule_id, rule in rules.items():
            if "itemtype" not in rule:
                self.error(
                    "Rule {} does not specify an itemtype in {}.".format(
                        rule_id,
                        rulesfile
                    )
                )
                failed = True
            elif rule["itemtype"] not in ["category", "product", "variant"]:
                self.error(
                    "Rule {} has an invalid itemtype in {}.".format(
                        rule_id,
                        rulesfile
                    )
                )
                failed = True

            if "test" not in rule:
                self.error(
                    "Rule {} does not specify a test in {}.".format(
                        rule_id,
                        rulesfile
                    )
                )
                failed = True
            else:
                if isinstance(rule["test"], list):
                    rule["test"] = " ".join(rule["test"])

            if "message" not in rule:
                self.error(
                    "ERROR: Rule {} does not specify a message in {}.".format(
                        rule_id,
                        rulesfile
                    )
                )
                failed = True

        if failed:
            sys.exit(1)

    def load_config_and_rules(self):
        """Load cctools configuration file and cclint rules files."""
        # Read config file.
        self.config = ConfigParser.RawConfigParser()
        self.config.optionxform = str  # preserve case of option names
        self.config.readfp(open(self.args.config))

        # Read rules files.
        for rulesfile in self.args.rules:
            try:
                constants_and_rules = json.load(
                    open(rulesfile),
                    object_pairs_hook=dupe_checking_hook
                )
            except Exception as ex:
                self.fatal(
                    "Error loading rules file {}:\n  {}".format(
                        rulesfile,
                        ex.message
                    )
                )

            #print(json.dumps(rules, indent=4, sort_keys=True))
            if "constants" in constants_and_rules:
                load_constants(
                    constants_and_rules["constants"],
                    self.eval_locals
                )
            if "rules" in constants_and_rules:
                file_rules = constants_and_rules["rules"]
                self.validate_rules(rulesfile, file_rules)

                # Merge rules from this file into master rules dict.
                if self.args.rule_ids is None:
                    rule_ids = None
                else:
                    rule_ids = self.args.rule_ids.split(",")
                for rule_id, rule in file_rules.items():
                    if (
                        (rule_ids is None or rule_id in rule_ids) and
                        "disabled" not in rule
                    ):
                        self.rules[rule_id] = rule

    def notify_checks_starting(self):
        """Notify user that time consuming checks are starting."""
        pass

    def notify_checks_completed(self):
        """Notify user that time consuming checks are complete."""
        pass

    def clear_error_list(self):
        """Clear error list."""
        pass

    def check_item(self, itemtype, item, item_name):
        """Check item for problems."""

        errors = []
        eval_globals = {"__builtins__": {"len": len, "re": re}}
        self.eval_locals["item"] = item

        for rule_id, rule in self.rules.items():
            # Skip rules that do not apply to this itemtype.
            if rule["itemtype"] != itemtype:
                continue

            try:
                success = eval(rule["test"], eval_globals, self.eval_locals)
            except SyntaxError:
                # SyntaxError probably means that the test is a statement,
                # not an expression.
                self.fatal(
                    "Syntax error in {} rule:\n{}".format(
                        rule_id,
                        rule["test"]
                    )
                )
            if not success:
                try:
                    # pylint: disable=W0142
                    message = rule["message"].format(**item)
                except Exception:
                    message = rule["message"]
                url = item_edit_url(self.config, itemtype, item)
                error = (itemtype, item_name, rule_id, message, url)
                errors.append(error)

        return errors

    def run_checks(self):
        """Run all checks, returning a list of errors."""

        self.clear_error_list()
        self.notify_checks_starting()

        errors = []

        # Create a connection to CoreCommerce.
        cc_browser = cctools.CCBrowser(
            self.config.get("website", "host"),
            self.config.get("website", "site"),
            self.config.get("website", "username"),
            self.config.get("website", "password"),
            clean=self.args.clean,
            cache_ttl=0 if self.args.refresh_cache else self.args.cache_ttl
        )

        # Any subsequent calls should ignore the cache.  If the user
        # clicks the Refresh button, it would be because they changed
        # something in CoreCommerce.
        self.args.refresh_cache = True

        # Check category list.
        categories = cc_browser.get_categories()
        self.eval_locals["items"] = categories
        for category in categories:
            errors.extend(
                self.check_item(
                    "category",
                    category,
                    category["Category Name"]
                )
            )

        # Check products list.
        cc_browser.guess_product_ids()
        products = sorted(
            cc_browser.get_products(),
            key=cc_browser.product_key_by_cat_and_name
        )
        self.eval_locals["items"] = products
        errors.extend(check_skus(self.config, products))
        for product in products:
            for key in ["Teaser"]:
                product[key] = cctools.html_to_plain_text(product[key])
            errors.extend(
                self.check_item(
                    "product",
                    product,
                    product_display_name(product)
                )
            )

        # Check variants list.
        variants = sorted(
            cc_browser.get_variants(),
            key=cc_browser.variant_key_by_cat_product
        )
        self.eval_locals["items"] = variants
        for variant in variants:
            errors.extend(
                self.check_item(
                    "variant",
                    variant,
                    variant_display_name(variant)
                )
            )

        self.notify_checks_completed()

        # Display errors.
        self.display_errors(errors)

    def error(self, msg):
        """Display an error message."""
        # pylint: disable=no-self-use
        print("ERROR: {}".format(msg), file=sys.stderr)

    def fatal(self, msg):
        """Display an error message and exit."""
        self.error(msg)
        sys.exit(1)

    def warning(self, msg):
        """Display a warning message."""
        # pylint: disable=no-self-use
        print("WARNING: {}".format(msg), file=sys.stderr)

    def display_errors(self, errors):
        """Print errors to stdout."""
        # pylint: disable=no-self-use
        for error in errors:
            # pylint: disable=W0142
            print("{0} '{1}'\n    {2}: {3}\n    {4}".format(*error))


#def sortby(tree, col, descending):
#    """Sort tree contents when a column header is clicked on."""
#    # grab values to sort
#    data = [(tree.set(child, col), child) \
#        for child in tree.get_children('')]
#    # if the data to be sorted is numeric change to float
#    #data =  change_numeric(data)
#    # now sort the data in place
#    data.sort(reverse=descending)
#    for index, item in enumerate(data):
#        tree.move(item[1], '', index)
#    # switch the heading so it will sort in the opposite direction
#    tree.heading(col, command=lambda col=col: sortby(tree, col, \
#        int(not descending)))


class AppGUI(AppUI):
    """Application graphical user interface."""
    # http://www.tkdocs.com/tutorial
    # http://www.tkdocs.com/tutorial/tree.html
    # https://www.daniweb.com/software-development/\
    #   python/threads/350266/creating-table-in-python

    # Width of filename text entry boxes.
    FILENAME_WIDTH = 95

    # pylint: disable=no-self-use
    def __init__(self, args, root):
        AppUI.__init__(self, args)
        self.root = root
        self.in_progress_window = None
        self.error_urls = {}
        self.frame = Tkinter.Frame(root)
        self.frame.pack(fill="both", expand=True)

        # Create treeview.
        self.tree = ttk.Treeview(self.frame, selectmode="browse")

        # Allow the tree cell to resize when the window is resized.
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(0, weight=1)

        # Configure first (tree) column
        self.tree.heading("#0", text="Item")
        # Setting the width dynamically in display_errors() doesn't
        # work, so set it explicitly here.
        self.tree.column("#0", width=350, minwidth=350, stretch=False)

        # Define data columns.
        # fix width for rule column
        self.tree["columns"] = ["Rule", "Problem"]
        self.tree.heading("Rule", text="Rule")
        self.tree.column(
            "Rule",
            width=tkFont.Font().measure("Rule"),
            stretch=False
        )
        self.tree.heading("Problem", text="Problem")
        self.tree.column("Problem", width=750)

        # Vertical scrollbar.
        # http://stackoverflow.com/questions/14359906
        # http://stackoverflow.com/questions/16746387/tkinter-treeview-widget
        ysb = ttk.Scrollbar(
            self.frame,
            orient="vertical",
            command=self.tree.yview
        )
        self.tree.configure(yscroll=ysb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")

        # Buttons frame.
        buttons_frame = Tkinter.Frame(self.frame)
        buttons_frame.grid(row=2, column=0, sticky="w")

        # Edit Item button
        self.edit_item_button = Tkinter.Button(
            buttons_frame,
            text="Edit Item",
            command=self.edit_item,
            default=Tkinter.ACTIVE
        )
        self.edit_item_button.pack(side=Tkinter.LEFT, padx=5, pady=5)

        # Recheck button
        self.recheck_button = Tkinter.Button(
            buttons_frame,
            text="Recheck",
            command=self.run_checks
        )
        self.recheck_button.pack(side=Tkinter.LEFT, padx=5, pady=5)

        # Quit button
        self.quit_button = Tkinter.Button(
            buttons_frame,
            text="Quit",
            command=self.frame.quit
        )
        self.quit_button.pack(side=Tkinter.LEFT, padx=5, pady=5)

        # Configure window resizing.
        self.root.minsize(600, 300)
        sizegrip = ttk.Sizegrip(self.frame)
        sizegrip.grid(row=2, column=1, sticky="se")

        self.load_config_and_rules()

        self.run_checks()

    def notify_checks_starting(self):
        """Notify user that time consuming checks are starting."""
        self.in_progress_window = Tkinter.Toplevel(self.root)
        Tkinter.Label(
            self.in_progress_window,
            text="Downloading data from CoreCommerce..."
        ).pack()
        self.in_progress_window.update()

    def notify_checks_completed(self):
        """Notify user that time consuming checks are complete."""
        self.in_progress_window.destroy()
        self.in_progress_window = None

    def clear_error_list(self):
        """Clear error list."""
        self.error_urls = {}
        for child in self.tree.get_children():
            self.tree.delete(child)

    def error(self, msg):
        """Display an error message."""
        tkMessageBox.showerror("Error", msg)

    def fatal(self, msg):
        """Display an error message and exit."""
        self.error(msg)
        self.frame.quit()
        sys.exit(1)

    def warning(self, msg):
        """Display a warning message."""
        tkMessageBox.showwarning("Warning", msg)

    def edit_item(self):
        """Display the selected item in a web browser."""
        selected = self.tree.focus()
        if selected in self.error_urls:
            url = self.error_urls[selected]
            webbrowser.open(url)
        else:
            self.error("URL for item {} not found.".format(selected))

    def display_errors(self, errors):
        """Display errors in table."""

        # Determine the minimum width of the Item column.
        item_col_width = tkFont.Font().measure("Item")

        # Insert errors in table.
        for itemtype in ("category", "product", "variant"):
            # Insert an itemtype branch.
            self.tree.insert(
                "",
                "end",
                itemtype,
                text=itemtype,
                open=True
            )
            width = tkFont.Font().measure(itemtype)
            if item_col_width < width:
                item_col_width = width

            # Insert items of the same itemtype.
            for error in errors:
                if error[0] == itemtype:
                    item_id = self.tree.insert(
                        error[0],
                        "end",
                        "",
                        text=error[1],
                        values=(error[2], error[3])
                    )
                    self.error_urls[item_id] = error[4]
                    width = tkFont.Font().measure(error[1])
                    if item_col_width < width:
                        item_col_width = width

        # Set width of the Item column.
        # This doesn't work, it makes the column too wide.
        # self.tree.column("#0", width=item_col_width)


class AppCLI(AppUI):
    """Application command line user interface."""
    # pylint: disable=no-self-use
    def __init__(self, args):
        AppUI.__init__(self, args)

        self.load_config_and_rules()

        self.run_checks()


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
        "--gui",
        action="store_true",
        default=False,
        help="display a GUI"
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

    if args.gui:
        root = Tkinter.Tk()
        root.title("cclint")
        AppGUI(args, root)
        root.mainloop()
    else:
        AppCLI(args)

    return 0


if __name__ == "__main__":
    main()
