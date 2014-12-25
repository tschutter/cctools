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

* recheck should flush cache
* in GUI, display message box when checking
* create href link to item
  * multiline CLI output
  * itemtype itemname
  *   message
  *   link
* fix width determination of item column
* freeze width of item and rule columns
* enable GUI scrollbars
* sort by column if column header clicked
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


def check_skus(products):
    """Check SKUs for uniqueness."""
    errors = []
    skus = dict()
    for product in products:
        display_name = product_display_name(product)
        sku = product["SKU"]
        if sku != "":
            if sku in skus:
                error = (
                    "product",
                    display_name,
                    "P00",
                    "SKU already used by '{}'".format(skus[sku])
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
                # FIXME
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
                error = (itemtype, item_name, rule_id, message)
                errors.append(error)

        return errors

    def run_checks(self):
        """Run all checks, returning a list of errors."""

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
        products = cc_browser.get_products()
        self.eval_locals["items"] = products
        errors.extend(check_skus(products))
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
        variants = cc_browser.get_variants()
        self.eval_locals["items"] = variants
        for variant in variants:
            errors.extend(
                self.check_item(
                    "variant",
                    variant,
                    variant_display_name(variant)
                )
            )

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
            print("{0} '{1}' {2} {3}".format(*error))


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

    # Width of filename text entry boxes.
    FILENAME_WIDTH = 95

    # pylint: disable=no-self-use
    def __init__(self, args, root):
        AppUI.__init__(self, args)
        self.root = root
        self.frame = Tkinter.Frame(root)
        self.frame.pack(fill="both", expand=True)

        # http://www.tkdocs.com/tutorial/tree.html
        # https://www.daniweb.com/software-development/\
        #   python/threads/350266/creating-table-in-python
        # Create treeview.
        self.tree = ttk.Treeview(self.frame)
        # selectmode, "browse" single select
        # tree.tag_bind('ttk', '<1>', itemClicked);
        # item clicked is tree.focus()

        # Configure first (tree) column
        self.tree.heading("#0", text="Item")
        self.tree.column("#0", width=350)

        # Define data columns.
        # fix width for rule column
        self.tree["columns"] = ["Rule", "Problem"]
        self.tree.heading("Rule", text="Rule")
        self.tree.column("Rule", width=tkFont.Font().measure("Rule"))
        self.tree.heading("Problem", text="Problem")
        self.tree.column("Problem", width=750)
            #sortable columns
            #self.tree.heading(
            #    column,
            #    text=column,
            #    command=lambda c=column: sortby(self.tree, c, 0)
            #)

#        self.tree.tag_configure("ttk")
        self.tree.pack(side=Tkinter.TOP, fill="both", expand=True)

# scrollbars
# http://stackoverflow.com/questions/16746387/tkinter-treeview-widget
#        ysb = ttk.Scrollbar(self, orient='vertical', command=self.tree.yview)
#        xsb = ttk.Scrollbar(self, orient='horizontal',
#             command=self.tree.xview)
#        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set)
#        self.tree.heading('#0', text='Path', anchor='w')
#
#        abspath = os.path.abspath(path)
#        root_node = self.tree.insert('', 'end', text=abspath, open=True)
#        self.process_directory(root_node, abspath)
#
#        self.tree.grid(row=0, column=0)
#        ysb.grid(row=0, column=1, sticky='ns')
#        xsb.grid(row=1, column=0, sticky='ew')
#        self.grid()

#        # CSV file
#        csv_group = Tkinter.LabelFrame(
#            self.frame,
#            text="CSV file",
#            padx=5,
#            pady=5
#        )
#        csv_group.pack(padx=10, pady=10)
#        self.csv_entry = Tkinter.Entry(csv_group, width=AppGUI.FILENAME_WIDTH)
#        self.csv_entry.pack(side=Tkinter.LEFT)
#
#        # QIF file
#        qif_group = Tkinter.LabelFrame(
#            self.frame,
#            text="QIF file",
#            padx=5,
#            pady=5
#        )
#        qif_group.pack(padx=10, pady=10)
#        self.qif_entry = Tkinter.Entry(qif_group, width=AppGUI.FILENAME_WIDTH)
#        self.qif_entry.pack(side=Tkinter.LEFT)

        # Recheck button
        self.recheck_button = Tkinter.Button(
            self.frame,
            text="Recheck",
            command=self.run_checks,
            default=Tkinter.ACTIVE
        )
        self.recheck_button.pack(side=Tkinter.LEFT, padx=5, pady=5)

        # Quit button
        self.quit_button = Tkinter.Button(
            self.frame,
            text="Quit",
            command=self.frame.quit
        )
        self.quit_button.pack(side=Tkinter.LEFT, padx=5, pady=5)

        self.load_config_and_rules()

        self.run_checks()

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

    def display_errors(self, errors):
        """Display errors in table."""

        # Delete existing items in table.
        for child in self.tree.get_children():
            self.tree.delete(child)

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
                    self.tree.insert(
                        error[0],
                        "end",
                        "",
                        text=error[1],
                        values=(error[2], error[3])
                    )
                    width = tkFont.Font().measure(error[1])
                    if item_col_width < width:
                        item_col_width = width

        # Set width of the Item column.
        # FIXME, doesn't work
        #self.tree.column("#0", width=item_col_width)
        #self.tree.pack(side=Tkinter.TOP, fill="both", expand=True)


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
        root.destroy()
    else:
        AppCLI(args)

    return 0


if __name__ == "__main__":
    main()
