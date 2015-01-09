#!/usr/bin/env python2

r"""
Detects problems in CoreCommerce product data.

Requires a [website] section in config file for login info.

The rules file (default=cclint.rules) contains product value checks in
YAML format.  The schema is:

    constants:
        CONST_NAME1: CONST_VALUE1
        CONST_NAME2: CONST_VALUE2
        ...
    },
    rules:
        RULE1_ID:
            disabled: message why rule disabled (key optional)
            itemtype: category|product|variant
            test: Python predicate, True means test passed
            message: message output if test returned False
        RULE2_ID:
            ...
    ]

The message is formatted using message.format(current_item), so it can
contain member variable references like "{SKU}".

For YAML syntax, see:
* http://wikipedia.org/wiki/YAML
* http://pyyaml.org/wiki/PyYAMLDocumentation
"""

from __future__ import print_function
import ConfigParser
import Tkinter
import argparse
import cctools
import logging
import os
import re
import sys
import tkFont
import tkMessageBox
import ttk
import webbrowser
import yaml  # sudo pip install pyyaml


# def dupe_checking_hook(pairs):
#     """
#     An object_pairs_hook for json.load() that raises a KeyError on
#     duplicate or blank keys.
#     """
#     result = dict()
#     for key, val in pairs:
#         if key.strip() == "":
#             raise KeyError("Blank key specified")
#         if key in result:
#             raise KeyError("Duplicate key specified: %s" % key)
#         result[key] = val
#     return result


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
    findings = []
    skus = dict()
    for product in products:
        display_name = product_display_name(product)
        sku = product["SKU"]
        if sku != "":
            if sku in skus:
                url = item_edit_url(config, "product", product)
                finding = (
                    "product",
                    display_name,
                    "P00",
                    "SKU already used by '{}'".format(skus[sku]),
                    url
                )
                findings.append(finding)
            else:
                skus[sku] = display_name
    return findings


def finding_to_string(finding):
    """Create a string representation of a finding."""
    # pylint: disable=W0142
    return "{0} '{1}'\n    {2}: {3}\n    {4}".format(*finding)


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
                rule["test"] = rule["test"].strip()

            if "message" not in rule:
                self.error(
                    "Rule {} does not specify a message in {}.".format(
                        rule_id,
                        rulesfile
                    )
                )
                failed = True
            else:
                rule["message"] = rule["message"].strip()

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
                constants_and_rules = yaml.load(open(rulesfile))
            except Exception as ex:
                self.fatal(
                    "Error loading rules file {}:\n  {}".format(
                        rulesfile,
                        str(ex).replace("\n", "\n  ")
                    )
                )

            if "constants" in constants_and_rules:
                load_constants(
                    constants_and_rules["constants"],
                    self.eval_locals
                )
            if "rules" in constants_and_rules:
                file_rules = constants_and_rules["rules"]
                # print(yaml.dump(file_rules, default_flow_style=False))
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

    def checks_completed(self):
        """Notify user that time consuming checks are complete."""
        pass

    def clear_finding_list(self):
        """Clear finding list."""
        pass

    def check_item(self, itemtype, item, item_name):
        """Check item for problems."""

        findings = []
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
                finding = (itemtype, item_name, rule_id, message, url)
                findings.append(finding)

        return findings

    def run_checks_core(self):
        """Run all checks, returning a list of findings."""

        self.clear_finding_list()

        findings = []

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
            findings.extend(
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
        findings.extend(check_skus(self.config, products))
        for product in products:
            for key in ["Teaser"]:
                product[key] = cctools.html_to_plain_text(product[key])
            findings.extend(
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
            findings.extend(
                self.check_item(
                    "variant",
                    variant,
                    variant_display_name(variant)
                )
            )

        self.checks_completed()

        return findings

    def run_checks(self):
        """Run all checks, displaying findings."""
        try:
            findings = self.run_checks_core()
            self.display_findings(findings)
        except Exception as ex:
            self.fatal(ex)

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

    def display_findings(self, findings):
        """Print findings to stdout."""
        # pylint: disable=no-self-use
        first = True
        for finding in findings:
            print(
                "{0}{1}".format(
                    "" if first else "\n",
                    finding_to_string(finding)
                )
            )
            first = False


class AppGUI(AppUI):
    """Application graphical user interface."""
    # http://www.tkdocs.com/tutorial
    # http://www.tkdocs.com/tutorial/tree.html
    # https://www.daniweb.com/software-development/\
    #   python/threads/350266/creating-table-in-python

    # pylint: disable=no-self-use
    def __init__(self, args, root):
        AppUI.__init__(self, args)
        self.root = root
        self.tree_item_finding = {}
        self.frame = Tkinter.Frame(root)
        self.frame.pack(fill="both", expand=True)

        # Create treeview.
        self.tree = ttk.Treeview(self.frame, selectmode="browse")

        # Allow the tree cell to resize when the window is resized.
        self.frame.columnconfigure(0, weight=1)
        self.frame.rowconfigure(0, weight=1)

        # Configure first (tree) column
        self.tree.heading("#0", text="Item")
        self.tree.column("#0", stretch=False)

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

        # Logging label
        logging_label = Tkinter.Label(buttons_frame)
        logging_label.pack(side=Tkinter.LEFT, padx=5, pady=5)

        # Configure window resizing.
        self.root.minsize(
            1000,  # width not including window border
            600    # height not including window border and title bar
        )
        sizegrip = ttk.Sizegrip(self.frame)
        sizegrip.grid(row=2, column=1, sticky="se")

        # Configure logging.
        logging.basicConfig(level=logging.INFO)
        logger = logging.getLogger()
        self.label_logging_handler = self.LabelLoggingHandler(logging_label)
        logger.addHandler(self.label_logging_handler)

        self.load_config_and_rules()

        # Setup key bindings.
        self.root.bind_all("<Control-c>", self.copy_to_clipboard)
        self.root.bind_all("<Alt-w>", self.copy_to_clipboard)

        # Set window title.
        self.root.title(
            "cclint - {}".format(self.config.get("website", "site"))
        )

        self.run_checks()

    class LabelLoggingHandler(logging.Handler):
        """A logging Handler that displays records in a main window label."""
        def __init__(self, logging_label):
            logging.Handler.__init__(self, level=logging.INFO)
            self.setFormatter(logging.Formatter('%(msg)s'))
            self.logging_label = logging_label

        def emit(self, record):
            """Override the default handler's emit method."""
            message = self.format(record)
            self.logging_label.configure(text=message)
            self.logging_label.update()

        def blank(self):
            """Clear the displayed text."""
            self.logging_label.configure(text="")
            self.logging_label.update()

    def copy_to_clipboard(self, _):
        """Copy selected finding to clipboard."""
        selected = self.tree.focus()
        if selected in self.tree_item_finding:
            finding = self.tree_item_finding[selected]
            self.root.clipboard_clear()
            self.root.clipboard_append(finding_to_string(finding) + "\n")

    def checks_completed(self):
        """Notify user that time consuming checks are complete."""
        self.label_logging_handler.blank()

    def clear_finding_list(self):
        """Clear finding list."""
        self.tree_item_finding = {}
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
        if selected in self.tree_item_finding:
            finding = self.tree_item_finding[selected]
            url = finding[4]
            webbrowser.open(url)
        else:
            self.error("URL for item {} not found.".format(selected))

    def display_findings(self, findings):
        """Display findings in a tree view."""

        # Starting minimum width of the Item column.
        item_col_width = 100

        # Second level items are indented.
        tree_indent = 38

        # Insert findings into tree view.
        for itemtype in ("category", "product", "variant"):
            # Insert an itemtype branch.
            self.tree.insert(
                "",
                "end",
                itemtype,
                text=itemtype,
                open=True
            )

            # Insert findings of the same itemtype.
            for finding in findings:
                if finding[0] == itemtype:
                    item_id = self.tree.insert(
                        finding[0],
                        "end",
                        "",
                        text=finding[1],
                        values=(finding[2], finding[3])
                    )
                    self.tree_item_finding[item_id] = finding
                    # Width returned by measure() appears to be 2
                    # pixels too wide for each character.
                    width = (
                        tkFont.Font().measure(finding[1])
                        - 2 * len(finding[1])
                        + tree_indent
                    )
                    if item_col_width < width:
                        item_col_width = width

        # Set width of the Item column.
        self.tree.column("#0", width=item_col_width)


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
