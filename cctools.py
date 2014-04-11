#!/usr/bin/env python2

"""
Web scraper interface to CoreCommerce.
"""

from __future__ import print_function
import csv
import json
import mechanize  # sudo apt-get install python-mechanize
import os
import re
import sys
import time

# http://wwwsearch.sourceforge.net/mechanize/
# https://views.scraperwiki.com/run/python_mechanize_cheat_sheet/?

# pylint seems to be confused by calling methods via self._browser.
# pylint: disable=E1102

class CCBrowser(object):
    """Encapsulate mechanize.Browser object."""
    def __init__(
        self,
        host,
        site,
        username,
        password,
        clean=True,
        verbose=True,
        cache_ttl=3600,
        proxy=None
    ):
        self._host = host
        self._base_url = "https://%s/~%s" % (host, site)
        self._admin_url = self._base_url + "/admin/index.php"
        self._username = username
        self._password = password
        self._clean = clean
        self._verbose = verbose
        self._cache_ttl = float(cache_ttl)
        self._cache_dir = os.environ["HOME"] + "/.cache/cctools"
        if not os.path.exists(self._cache_dir):
            os.mkdir(self._cache_dir, 0o700)
        self._browser = mechanize.Browser()
        if proxy != None:
            self._browser.set_proxies({"https": proxy})
        self._logged_in = False
        self._personalizations = None
        self._products = None
        self._categories = None
        self._category_sort = None

    def _login(self):
        """Login to site."""

        # No need to login if we have already done so.
        if self._logged_in:
            return

        # Notify user of time consuming step.
        if self._verbose:
            sys.stderr.write("Logging into %s\n" % self._host)

        # Open the login page.
        self._browser.open(self._admin_url)

        # Find the login form.
        self._browser.select_form(name="digiSHOP")

        # Set the form values.
        self._browser["userId"] = self._username
        self._browser["password"] = self._password

        # Submit the form (press the "Login" button).
        self._browser.submit()
        self._logged_in = True

    def _is_file_expired(self, filename):
        """Determine if a file doesn't exist or has expired."""
        if not os.path.exists(filename):
            return True
        mtime = os.stat(filename).st_mtime
        expire_time = mtime + self._cache_ttl
        now = time.time()
        return expire_time < now

    def _do_export(self, url, filename):
        """Export a file from CoreCommerce."""
        # This method was derived from the following javascript code
        # returned by pressing the "Export" button.
        #
        # jQuery.ajax({
        #     type : "GET",
        #     cache : false,
        #     url : 'https://HOST/~SITE/controllers/ajaxController.php',
        #     data : {
        #         object : 'ExportAjax',
        #         'function' : 'processExportCycle',
        #         current : current
        #     }
        # })
        # .done(function(response) {
        #     var responseObject = jQuery.parseJSON(response);
        #     var current = responseObject.current;
        #     var percentComplete = responseObject.percentComplete;
        #     if(percentComplete == '100'){
        #         var url = 'https://HOST/~SITE/admin/index.php?m=ajax_export_send';
        #         var parent = window.opener;
        #         parent.location = url;
        #     } else {
        #         doExport(current);
        #     }
        # })

        # Call the processExportCycle function until percentComplete == 100.
        ajax_controller_url = self._base_url + "/controllers/ajaxController.php"
        current = 0
        while True:
            url = "%s?object=ExportAjax&function=processExportCycle&current=%i" % (
                ajax_controller_url,
                current
            )
            response = self._browser.open(url).read()
            response_object = json.loads(response)
            if response_object["percentComplete"] == 100:
                break
            current = response_object["current"]

        # Fetch the result file.
        url = self._admin_url + "?m=ajax_export_send"
        self._browser.retrieve(url, filename)

    def _download_personalizations_csv(self, filename):
        """Download personalization list to a CSV file."""

        # Login if necessary.
        self._login()

        # Notify user of time consuming step.
        if self._verbose:
            sys.stderr.write("Downloading personalizations\n")

        # Load export page.
        url = "%s?%s" % (
            self._admin_url,
            "m=ajax_export" +
            "&instance=personalization_products&checkAccess=products"
        )
        self._browser.open(url)

        # Call the doExport function.
        self._do_export(url, filename)

    def _clean_personalizations(self):
        """Normalize suspect personalization data."""
        # Boolean value of "" appears to mean "N".
        booleans = [
            "Answer Enabled",
            "Default",
            "Exclude from best seller report",
            "Question Enabled",
            "Required",
            "Track Inventory"
        ]
        for personalization in self._personalizations:
            # Booleans should be Y|N, but we sometimes see "".
            for boolean in booleans:
                if not personalization[boolean] in ("Y", "N"):
                    personalization[boolean] = "N"

    def get_personalizations(self):
        """Return a list of per-personalization dictionaries."""

        if self._personalizations == None:
            filename = os.path.join(self._cache_dir, "personalizations.csv")

            # Download products file if it is out of date.
            if self._is_file_expired(filename):
                self._download_personalizations_csv(filename)

            # Read personalizations file.
            self._personalizations = list(csv.DictReader(open(filename)))

            # Cleanup suspect data.
            if self._clean:
                self._clean_personalizations()

        return self._personalizations

    def _download_products_csv(self, filename):
        """Download products list to a CSV file."""

        # Login if necessary.
        self._login()

        # Notify user of time consuming step.
        if self._verbose:
            sys.stderr.write("Downloading products\n")

        # Load the export page.
        url = (
            self._admin_url +
            "?m=ajax_export&instance=products&checkAccess=products"
        )
        self._browser.open(url)

        # Select form.
        self._browser.select_form("jsform")

        # Ensure that "All Categories" is selected.
        category_list = self._browser.form.find_control("category")
        if False:  # debug
            for item in category_list.items:
                print(
                    " name=%s values=%s" % (
                        item.name,
                        str([label.text for label in item.get_labels()])
                    )
                )
        category_list.value = [""]  # name where values = ["All Categories"]

        # Submit the form (press the "Export" button).
        resp = self._browser.submit()
        if False:  # debug
            # Examine the source of the doExport method.
            print("Response from %s:\n" % url)
            print(resp.read().replace("\r", ""))

        # Call the doExport function.
        self._do_export(url, filename)

    def _clean_products(self):
        """Normalize suspect product data."""
        # Boolean value of "" appears to mean "N".
        booleans = [
            "Available",
            "Customer Must Add To Cart To See Sales Price",
            "Discontinued Item",
            "Display Facebook LIKE",
            "Display Facebook Link",
            "Display Twitter Link",
            "Eligible For Reward Points",
            "Featured Product",
            "Ignore Default Images",
            "Include in Bing Product Feed",
            "Include in Google Product Feed",
            "Password protect this product",
            "Request a Lower Price",
            "Taxable (GST)",
            "Taxable (HST)",
            "Taxable (PST)",
            "Taxable",
            "Use Main Photo as Product Detail Thumbnail",
            "Use Sale Price",
            "Use Tab Navigation"
        ]
        for product in self._products:
            # Booleans should be Y|N, but we sometimes see "".
            for boolean in booleans:
                if not product[boolean] in ("Y", "N"):
                    product[boolean] = "N"

    def get_products(self):
        """Return a list of per-product dictionaries."""

        if self._products == None:
            filename = os.path.join(self._cache_dir, "products.csv")

            # Download products file if it is out of date.
            if self._is_file_expired(filename):
                self._download_products_csv(filename)

            # Read products file.
            self._products = list(csv.DictReader(open(filename)))

            # Cleanup suspect data.
            if self._clean:
                self._clean_products()

        return self._products

    _PRODUCT_KEY_MAP = {
        "Price": "pPrice"
    }

    def is_valid_product_update_key(self, key):
        """Return True if key is a valid product update key."""
        return key in self._PRODUCT_KEY_MAP

    def update_product(self, sku, key, value):
        """Login to site."""

        # Login if necessary.
        self._login()

        # Notify user of time consuming step.
        if self._verbose:
            sys.stderr.write(
                "Updating product SKU=%s, setting %s to %s\n" % (
                    sku,
                    key,
                    value
                )
            )

        # Open the upload page.
        self._browser.open(
            self._admin_url + "?m=ajax_import&instance=product_import"
        )
        for form in self._browser.forms():
            print("Form name:", form.name)
            print(form)

        # Select first and only form on page.
        self._browser.select_form(nr=0)

        #print [item.name for item in form.find_control('useFile').items]
        # Set the form values.
        #self._browser["instance"] = "product_import"
        #self._browser["xsubmit"] = "true"
        #self._browser["file"] = "cctools.csv"
        #self._browser["useFile"] = ["Y",]
        with open("/tmp/cctools.csv", "wt") as tfile:
            tfile.write("SKU,%s\n%s,%s\n" % (key, sku, value))
        tfile.close()
        self._browser.form.add_file(
            open("/tmp/cctools.csv"),
            "text/csv",
            "/tmp/cctools.csv",
            name="importFile"
        )
        #self._browser["importFile"] = "SKU,%s\n%s,%s\n" % (key, sku, value)
        self._browser["updateType"] = "update"

        # Submit the form (press the "????" button).
        self._browser.submit()

        # https://www16.corecommerce.com/~cohu1/admin/index.php?m=ajax_import&instance=product_import
        # POST https://www16.corecommerce.com/~cohu1/admin/index.php
        #   m:           ajax_import
        #   instance:    product_import
        #   xsubmit:     true
        #   file:
        #   useFile:     Y
        #   importFile:  SKU,Price
        #                00000,5.62
        #   updateType:  update

        # Open the update page.
        self._browser.open(
            self._admin_url + "?m=ajax_import_save&instance=product_import"
        )
        for form in self._browser.forms():
            print("Form name:", form.name)
            print(form)

        # Select first and only form on page.
        self._browser.select_form(nr=0)

        # Set the form values.
        #self._browser["instance"] = "product_import"
        self._browser["go"] = self._admin_url + "?m=ajax_import&instance=product_import"
        self._browser["submit"] = "true"
        self._browser["file"] = "cctools.csv"
        self._browser["useFile"] = "Y"
        self._browser["instance"] = "product_import"
        self._browser["updateType"] = "update"
        self._browser["fields[0]"] = "pNum"
        self._browser["fields[1]"] = self._PRODUCT_KEY_MAP[key]
        self._browser["ignore"] = "Y"

        # Submit the form (press the "????" button).
        self._browser.submit()

        # POST https://www16.corecommerce.com/~cohu1/admin/index.php?m=ajax_import_save&instance=product_import
        #   m:           ajax_import_save
        #   go:          https://www16.corecommerce.com/~cohu1/admin/index.php?m=ajax_import&instance=product_import
        #   submit:      true
        #   file:        test-prod-update.csv
        #   useFile:     Y
        #   instance:    product_import
        #   updateType:  update
        #   fields[0]:   pNum
        #   fields[1]:   pPrice
        #   ignore:      Y

    def _download_categories_csv(self, filename):
        """Download categories list to a CSV file."""

        # Login if necessary.
        self._login()

        # Notify user of time consuming step.
        if self._verbose:
            sys.stderr.write("Downloading categories\n")

        # Load the export page.
        url = (
            self._admin_url +
            "?m=ajax_export&instance=categories&checkAccess=categories"
        )
        self._browser.open(url)

        # Call the doExport function.
        self._do_export(url, filename)

    def _clean_categories(self):
        """Normalize suspect product data."""
        # Boolean value of "" appears to mean "N".
        for product in self._categories:
            if not product["Hide This Category From Customers"] in ("Y", "N"):
                product["Available"] = "N"

    def get_categories(self):
        """Return a list of per-category dictionaries."""

        if self._categories == None:
            filename = os.path.join(self._cache_dir, "categories.csv")

            # Download categories file if it is out of date.
            if self._is_file_expired(filename):
                self._download_categories_csv(filename)

            # Read categories file.
            self._categories = list(csv.DictReader(open(filename)))

            # Cleanup suspect data.
            if self._clean:
                self._clean_categories()

        return self._categories

    def set_category_sort_order(self, categories):
        """Build the dictionary used for sorting by category based
        upon the specified category names.
        """
        self._category_sort = dict()
        for sort, name in enumerate(categories):
            self._category_sort[name] = sort

    def _init_category_sort(self):
        """Build the dictionary used for sorting by category based
        upon the category Sort value from CoreCommerce.
        """
        if self._category_sort == None:
            if self._categories == None:
                self.get_categories()
            self._category_sort = dict()
            for category in self._categories:
                name = category["Category Name"]
                sort = int(category["Sort"])
                self._category_sort[name] = sort

    def sort_key_by_category_and_name(self, product):
        """Return a key for a product dictionary used to sort by
        category, product_name.
        """
        if self._category_sort == None:
            self._init_category_sort()
        category = product["Category"]
        if category in self._category_sort:
            category_sort_key = "%05i" % self._category_sort[category]
        else:
            category_sort_key = category
        return "%s:%s" % (category_sort_key, product["Product Name"])

    def sort_key_by_category(self, product):
        """Return a key for a product dictionary used to sort by
        category, product_name.
        """
        if self._category_sort == None:
            self._init_category_sort()
        category = product["Category"]
        if category in self._category_sort:
            category_sort_key = "%05i" % self._category_sort[category]
        else:
            category_sort_key = category
        return category_sort_key

    def sort_key_by_sku(self, product):
        """Return a key for a product dictionary used to sort by sku."""
        return product["SKU"]

    def personalization_sort_key(self, personalization):
        """Return a key to sort personalizations."""
        return (
            personalization["Product SKU"],
            personalization["Question Sort Order"],
            personalization["Answer Sort Order"]
        )

_HTML_TO_PLAIN_TEXT_DICT = {
    "&quot;": "\"",
    "&amp;": "&",
    "<p>": " ",
    "</p>": " "
}
_HTML_TO_PLAIN_TEXT_RE = re.compile('|'.join(_HTML_TO_PLAIN_TEXT_DICT.keys()))

def html_to_plain_text(string):
    """Convert HTML markup to plain text."""

    # Replace HTML markup with plain text."""
    string = _HTML_TO_PLAIN_TEXT_RE.sub(
        lambda m: _HTML_TO_PLAIN_TEXT_DICT[m.group(0)],
        string
    )

    # Collapse all whitespace to a single space.
    string = re.sub("\s+", " ", string)

    # Strip leading and trailing whitespace.
    string = string.strip()

    return string


def plain_text_to_html(string):
    """Convert plain text to HTML markup."""
    string = string.replace("&", "&amp;")
    return string
