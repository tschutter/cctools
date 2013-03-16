#!/usr/bin/env python

"""
Web scraper interface to CoreCommerce.
"""

import csv
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
        self._base_url = "https://%s/~%s/admin/index.php" % (host, site)
        self._username = username
        self._password = password
        self._clean = clean
        self._verbose = verbose
        self._cache_ttl = float(cache_ttl)
        if "USER" in os.environ:
            username = os.environ["USER"]
        elif "USERNAME" in os.environ:
            username = os.environ["USERNAME"]
        else:
            username = "UNKNOWN"
        self._cache_dir = "/tmp/cctools-cache-" + username
        if not os.path.exists(self._cache_dir):
            os.mkdir(self._cache_dir, 0700)
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
        self._browser.open(self._base_url)

        # Find the login form.
        self._browser.select_form(name="digiSHOP")

        # Set the form values.
        self._browser['userId'] = self._username
        self._browser['password'] = self._password

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

    def _sajax_do_call(self, url, method, args):
        """Call a sajax method and return the response."""
        url += "&rs=" + method
        for arg in args:
            url += "&rsargs[]=" + arg
        response = self._browser.open(url).read()
        txt = response.strip()
        status = txt[0]
        data = txt[2:]
        if (status == "-"):
            print "ERROR: " + data
            return
        # Javascript uses eval(data) here to convert data to objects.
        # For now, we hack.
        match = re.match("var res = '(.*)'; res;", data)
        if match:
            return match.group(1)

    def _do_export(self, url, filename):
        """Export a file from CoreCommerce."""

        # Call the doExport function until percent_complete == 100.
        current = ""
        while True:
            res = self._sajax_do_call(url, "doExport", [current])
            pieces = res.split("|")
            if len(pieces) >= 1 and pieces[1] == "100":
                break
            current = pieces[0]

        # Fetch the result file.
        url = self._base_url + "?m=ajax_export_send"
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
            self._base_url,
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
            self._base_url +
            "?m=ajax_export&instance=products&checkAccess=products"
        )
        self._browser.open(url)

        # Select form.
        self._browser.select_form("jsform")

        # Ensure that "All Categories" is selected.
        category_list = self._browser.form.find_control("category")
        if False:  # debug
            for item in category_list.items:
                print " name=%s values=%s" % (
                    item.name,
                    str([label.text for label in item.get_labels()])
                )
        category_list.value = [""]  # name where values = ["All Categories"]

        # Submit the form (press the "" button).
        resp = self._browser.submit()
        if False:  # debug
            print url + " response:\n"
            print resp.read().replace("\r", "")

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

    def _download_categories_csv(self, filename):
        """Download categories list to a CSV file."""

        # Login if necessary.
        self._login()

        # Notify user of time consuming step.
        if self._verbose:
            sys.stderr.write("Downloading categories\n")

        # Load the export page.
        url = (
            self._base_url +
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

    def _init_category_sort(self):
        """Build the dictionary used for sorting by category."""
        if self._category_sort == None:
            if self._categories == None:
                self.get_categories()
            category_sort = dict()
            for category in self._categories:
                name = category["Category Name"]
                sort = int(category["Sort"])
                category_sort[name] = sort
            self._category_sort = category_sort

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

    def personalization_sort_key(self, personalization):
        """Return a key to sort personalizations."""
        return (
            personalization["Product Name"],
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
