#!/usr/bin/env python

"""
Web scraper interface to CoreCommerce.

TODO: download categories for sorting
TODO: use by gen-po-invoice.py
TODO: use by gen-art-mart-checkin.py
"""

#
# http://wwwsearch.sourceforge.net/mechanize/
# https://views.scraperwiki.com/run/python_mechanize_cheat_sheet/?

import csv
import mechanize  # sudo apt-get install python-mechanize
import os
import sys
import time

EXPORT_PRODUCTS = "m=ajax_export&instance=products&checkAccess=products"

CATEGORIES = [
    "Necklaces",
    "Bracelets",
    "Bags & Purses",
    "Baskets, Trivets & Bowls",
    "Miscellaneous"
]

class CCBrowser(object):
    """Encapsulate mechanize.Browser object."""
    def __init__(
        self,
        host,
        site,
        username,
        password,
        verbose=True,
        cache_ttl=3600
    ):
        self._base_url = "https://%s/~%s/admin/index.php" % (host, site)
        self._username = username
        self._password = password
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
        self._br = mechanize.Browser()
        self._logged_in = False

    def login(self):
        """Login to site."""

        # No need to login if we have already done so.
        if self._logged_in:
            return

        # Notify user of time consuming step.
        if self._verbose:
            sys.stderr.write("Logging into corecommerce.com\n")

        # Open the login page.
        self._br.open(self._base_url)

        # Find the login form.
        self._br.select_form(name="digiSHOP")

        # Set the form values.
        self._br['userId'] = self._username
        self._br['password'] = self._password

        # Submit the form (press the "Login" button).
        self._br.submit()
        self._logged_in = True

    def _load_export_page(self):
        """Load ajax_export page.  Not required to download
        products.csv, but it is useful for debugging.
        """

        # Open page.
        url = self._base_url + "?" + EXPORT_PRODUCTS
        self._br.open(url)

        # Select form.
        self._br.select_form("jsform")

        # Ensure that "All Categories" is selected.
        category_list = self._br.form.find_control("category")
        if False:  # debug
            for item in category_list.items:
                print " name=%s values=%s" % (
                    item.name,
                    str([label.text for label in item.get_labels()])
                )
        category_list.value = [""]  # name where values = ["All Categories"]

        # Submit the form (press the "" button).
        resp = self._br.submit()
        if False:  # debug
            print url + " response:\n"
            print resp.read().replace("\r", "")

    def _do_export(self):
        """Call doExport function.  This prepares a product list for
        download.
        """

        # Call the doExport function.
        url = self._base_url + "?" + EXPORT_PRODUCTS + "&rs=doExport"
        response = self._br.open(url)

        # Read the entire response.  This ensures that we do not
        # return until the server side prep is totally done.  The
        # response contains percentages for a progress bar.
        response.read()

    def download_products_csv(self, filename):
        """Download products list to a CSV file."""

        # Login if necessary.
        self.login()

        # Notify user of time consuming step.
        if self._verbose:
            sys.stderr.write("Downloading products\n")

        # Load the export page.
        self._load_export_page()

        # Prepare a product list for download.
        self._do_export()

        # Fetch the result file.
        url = self._base_url + "?m=ajax_export_send"
        self._br.retrieve(url, filename)

    def is_file_valid(self, filename):
        """Determine if a file exists and has not expired."""
        if not os.path.exists(filename):
            return False
        mtime = os.stat(filename).st_mtime
        expire_time = mtime + self._cache_ttl
        now = time.time()
        return now < expire_time

    def get_products(self):
        """Generate dictionary for each product downloaded from CoreCommerce."""

        filename = os.path.join(self._cache_dir, "products.csv")
        if not self.is_file_valid(filename):
            # Download products.csv.
            self.download_products_csv(filename)

        # Yield the product dictionaries.
        for product in csv.DictReader(open(filename)):
            yield product

    def sort_key_by_category_and_name(self, product):
        """Return a key for a product dictionary used to sort by
        category, product_name.
        """
        category = product["Category"]
        if category in CATEGORIES:
            category_index = CATEGORIES.index(category)
        else:
            category_index = len(CATEGORIES)
        return "%03i:%s" % (category_index, product["Product Name"])

    def sort_key_by_category(self, product):
        """Return a key for a product dictionary used to sort by
        category, product_name.
        """
        category = product["Category"]
        if category in CATEGORIES:
            category_index = CATEGORIES.index(category)
        else:
            category_index = len(CATEGORIES)
        return "%03i" % category_index
