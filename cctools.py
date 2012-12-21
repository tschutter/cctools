#!/usr/bin/env python

"""
Web scraper interface to CoreCommerce.

TODO: drop generator for products; keep _products list
"""

# http://wwwsearch.sourceforge.net/mechanize/
# https://views.scraperwiki.com/run/python_mechanize_cheat_sheet/?

import csv
import mechanize  # sudo apt-get install python-mechanize
import os
import sys
import time

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
        self._categories = None
        self._category_sort = None

    def _login(self):
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

    def _is_file_expired(self, filename):
        """Determine if a file doesn't exist or has expired."""
        if not os.path.exists(filename):
            return True
        mtime = os.stat(filename).st_mtime
        expire_time = mtime + self._cache_ttl
        now = time.time()
        return expire_time < now

    _EXPORT_PRODUCTS_PAGE =\
        "m=ajax_export&instance=products&checkAccess=products"

    def _load_export_products_page(self):
        """Load ajax_export products page."""

        # Open page.
        url = "%s?%s" % (self._base_url, self._EXPORT_PRODUCTS_PAGE)
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

    def _do_export_products(self):
        """Call doExport function on the ajax_export products page.
        This prepares a product list for download.
        """

        # Call the doExport function.
        url = "%s?%s&rs=doExport" % (self._base_url, self._EXPORT_PRODUCTS_PAGE)
        response = self._br.open(url)

        # Read the entire response.  This ensures that we do not
        # return until the server side prep is totally done.  The
        # response contains percentages for a progress bar.
        response.read()

    def _download_products_csv(self, filename):
        """Download products list to a CSV file."""

        # Login if necessary.
        self._login()

        # Notify user of time consuming step.
        if self._verbose:
            sys.stderr.write("Downloading products\n")

        # Load the export page.
        self._load_export_products_page()

        # Prepare a product list for download.
        self._do_export_products()

        # Fetch the result file.
        url = self._base_url + "?m=ajax_export_send"
        self._br.retrieve(url, filename)

    def get_products(self):
        """Generate dictionary for each product downloaded from CoreCommerce."""

        filename = os.path.join(self._cache_dir, "products.csv")
        if self._is_file_expired(filename):
            # Download products.csv.
            self._download_products_csv(filename)

        # Yield the product dictionaries.
        for product in csv.DictReader(open(filename)):
            yield product

    _EXPORT_CATEGORIES_PAGE =\
        "m=ajax_export&instance=categories&checkAccess=categories"

    def _load_export_categories_page(self):
        """Load ajax_export categories page."""

        # Open page.
        url = "%s?%s" % (self._base_url, self._EXPORT_CATEGORIES_PAGE)
        self._br.open(url)

    def _do_export_categories(self):
        """Call doExport function on the ajax_export categories page.
        This prepares a category list for download.
        """

        # Call the doExport function.
        url = "%s?%s&rs=doExport" % (
            self._base_url,
            self._EXPORT_CATEGORIES_PAGE
        )
        response = self._br.open(url)

        # Read the entire response.  This ensures that we do not
        # return until the server side prep is totally done.  The
        # response contains percentages for a progress bar.
        response.read()

    def _download_categories_csv(self, filename):
        """Download categories list to a CSV file."""

        # Login if necessary.
        self._login()

        # Notify user of time consuming step.
        if self._verbose:
            sys.stderr.write("Downloading categories\n")

        # Load the export page.
        self._load_export_categories_page()

        # Prepare a category list for download.
        self._do_export_categories()

        # Fetch the result file.
        url = self._base_url + "?m=ajax_export_send"
        self._br.retrieve(url, filename)

    def get_categories(self):
        """Return a dictionary for each category downloaded from
        CoreCommerce.
        """

        if self._categories == None:
            filename = os.path.join(self._cache_dir, "categories.csv")
            if self._is_file_expired(filename):
                # Download categories.csv.
                self._download_categories_csv(filename)

            self._categories = [
                category for category in csv.DictReader(open(filename))
            ]

        return self._categories

    def _init_category_sort(self):
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
