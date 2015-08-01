#!/usr/bin/env python2

"""
Web scraper interface to CoreCommerce.
"""

from __future__ import print_function
import csv
import json
import lockfile  # sudo apt-get install python-lockfile
import logging
import mechanize  # sudo apt-get install python-mechanize
import os
import re
import tempfile
import time
try:
    # pylint: disable=F0401
    from xdg.BaseDirectory import xdg_cache_home
except ImportError:
    # xdg not available on all platforms
    # pylint: disable=C0103
    xdg_cache_home = os.path.join(os.path.expanduser("~"), ".cache")

# Notes:
#
# http://wwwsearch.sourceforge.net/mechanize/
# https://views.scraperwiki.com/run/python_mechanize_cheat_sheet/?
#
# CoreCommerce is schizophrenic regarding personalizations
# vs. variants.  Most likely they were originally called
# personalizations and were later renamed to variants.  Except they
# didn't change all of the references on the web site.  I chose to
# call them variants here.

# pylint seems to be confused by calling methods via self._browser.
# pylint: disable=E1102

LOGGER = logging.getLogger(__name__)


class CCBrowser(object):
    """Encapsulate mechanize.Browser object."""
    def __init__(
        self,
        host,
        site,
        username,
        password,
        clean=True,
        cache_ttl=3600,
        proxy=None
    ):
        self._host = host
        self._base_url = "https://{}/~{}".format(host, site)
        self._admin_url = self._base_url + "/admin/index.php"
        self._username = username
        self._password = password
        self._clean = clean
        self._cache_ttl = float(cache_ttl)
        self._cache_dir = os.path.join(xdg_cache_home, "cctools")
        if not os.path.exists(self._cache_dir):
            os.mkdir(self._cache_dir, 0o700)
        # A single lockfile is used for all download operations.  We
        # have no idea if the CoreCommerce ajax_export can support
        # simultaneous downloads so we play it safe.  The lockfile
        # module will append ".lock" to the filename.
        self._download_lock_filename = os.path.join(
            self._cache_dir,
            "download"
        )
        self._browser = mechanize.Browser()
        if proxy is not None:
            self._browser.set_proxies({"https": proxy})
        self._logged_in = False
        self._variants = None
        self._questions = None
        self._products = None
        self._categories = None
        self._category_sort = None

    def _select_form(self, name):
        """Select a form in the browser."""
        try:
            self._browser.select_form(name)
        except mechanize.FormNotFoundError as ex:
            forms = []
            for form in self._browser.forms():
                forms.append(form.name)
            raise type(ex)(str(ex) + " in {}".format(forms))

    def _login(self):
        """Login to site."""

        # No need to login if we have already done so.
        if self._logged_in:
            return

        # Log time consuming step.
        LOGGER.info("Logging into {}".format(self._host))
        LOGGER.debug("Username = {}".format(self._username))

        # Open the login page.
        self._browser.open(self._admin_url)

        # Find the login form.
        self._select_form("digiSHOP")

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
        #   type : "GET",
        #   cache : false,
        #   url : 'https://HOST/~SITE/controllers/ajaxController.php',
        #   data : {
        #     object : 'ExportAjax',
        #     'function' : 'processExportCycle',
        #     current : current
        #   }
        # })
        # .done(function(response) {
        #   var responseObject = jQuery.parseJSON(response);
        #   var current = responseObject.current;
        #   var percentComplete = responseObject.percentComplete;
        #   if(percentComplete == '100'){
        #     var url = 'https://HST/~SITE/admin/index.php?m=ajax_export_send';
        #     var parent = window.opener;
        #     parent.location = url;
        #   } else {
        #     doExport(current);
        #   }
        # })

        # Call the processExportCycle function until percentComplete == 100.
        ajax_controller_url =\
            self._base_url + "/controllers/ajaxController.php"
        current = 0
        while True:
            url = (
                "{}"
                "?object=ExportAjax"
                "&function=processExportCycle"
                "&current={}"
            ).format(ajax_controller_url, current)
            response = self._browser.open(url).read()
            response_object = json.loads(response)
            if response_object["percentComplete"] == 100:
                break
            current = response_object["current"]

        # Fetch the result file.
        url = self._admin_url + "?m=ajax_export_send"
        self._browser.retrieve(url, filename)

    def _download_variants_csv(self, filename):
        """Download variant list to a CSV file."""

        # Login if necessary.
        self._login()

        # Log time consuming step.
        LOGGER.info("Downloading variants")

        # Load export page.
        url = "{}?{}{}".format(
            self._admin_url,
            "m=ajax_export",
            "&instance=personalization_products&checkAccess=products"
        )
        self._browser.open(url)

        # Call the doExport function.
        self._do_export(url, filename)

    def _clean_variants(self):
        """Normalize suspect variant data."""
        # Boolean value of "" appears to mean "N".
        booleans = [
            "Answer Enabled",
            "Default",
            "Exclude from best seller report",
            "Question Enabled",
            "Required",
            "Track Inventory"
        ]
        for variant in self._variants:
            # Booleans should be Y|N, but we sometimes see "".
            for boolean in booleans:
                if not variant[boolean] in ("Y", "N"):
                    variant[boolean] = "N"

    def get_variants(self):
        """Return a list of per-variant dictionaries."""

        if self._variants is None:
            filename = os.path.join(self._cache_dir, "variants.csv")

            with lockfile.FileLock(self._download_lock_filename):
                # Download products file if it is out of date.
                if self._is_file_expired(filename):
                    self._download_variants_csv(filename)

                # Read variants file.
                self._variants = list(csv.DictReader(open(filename)))

                # Cleanup suspect data.
                if self._clean:
                    self._clean_variants()

        return self._variants

    def get_questions(self):
        """Return a list of per-question dictionaries."""

        # Variant values that are the same for all answers.
        question_copy_keys = (
            "Answer Input Type",
            "Exclude from best seller report",
            "In Line Help",
            "Product Name",
            "Product SKU",
            "Question Enabled",
            "Question Sort Order",
            "Required",
            "Track Inventory"
        )

        if self._questions is None:
            # Process variants sorted by product, question.
            variants = sorted(
                self.get_variants(),
                key=lambda variant: (
                    variant["Product Id"],
                    variant["Question ID|Answer ID"].split("|")[0]
                )
            )

            self._questions = list()
            prev_product_id = None
            prev_question_id = None
            for variant in variants:
                product_id = variant["Product Id"]
                question_id = variant["Question ID|Answer ID"].split("|")[0]
                is_first_answer = (
                    product_id != prev_product_id or
                    question_id != prev_question_id
                )
                if is_first_answer:
                    question = {
                        "Product Id": product_id,
                        "Question ID": question_id,
                        "Question": variant["Question|Answer"].split("|")[0],
                        "_n_answers": 1
                    }
                    for key in question_copy_keys:
                        question[key] = variant[key]
                    self._questions.append(question)
                    prev_product_id = product_id
                    prev_question_id = question_id
                else:
                    self._questions[-1]["_n_answers"] += 1

        return self._questions

    def _download_products_csv(self, filename):
        """Download products list to a CSV file."""

        # Login if necessary.
        self._login()

        # Log time consuming step.
        LOGGER.info("Downloading products")

        # Load the export page.
        url = (
            self._admin_url +
            "?m=ajax_export&instance=products&checkAccess=products"
        )
        self._browser.open(url)

        # Select form.
        self._select_form("jsform")

        # Ensure that "All Categories" is selected.
        category_list = self._browser.form.find_control("category")
        if False:  # debug
            for item in category_list.items:
                print(
                    " name={} values={}".format(
                        item.name,
                        str(label.text for label in item.get_labels())
                    )
                )
        category_list.value = [""]  # name where values = ["All Categories"]

        # Submit the form (press the "Export" button).
        resp = self._browser.submit()
        if False:  # debug
            # Examine the source of the doExport method.
            print("Response from {}:\n".format(url))
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

        if self._products is None:
            filename = os.path.join(self._cache_dir, "products.csv")

            with lockfile.FileLock(self._download_lock_filename):
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

    def get_product_update_keys(self):
        """Return list of product keys where the value can be updated."""
        return self._PRODUCT_KEY_MAP.keys()

    def update_product(self, sku, key, value):
        """Login to site."""

        # Login if necessary.
        self._login()

        # Log time consuming step.
        LOGGER.info(
            "Updating product SKU={}, setting {} to {}".format(
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

        # print [item.name for item in form.find_control("useFile").items]
        # Set the form values.
        # self._browser["instance"] = "product_import"
        # self._browser["xsubmit"] = "true"
        # self._browser["file"] = "cctools.csv"
        # self._browser["useFile"] = ["Y",]
        named_tfile = tempfile.NamedTemporaryFile(
            mode="w",
            suffix=".csv",
            delete=False
        )
        with named_tfile.file as tfile:
            tfile.write("SKU,{}\n{},{}\n".format(key, sku, value))
        with open(named_tfile.name, mode="w") as tfile:
            self._browser.form.add_file(
                tfile,
                "text/csv",
                named_tfile.name,
                name="importFile"
            )
        # self._browser["importFile"] =\
        #     "SKU,{}\n{},{}\n".format(key, sku, value)
        self._browser["updateType"] = "update"

        # Submit the form (press the "????" button).
        self._browser.submit()

        # https://www16.corecommerce.com/~cohu1/admin/index.php?\
        #     m=ajax_import&instance=product_import
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
        # self._browser["instance"] = "product_import"
        self._browser["go"] =\
            self._admin_url + "?m=ajax_import&instance=product_import"
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

        # POST https://www16.corecommerce.com/~cohu1/admin/index.php?\
        #          m=ajax_import_save&instance=product_import
        #   m:           ajax_import_save
        #   go:          https://www16.corecommerce.com/~cohu1/admin/index.php\
        #                    ?m=ajax_import&\
        #                    instance=product_import
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

        # Log time consuming step.
        LOGGER.info("Downloading categories")

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

        if self._categories is None:
            filename = os.path.join(self._cache_dir, "categories.csv")

            with lockfile.FileLock(self._download_lock_filename):
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
        if self._category_sort is None:
            if self._categories is None:
                self.get_categories()
            self._category_sort = dict()
            for category in self._categories:
                name = category["Category Name"]
                sort = int(category["Sort"])
                self._category_sort[name] = sort

    def product_key_by_cat_and_name(self, product):
        """Return a key for a product dictionary used to sort by
        category, product_name.
        """
        if self._category_sort is None:
            self._init_category_sort()
        category = product["Category"]
        if category in self._category_sort:
            category_sort_key = "{:05d}".format(self._category_sort[category])
        else:
            category_sort_key = category
        return "{}:{}".format(category_sort_key, product["Product Name"])

    def product_key_by_category(self, product):
        """Return a key for a product dictionary used to sort by
        category, product_name.
        """
        if self._category_sort is None:
            self._init_category_sort()
        category = product["Category"]
        if category in self._category_sort:
            category_sort_key = "{:05d}".format(self._category_sort[category])
        else:
            category_sort_key = category
        return category_sort_key

    def product_key_by_sku(self, product):
        """Return a key for a product dictionary used to sort by sku."""
        # pylint: disable=R0201
        return product["SKU"]

    def variant_key(self, variant):
        """Return a key to sort variants."""
        # pylint: disable=R0201
        return (
            variant["Product SKU"],
            variant["Question Sort Order"],
            variant["Answer Sort Order"]
        )

    def variant_key_by_cat_product(self, variant):
        """Return a key for a variant dictionary used to sort by
        category, product_name, question, answer.
        """
        if self._category_sort is None:
            self._init_category_sort()
        self.get_products()

        category = None
        product_sort_key = None
        for product in self._products:
            if (
                product["Product Name"] == variant["Product Name"] and
                product["SKU"] == variant["Product SKU"]
            ):
                category = product["Category"]
                if category in self._category_sort:
                    category_sort_key = "{:05d}".format(
                        self._category_sort[category]
                    )
                else:
                    category_sort_key = category
                product_sort_key = product["Product Name"]
                break

        return (
            category_sort_key,
            product_sort_key,
            variant["Question Sort Order"],
            variant["Answer Sort Order"]
        )

    def question_key(self, question):
        """Return a key to sort questions."""
        # pylint: disable=R0201
        return (
            question["Product SKU"],
            question["Question Sort Order"]
        )

    def guess_product_ids(self):
        """
        The product list returned by CoreCommerce does not include product
        IDs.  Guess the ones we can.
        """

        # Download products and variants.
        self.get_products()
        self.get_variants()

        # Guess an ID for each product.
        for product in self._products:
            # Check to see if the product already has an ID.  As of
            # 2014-12-27, a product never has a Product Id, but maybe
            # that will change in the future.
            if "Product Id" in product and product["Product Id"] != "":
                continue

            # Find a variant that matches in name and SKU.
            product_id = ""
            name = product["Product Name"]
            sku = product["SKU"]
            for variant in self._variants:
                if (
                    variant["Product Name"] == name and
                    variant["Product SKU"] == sku
                ):
                    product_id = variant["Product Id"]
                    break

            # Assign the guessed ID to the product.
            product["Product Id"] = product_id


_HTML_TO_PLAIN_TEXT_DICT = {
    "&quot;": "\"",
    "&amp;": "&",
    "<p>": " ",
    "</p>": " "
}
_HTML_TO_PLAIN_TEXT_RE = re.compile("|".join(_HTML_TO_PLAIN_TEXT_DICT.keys()))


def html_to_plain_text(string):
    """Convert HTML markup to plain text."""

    # Replace HTML markup with plain text."""
    string = _HTML_TO_PLAIN_TEXT_RE.sub(
        lambda m: _HTML_TO_PLAIN_TEXT_DICT[m.group(0)],
        string
    )

    # Collapse all whitespace to a single space.
    string = re.sub(r"\s+", " ", string)

    # Strip leading and trailing whitespace.
    string = string.strip()

    return string


def plain_text_to_html(string):
    """Convert plain text to HTML markup."""
    string = string.replace("&", "&amp;")
    return string
