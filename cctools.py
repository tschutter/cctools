#!/usr/bin/env python2

"""
Web scraper interface to CoreCommerce.

Product Variants
================

Product variants are a combination of personalizations and options.
Both can exist simultaneously.

Product Options
---------------

Product options are downloaded as a single CSV file, even though there
should be 8 separate tables.  If the schema was truly normalized,
those tables would probably be:

*) Option (OptionId, OptionName, OptionSort)
*) OptionGroup (
       OptionGroupId, OptionGrouopName,
       FirstOptionValue, UseFirstOptionValue
   )
*) OptionSet (ProductId, OptionSetSKU, Price, Cost, other?)
*) OptionXOptionGroup (OptionId, OptionGroupId)
*) OptionXOptionSet (OptionId, ProductId)
*) OptionSetXOptionGroup (ProductId, OptionGroupId)

There can only be one OptionSet per Product, so the ProductId is
equivalent to a OptionSetId.
"""

from __future__ import print_function
import csv
import json
import logging
import os
import re
import tempfile
import time

import lockfile  # sudo apt-get install python-lockfile
import mechanize  # sudo apt-get install python-mechanize
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

# pylint seems to be confused by calling methods via self._browser.
# pylint: disable=E1102

LOGGER = logging.getLogger(__name__)


def uniquify_header(line):
    """Make names unique in a CSV header line."""
    fields = line.strip().split(",")
    for i, field in enumerate(fields):
        if field in fields[i + 1:]:
            count = 1
            for j in range(i, len(fields)):
                if fields[j] == field:
                    fields[j] = "{} [{}]".format(
                        field,
                        count
                    )
                    count += 1

    return ",".join(fields) + "\n"


def repair_product_options_csv(fixed_filename, tmp_filename):
    """
    Repairs product_options CSV header which has duplicate keys.  Who exports
    a CSV file with identical column names?  That is criminal.  Each
    row consists of data about the "option set", followed by a number
    of option groups.  Each option group has 4 fields for the group,
    followed by 3 fields for the option that is in the option set and
    the option group.
    """

    with open(tmp_filename) as tmp_file:
        with open(fixed_filename, "w") as fixed_file:
            first = True
            for line in tmp_file.readlines():
                if first:
                    line = uniquify_header(line)
                    first = False
                fixed_file.write(line)


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
        self._browser.set_handle_robots(False)
        if proxy is not None:
            self._browser.set_proxies({"https": proxy})
        self._logged_in = False
        self._personalizations = None
        self._product_options = None
        self._option_sets = None
        self._option_groups = None
        self._options = None
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
            # If the form is not found, and the form list only
            # contains "digiSHOP", then the hostname is probably
            # wrong.  "digiSHOP" is on a redirected login page.
            # Change the requested name to "digiSHOP" and then print
            # self._browser.form to see clues.
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
        LOGGER.info("Logging into {}".format(self._admin_url))
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

    def _download_personalizations_csv(self, filename):
        """Download personalization list to a CSV file."""

        # Login if necessary.
        self._login()

        # Log time consuming step.
        LOGGER.info("Downloading personalizations")

        # Load export page.
        url = "{}?{}{}".format(
            self._admin_url,
            "m=ajax_export",
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

        if self._personalizations is None:
            filename = os.path.join(self._cache_dir, "personalizations.csv")

            with lockfile.FileLock(self._download_lock_filename):
                # Download products file if it is out of date.
                if self._is_file_expired(filename):
                    self._download_personalizations_csv(filename)

                # Read personalizations file.
                self._personalizations = list(csv.DictReader(open(filename)))

                # Cleanup suspect data.
                if self._clean:
                    self._clean_personalizations()

        return self._personalizations

    def _download_product_options_csv(self, filename):
        """Download product_option list to a CSV file."""

        # Login if necessary.
        self._login()

        # Log time consuming step.
        LOGGER.info("Downloading product_options")

        # Load export page.
        url = "{}?{}{}".format(
            self._admin_url,
            "m=ajax_export",
            "&instance=product_options&checkAccess=products"
        )
        self._browser.open(url)

        # Call the doExport function.
        self._do_export(url, filename)

    def _clean_product_options(self):
        """Normalize suspect product_option data."""
        # Boolean value of "" appears to mean "N".
        for product_option in self._product_options:
            # Booleans should be Y|N, but we sometimes see "".
            for key, value in product_option.items():
                if key and key.startswith("Use First Option Value"):
                    if value not in ("Y", "N"):
                        product_option[key] = "N"

    def get_product_options(self):
        """Return a list of per-product_option dictionaries."""

        if self._product_options is None:
            filename = os.path.join(self._cache_dir, "product_options.csv")

            with lockfile.FileLock(self._download_lock_filename):
                # Download products file if it is out of date.
                if self._is_file_expired(filename):
                    self._download_product_options_csv(filename + ".tmp")
                    repair_product_options_csv(filename, filename + ".tmp")

                # Read product_options file.
                # The header has 47 fields, but each data record has
                # 48 fields.  By setting restkey to "Extra", we
                # prevent the extra field from having a key of None.
                self._product_options = list(
                    csv.DictReader(open(filename), restkey="Extra")
                )

                # Cleanup suspect data.
                if self._clean:
                    self._clean_product_options()

        return self._product_options

    def get_option_sets(self):
        """
        Return a list of option sets.  The list is derived from the
        product options list.
        """

        if self._option_sets is None:
            # Product option values that belong in the option set.
            copy_keys = (
                "Product SKU",
                "Product Id",
                "Product Name"
            )

            # Process product options sorted by product, option set sku.
            product_options = sorted(
                self.get_product_options(),
                key=lambda product_option: (
                    product_option["Product Id"],
                    product_option["Option Set SKU"]
                )
            )

            self._option_sets = []
            for product_option in product_options:
                option_set = {}
                for key, value in product_option.items():
                    if key in copy_keys or key.startswith("Option Set "):
                        option_set[key] = value
                self._option_sets.append(option_set)

        return self._option_sets

    def get_option_groups(self):
        """
        Return a list of option groups.  The list is derived from the
        product options list.
        """

        if self._option_groups is None:
            # Product option values that belong in the option group.
            copy_option_set_keys = (
                "Product SKU",
                "Product Id",
                "Product Name",
            )
            copy_option_group_keys = (
                "Option Group Name",
                "First Option Value",
                "Use First Option Value"
            )

            # Process product options sorted by product, option group sku.
            product_options = sorted(
                self.get_product_options(),
                key=lambda product_option: (
                    product_option["Product Id"],
                    product_option["Option Set SKU"]
                )
            )

            option_group_ids = []
            self._option_groups = []
            for product_option in product_options:
                for idx in range(1, 10):
                    group_id_key = "Option Group Id [{}]".format(idx)
                    if group_id_key not in product_option:
                        break
                    option_group_id = product_option[group_id_key]
                    if option_group_id in option_group_ids:
                        continue
                    option_group_ids.append(option_group_id)
                    option_group = {}
                    option_group["Option Group Id"] = option_group_id
                    for key in copy_option_set_keys:
                        option_group[key] = product_option[key]
                    for key in copy_option_group_keys:
                        product_option_key = "{} [{}]".format(key, idx)
                        option_group[key] = product_option[product_option_key]
                    self._option_groups.append(option_group)

        return self._option_groups

    def get_options(self):
        """
        Return a list of options.  The list is derived from the product
        options list.
        """

        if self._options is None:
            # Product option values that belong in the option group.
            copy_option_set_keys = (
                "Product SKU",
                "Product Id",
                "Product Name",
            )
            copy_option_group_keys = (
                "Option Group Id",
                "Option Group Name"
            )
            copy_option_keys = (
                "Option Name",
                "Option Sort"
            )

            # Process product options sorted by product, option group sku.
            product_options = sorted(
                self.get_product_options(),
                key=lambda product_option: (
                    product_option["Product Id"],
                    product_option["Option Set SKU"]
                )
            )

            option_ids = []
            self._options = []
            for product_option in product_options:
                for idx in range(1, 10):
                    option_id_key = "Option Id [{}]".format(idx)
                    if option_id_key not in product_option:
                        break
                    option_id = product_option[option_id_key]
                    if option_id in option_ids:
                        continue
                    option_ids.append(option_id)
                    option = {}
                    option["Option Id"] = option_id
                    for key in copy_option_set_keys:
                        option[key] = product_option[key]
                    for key in copy_option_group_keys:
                        product_option_key = "{} [{}]".format(key, idx)
                        option[key] = product_option[product_option_key]
                    for key in copy_option_keys:
                        product_option_key = "{} [{}]".format(key, idx)
                        option[key] = product_option[product_option_key]
                    self._options.append(option)

        return self._options

    def get_variants(self):
        """Return a list of per-variant dictionaries."""

        if self._variants is None:
            self._variants = list()

            personalization_keys = {
                "Product SKU": "Product SKU",
                "Product Name": "Product Name",
                # "Question ID|Answer ID",
                # "Question|Answer",
                # "Answer Input Type",
                # "In Line Help",
                # "Max Characters",
                # "Required",
                # "Track Inventory",
                # "Question Sort Order",
                # "Exclude from best seller report",
                # "Required Quantity",
                "Size": "Variant Size",
                "Price": "Variant Add Price",
                "SKU": "Variant SKU",
                # "Swatch Image",
                # "Swatch Image Alt / Title Tag",
                # "Swatch Image Image URL Link",
                # "Swatch Image Height",
                # "Swatch Image Width",
                # "Swatch Image New Window URL",
                "Main Photo": "Variant Main Photo (Image)",
                "Main Photo Alt / Title Tag":
                    "Variant Main Photo (Alt / Title Tag)",
                "Main Photo Caption": "Variant Main Photo (Caption)",
                "Main Photo Image URL Link": "Variant Main Photo URL",
                "Main Photo Height": "Variant Main Photo URL Height",
                "Main Photo Width": "Variant Main Photo URL Width",
                "Large Photo": "Variant Large Pop-Up Photo (Image)",
                "Large Photo Alt / Title Tag":
                    "Variant Large Pop-Up Photo (Alt / Title Tag)",
                "Large Photo Image URL Link": "Variant Large Popup Photo URL",
                "Large Photo Height": "Variant Large Popup Photo URL Height",
                "Large Photo Width": "Variant Large Popup Photo URL Width",
                "Default": "Variant Default",
                # "Price Type",
                "Inventory Level": "Variant Inventory Level",
                "Low Inventory Notify Level": "Variant Notify Level",
                "Weight": "Variant Weight",
                "Cost": "Variant Add Cost"
            }
            product_option_keys = (
                "Product SKU",
                "Product Name",
                "Option Set SKU",
                "Option Set Price",
                "Option Set Weight",
                "Option Set Cost",
                "Option Set MSRP",
                "Option Set Main Photo (Image)",
                "Option Set Main Photo (Caption)",
                "Option Set Main Photo (Alt / Title Tag)",
                "Option Set Main Photo URL",
                "Option Set Main Photo URL Width",
                "Option Set Main Photo URL Height",
                "Option Set Large Pop-Up Photo (Image)",
                "Option Set Large Pop-Up Photo (Alt / Title Tag)",
                "Option Set Large Popup Photo URL",
                "Option Set Large Popup Photo URL Width",
                "Option Set Large Popup Photo URL Height",
                "Option Set Inventory Level",
                "Option Set Notify Level"
            )

            # Get source lists.
            products = self.get_products()
            personalizations = self.get_personalizations()
            product_options = self.get_product_options()

            for product in products:
                product_name = product["Product Name"]
                product_sku = product["SKU"]

                # Convert personalizations to variants.
                relevant_personalizations = [
                    personalization for personalization in personalizations
                    if (
                        personalization["Product Name"] == product_name and
                        personalization["Product SKU"] == product_sku
                    )
                ]
                if len(relevant_personalizations) > 0:
                    for personalization in relevant_personalizations:
                        variant = {
                            "Variant Type": "Personalization"
                        }
                        for p_key, v_key in personalization_keys.items():
                            variant[v_key] = personalization[p_key]
                        question_answer = personalization["Question|Answer"]
                        (
                            variant["Variant Group"],
                            variant["Variant Name"]
                        ) = question_answer.split("|")
                        variant["Variant Sort"] = personalization[
                            "Answer Sort Order"
                        ]
                        if (
                            personalization["Answer Enabled"] == "Y" and
                            personalization["Answer Enabled"] == "Y"
                        ):
                            variant["Variant Enabled"] = "Y"
                        else:
                            variant["Variant Enabled"] = "N"
                        self._variants.append(variant)

                # Convert option sets to variants.
                relevant_product_options = [
                    product_option for product_option in product_options
                    if (
                        product_option["Product Name"] == product_name and
                        product_option["Product SKU"] == product_sku
                    )
                ]
                if len(relevant_product_options) > 0:
                    for product_option in relevant_product_options:
                        variant = {
                            "Variant Type": "Option"
                        }
                        for key in product_option_keys:
                            variant_key = key.replace(
                                "Option Set", "Variant"
                            ).replace(
                                "Cost", "Add Cost"
                            ).replace(
                                "Price", "Add Price"
                            )
                            variant[variant_key] = product_option[key]
                        option_group_names = []
                        option_names = []
                        option_sorts = []
                        for idx in range(1, 10):
                            option_group_name_key = (
                                "Option Group Name [{}]".format(idx)
                            )
                            if option_group_name_key not in product_option:
                                break
                            option_group_names.append(
                                product_option[option_group_name_key]
                            )
                            option_name_key = "Option Name [{}]".format(idx)
                            option_names.append(
                                product_option[option_name_key]
                            )
                            option_sort_key = "Option Sort [{}]".format(idx)
                            option_sorts.append(
                                product_option[option_sort_key]
                            )
                        variant["Variant Group"] = ":".join(option_group_names)
                        variant["Variant Name"] = ":".join(option_names)
                        variant["Variant Sort"] = ":".join(option_sorts)
                        variant["Variant Enabled"] = "Y"
                        self._variants.append(variant)

        return self._variants

    def get_questions(self):
        """
        Return a list of per-question dictionaries.  The list is derived
        from the personalization list.
        """

        if self._questions is None:
            # Personalization values that are the same for all answers.
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

            # Process personalizations sorted by product, question.
            personalizations = sorted(
                self.get_personalizations(),
                key=lambda personalization: (
                    personalization["Product Id"],
                    personalization["Question ID|Answer ID"].split("|")[0]
                )
            )

            self._questions = list()
            prev_product_id = None
            prev_question_id = None
            for personalization in personalizations:
                product_id = personalization["Product Id"]
                question_id = personalization["Question ID|Answer ID"]
                question_id = question_id.split("|")[0]
                is_first_answer = (
                    product_id != prev_product_id or
                    question_id != prev_question_id
                )
                if is_first_answer:
                    question_name = personalization["Question|Answer"]
                    question_name = question_name.split("|")[0]
                    question = {
                        "Product Id": product_id,
                        "Question ID": question_id,
                        "Question": question_name,
                        "_n_answers": 1
                    }
                    for key in question_copy_keys:
                        question[key] = personalization[key]
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

    def personalization_key(self, personalization):
        """Return a key to sort personalizations."""
        # pylint: disable=R0201
        return (
            personalization["Product SKU"],
            int(personalization["Question Sort Order"]),
            int(personalization["Answer Sort Order"])
        )

    def personalization_key_by_cat_product(self, personalization):
        """Return a key for a personalization dictionary used to sort by
        category, product_name, question, answer.
        """
        if self._category_sort is None:
            self._init_category_sort()
        self.get_products()

        category = None
        product_sort_key = None
        for product in self._products:
            if (
                product["Product Name"] == personalization["Product Name"] and
                product["SKU"] == personalization["Product SKU"]
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
            int(personalization["Question Sort Order"]),
            int(personalization["Answer Sort Order"])
        )

    def product_option_key(self, product_option):
        """Return a key to sort product_options."""
        # pylint: disable=R0201
        return (
            product_option["Product SKU"],
            product_option["Option Set SKU"]
        )

    def product_option_key_by_cat_product(self, product_option):
        """
        Return a key for a product_option dictionary used to sort by
        category, product_name, question, answer.
        """
        if self._category_sort is None:
            self._init_category_sort()
        self.get_products()

        category = None
        product_sort_key = None
        for product in self._products:
            if (
                product["Product Name"] == product_option["Product Name"] and
                product["SKU"] == product_option["Product SKU"]
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
            int(product_option["Option Sort"])
        )

    def option_set_key(self, option_set):
        """Return a key to sort option_sets."""
        # pylint: disable=R0201
        return (
            option_set["Product Id"],
            option_set["Option Set SKU"]
        )

    def option_group_key(self, option_group):
        """Return a key to sort option_groups."""
        # pylint: disable=R0201
        return (
            option_group["Product Name"],
            option_group["Option Group Name"]
        )

    def option_key(self, option):
        """Return a key to sort options."""
        # pylint: disable=R0201
        return (
            option["Product Id"],
            option["Option Group Name"],
            int(option["Option Sort"])
        )

    def variant_key(self, variant):
        """Return a key to sort variants."""
        # pylint: disable=R0201
        return (
            variant["Product SKU"],
            variant["Variant Type"],
            variant["Variant Sort"]
        )

    def question_key(self, question):
        """Return a key to sort questions."""
        # pylint: disable=R0201
        return (
            question["Product SKU"],
            int(question["Question Sort Order"])
        )

    def guess_product_ids(self):
        """
        The product list returned by CoreCommerce does not include product
        IDs.  Guess the ones we can.
        """

        # Download products and personalizations.
        self.get_products()
        self.get_personalizations()

        # Guess an ID for each product.
        for product in self._products:
            # Check to see if the product already has an ID.  As of
            # 2014-12-27, a product never has a Product Id, but maybe
            # that will change in the future.
            if "Product Id" in product and product["Product Id"] != "":
                continue

            # Find a personalization that matches in name and SKU.
            product_id = ""
            name = product["Product Name"]
            sku = product["SKU"]
            for personalization in self._personalizations:
                if (
                    personalization["Product Name"] == name and
                    personalization["Product SKU"] == sku
                ):
                    product_id = personalization["Product Id"]
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
