cctools
=======

cctools is a set of utilities that work with CoreCommerce product
data.

calc-online-price.py
    Calculates an online price (pre-tax) based upon a tax-included
    price.

ccc
    Core Commerce Command line tool.  Interact with CoreCommerce from
    the command line.  List categories, products, and
    personalizations.

cclint.py
    Detects problems in CoreCommerce product data.

gen-art-mart-checkin.py
    Generates an Art Mart Inventory Sheet.

gen-inventory.py
    Generates an inventory report in spreadsheet form.

gen-po-invoice.py
    Generates a Purchase Order / Commercial Invoice in XLSX form.

gen-price-list.py
    Generates a price list from CoreCommerce data in PDF form.  Prices
    are adjusted to include sales tax and are rounded to even dollar
    amounts.  The intent is to use the price list at fairs and shows
    to avoid the handling of change.  It also makes accounting of cash
    and checks easier because you can deal with round numbers.

gen-wholesale-order.py
    Generates a wholesale order form in spreadsheet form.

All of these tools use a common cctools.cfg configuration file.  The
configuration file uses `INI syntax
<http://docs.python.org/2/library/configparser.html>`_.  The config
file must have a [website] section that specifies how the tools can
login to CoreCommerce to download product and category lists::

    [website]
    host: www16.corecommerce.com
    site: yoursite
    username: cctools
    password: super!secret

I strongly suggest that you create a cctools limited user in
CoreCommerce that only has access to Products/Categories.

Notes
-----

As of 2012-12-20, when CoreCommerce exports boolean values, the values
are in the set {"Y", "N", ""}.  `According to Megan Heikkinen @
CoreCommerce
<https://getsatisfaction.com/corecommerce/topics/when_exporting_products_what_does_a_space_for_discontinued_item_mean>`_,
for boolean values a blank space is essentially the same as "N"
meaning that that particular setting/feature is off/not being used.
cctools.py therefore translates all boolean values that are not "Y" to
"N".

The inline HTML comments in the inline JavaScript are to prevent older
browsers that do not understand the script element from displaying the
script code in plain text::

    <script type="text/javascript">
        <!--
        // JavaScript code here
        -->
    </script>
