cctools
=======

``cctools`` is a set of utilities that work with CoreCommerce product
data.

``calc_price.py``
    Converts an event price (tax-included) to a retail price (pre-tax)
    or vice versa.

``ccc``
    Core Commerce Command line tool.  Interact with CoreCommerce from
    the command line.  List categories, products, and variants
    (personalizations).

``cclint.py``
    Detects problems in CoreCommerce product data.

``gen-art-mart-checkin.py``
    Generates an Art Mart Inventory Sheet.

``gen-inventory.py``
    Generates an inventory report in spreadsheet form.

``gen-po-comm-invoice.py``
    Generates a Purchase Order / Commercial Invoice in XLSX form.

``gen-price-list.py``
    Generates a price list from CoreCommerce data in PDF form.  Prices
    are adjusted to include sales tax and are rounded to even dollar
    amounts.  The intent is to use the price list at fairs and shows
    to avoid the handling of change.  It also makes accounting of cash
    and checks easier because you can deal with round numbers.

``gen-wholesale-line-sheet.py``
    Generates a wholesale line sheet in spreadsheet form.

``gen-wholesale-order.py``
    Generates a wholesale order form in spreadsheet form.

Installation
------------

``cctools`` uses the Python ``mechanize`` package.  Unfortunately as
of 2015-03-19, ``mechanize`` has not been ported to Python3.
Therefore you must install Python2.

Configuration
-------------

All of these tools use a common ``cctools.cfg`` configuration file.
The configuration file uses `INI syntax
<http://docs.python.org/2/library/configparser.html>`_.  The config
file must have a ``[website]`` section that specifies how the tools
can login to CoreCommerce to download product and category lists::

    [website]
    host: www16.corecommerce.com
    site: yoursite
    username: cctools
    password: super!secret

It is strongly suggested that you create a ``cctools`` limited user in
CoreCommerce that only has access to Products/Categories.

Linux
+++++

1) Install required packages::

    sudo apt-get install python-lockfile python-mechanize
    sudo apt-get install python-reportlab python-openpyxl

2) Clone the ``cctools`` repository::

    cd ~
    git clone https://github.com/tschutter/cctools.git

3) Create a ``cctools.cfg`` configuration file in the ``cctools``
   directory.

4) Create links on your Desktop to used ``.sh`` files.

Windows
+++++++

1) Install Python 2.7.9 or later to ensure that the pip package
   manager is installed.

2) Install required packages::

    pip install python-lockfile python-mechanize
    pip install python-reportlab python-openpyxl

3) Download the `ZIP file
   <https://github.com/tschutter/cctools/archive/master.zip>`_ of the
   ``cctools`` repository, and unzip it WHERE?

4) Create a ``cctools.cfg`` configuration file in the ``cctools``
   directory.

5) Create links on your desktop to frequently used tools.

Notes
-----

As of 2012-12-20, when CoreCommerce exports boolean values, the values
are in the set {"Y", "N", ""}.  `According to Megan Heikkinen @
CoreCommerce
<https://getsatisfaction.com/corecommerce/topics/when_exporting_products_what_does_a_space_for_discontinued_item_mean>`_,
for boolean values a blank space is essentially the same as "N"
meaning that that particular setting/feature is off/not being used.
``cctools.py`` therefore translates all boolean values that are not
"Y" to "N".

The inline HTML comments in the inline JavaScript are to prevent older
browsers that do not understand the script element from displaying the
script code in plain text::

    <script type="text/javascript">
        <!--
        // JavaScript code here
        -->
    </script>
