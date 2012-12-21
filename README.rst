cctools
=======

CoreCommerce tools

gen-price-list.py
    Generates a price list from CoreCommerce data in PDF form.  Prices
    are adjusted to include sales tax and are rounded to even dollar
    amounts.  The intent is to use the price list at fairs and shows
    to avoid the handling of change.  It also makes accounting easier
    because you can deal with round numbers.

gen-art-mart-checkin.py
    Generates an Art Mart Inventory Sheet.

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
