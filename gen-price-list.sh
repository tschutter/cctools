#!/bin/sh
#
# Generate and display PriceListRetailTaxInc.pdf
#

SCRIPT=`readlink --canonicalize "$0"`
SCRIPTDIR=`dirname "${SCRIPT}"`

cd "${SCRIPTDIR}"
./gen-price-list.py --verbose
evince PriceListRetailTaxInc.pdf
