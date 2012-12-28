#!/bin/sh
#
# Generate and display PriceListRetailTaxInc.pdf
#

SCRIPT=`readlink --canonicalize "$0"`
SCRIPTDIR=`dirname "${SCRIPT}"`

FILENAME="PriceListRetailTaxInc.pdf"

# Generate the price list.
"${SCRIPTDIR}/gen-price-list.py" --pdf-file="${FILENAME}" --verbose

# Display the price list if it was successfully created.
if [ -f "${FILENAME}" ]; then
    evince "${FILENAME}"
fi
