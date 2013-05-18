#!/bin/sh
#
# Generate and display YYYY-MM-DD-WholesalePrices.xlsx
#

SCRIPT=`readlink --canonicalize "$0"`
SCRIPTDIR=`dirname "${SCRIPT}"`

# Determine the output filename.
FILENAME="`date +%Y-%m-%d`-WholesalePrices.xlsx"

# Generate the po/invoice.
"${SCRIPTDIR}/gen-wholesale-price-list.py" --outfile="${FILENAME}" --verbose

# Display the po/invoice if it was successfully created.
if [ -f "${FILENAME}" ]; then
    localc "${FILENAME}"
fi
