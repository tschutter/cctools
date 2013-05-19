#!/bin/sh
#
# Generate and display YYYY-MM-DD-WholesalePrices.xlsx
#

SCRIPT=`readlink --canonicalize "$0"`
SCRIPTDIR=`dirname "${SCRIPT}"`

# Determine the output filename.
FILENAME="`date +%Y-%m-%d`-WholesaleOrder.xlsx"

# Generate the order form.
"${SCRIPTDIR}/gen-wholesale-order.py" --outfile="${FILENAME}" --verbose

# Display the order form if it was successfully created.
if [ -f "${FILENAME}" ]; then
    localc "${FILENAME}"
fi
