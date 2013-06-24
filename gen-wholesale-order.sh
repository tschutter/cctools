#!/bin/sh
#
# Generate and display YYYY-MM-DD-WholesalePrices.xlsx
#

SCRIPT=`readlink --canonicalize "$0"`
SCRIPTDIR=`dirname "${SCRIPT}"`

# Determine the output filename.
FILENAME="`date +%Y-%m-%d`-WholesaleOrder.xlsx"
rm -f ${FILENAME}

# Generate the order form.
"${SCRIPTDIR}/gen-wholesale-order.py" --outfile="${FILENAME}" --verbose "$@"

# Display the order form if it was successfully created.
if [ -f "${FILENAME}" ]; then
    localc "${FILENAME}" &
else
    echo "ERROR: '${FILENAME}' not created"
    read -p "Press [Enter] to continue..." key
fi
