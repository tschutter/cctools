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
# SKU 30001 = All other Earrings
"${SCRIPTDIR}/gen-wholesale-order.py"\
    --outfile="${FILENAME}"\
    --exclude-sku=30001\
    --verbose\
    "$@"

# Display the order form if it was successfully created.
if [ -f "${FILENAME}" ]; then
    # Enable monitor mode (job control).  If not set, localc will exit
    # when this shell exits.
    set -m
    localc "${FILENAME}" &
else
    echo "ERROR: '${FILENAME}' not created"
    read -p "Press [Enter] to continue..." key
fi
