#!/bin/sh
#
# Generate and display PriceListRetailTaxInc.pdf
#

SCRIPT=`readlink --canonicalize "$0"`
SCRIPTDIR=`dirname "${SCRIPT}"`

# Determine the output filename.
FILENAME="PriceListRetailTaxInc.pdf"
rm -f ${FILENAME}

# Generate the price list.
"${SCRIPTDIR}/gen-price-list.py"\
    --exclude-sku=40025\
    --exclude-sku=40027\
    --exclude-sku=40043\
    --exclude-sku=40052\
    --exclude-sku=40064\
    --exclude-sku=40068\
    --exclude-sku=40073\
    --exclude-sku=70074\
    --exclude-sku=70075\
    --exclude-sku=70076\
    --exclude-sku=70077\
    --exclude-sku=70078\
    --exclude-sku=70079\
    --pdf-file="${FILENAME}"\
    --verbose\
    "$@"

# Display the price list if it was successfully created.
if [ -f "${FILENAME}" ]; then
    # Enable monitor mode (job control).  If not set, evince will exit
    # when this shell exits.
    set -m
    evince "${FILENAME}" &
else
    echo "ERROR: '${FILENAME}' not created"
    read -p "Press [Enter] to continue..." key
fi
