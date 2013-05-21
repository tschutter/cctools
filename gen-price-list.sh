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
    --exclude-category=Earrings\
    --pdf-file="${FILENAME}"\
    --verbose\
    "$@"

# Display the price list if it was successfully created.
if [ -f "${FILENAME}" ]; then
    # Enable monitor mode (job control).  If not set, evince will exit
    # when this shell exits.
    set -m
    evince "${FILENAME}" &
fi
