#!/bin/sh
#
# Generate and display YYYY-MM-DD-PurchaseOrder.xlsx
#

SCRIPT=`readlink --canonicalize "$0"`
SCRIPTDIR=`dirname "${SCRIPT}"`

# Determine the output filename and PO number.
BASENAME="`date +%Y-%m-%d`-PurchaseOrder"
BASENUM="`date +%y%m%d`"
FILENAME="${BASENAME}.xlsx"
NUMBER="${BASENUM}00"
if [ -f "${FILENAME}" ]; then
    for NUM in `seq --format "%02.0f" 1 99`; do
        FILENAME="${BASENAME}-r${NUM}.xlsx"
        NUMBER="${BASENUM}${NUM}"
        if [ ! -f "${FILENAME}" ]; then
            break
        fi
    done
fi

# Generate the po/invoice.
# SKU 30001 = All other Earrings
"${SCRIPTDIR}/gen-po-invoice.py"\
    --number="${NUMBER}"\
    --outfile="${FILENAME}"\
    --exclude-sku=30001\
    --verbose\
    "$@"

# Display the po/invoice if it was successfully created.
if [ -f "${FILENAME}" ]; then
    # Enable monitor mode (job control).  If not set, localc will exit
    # when this shell exits.
    set -m
    localc "${FILENAME}" &
else
    echo "ERROR: '${FILENAME}' not created"
    read -p "Press [Enter] to continue..." key
fi
