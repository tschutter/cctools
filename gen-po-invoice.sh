#!/bin/sh
#
# Generate and display YYYY-MM-DD-PurchaseOrder.xlsx
#

SCRIPT=`readlink --canonicalize "$0"`
SCRIPTDIR=`dirname "${SCRIPT}"`

cd "${SCRIPTDIR}"

# Determine the output filename.
BASENAME="`date +%Y-%m-%d`-PurchaseOrder"
FILENAME="${BASENAME}.xlsx"
if [ -f "${FILENAME}" ]; then
    for NUM in `seq --format "%02.0f" 1 12`; do
        FILENAME="${BASENAME}-r${NUM}.xlsx"
        if [ ! -f "${FILENAME}" ]; then
            break
        fi
    done
fi

# Generate the po/invoice.
./gen-po-invoice.py --outfile="${FILENAME}" --verbose

# Display the po/invoice if it was successfully created.
if [ -f "${FILENAME}" ]; then
    localc "${FILENAME}"
fi
