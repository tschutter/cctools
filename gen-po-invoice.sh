#!/bin/sh
#
# Generate and display YYYY-MM-DD-PurchaseOrder.xlsx
#

SCRIPT=`readlink --canonicalize "$0"`
SCRIPTDIR=`dirname "${SCRIPT}"`

cd "${SCRIPTDIR}"
FILENAME="`date +%Y-%m-%d`-PurchaseOrder.xlsx"
./gen-po-invoice.py --outfile="${FILENAME}" --verbose
localc "${FILENAME}"
