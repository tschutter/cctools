#!/bin/sh
#
# Generate and display YYYY-MM-DD-OnlineInventory.xlsx
#

SCRIPT=`readlink --canonicalize "$0"`
SCRIPTDIR=`dirname "${SCRIPT}"`

# Determine the output filename.
FILENAME="`date +%Y-%m-%d`-OnlineInventory.xlsx"
rm -f "${FILENAME}"

# Generate the inventory report.
"${SCRIPTDIR}/gen-inventory.py" --outfile="${FILENAME}" --verbose

# Display the inventory report if it was successfully created.
if [ -f "${FILENAME}" ]; then
    localc "${FILENAME}"
fi
