#!/bin/sh

usage() {
    echo "Generate and display YYYY-MM-DD-WholesalePrices.xlsx" >&2
    echo "" >&2
    echo "USAGE:" >&2
    echo "  $0 [options]" >&2
    echo "" >&2
    echo "OPTIONS:" >&2
    echo "  --dir=DIR = specify output directory" >&2
}

ARGS=""
OUTPUT_DIR=""
for ARG in "$@"; do
    case ${ARG} in
        --help)
            usage
            exit 1
            ;;
        --dir=*)
            OUTPUT_DIR="${ARG#*=}"
            ;;
        *)
            ARGS="${ARGS} ${ARG}"
            ;;
    esac
done

SCRIPT=`readlink -k "$0"`
SCRIPTDIR=`dirname "${SCRIPT}"`

# Determine the output filename.
FILENAME="`date +%Y-%m-%d`-WholesaleLineSheet.xlsx"
if [ "${OUTPUT_DIR}" ]; then
    FILENAME="${OUTPUT_DIR}/${FILENAME}"
fi
rm -f ${FILENAME}

# Generate the line sheet.  The "--exclude-sku" options indicate which
# products should be excluded.
#
# SKU 30001 = All other Earrings
"${SCRIPTDIR}/gen-wholesale-line-sheet.py"\
    --outfile="${FILENAME}"\
    --exclude-sku=30001\
    --verbose\
    ${ARGS}

# Display the line sheet if it was successfully created.
if [ -f "${FILENAME}" ]; then
    # Enable monitor mode (job control).  If not set, localc will exit
    # when this shell exits.
    echo "Opening ${FILENAME}"
    set -m
    localc "${FILENAME}" &
else
    echo "ERROR: '${FILENAME}' not created"
    read -p "Press [Enter] to continue..." key
fi
