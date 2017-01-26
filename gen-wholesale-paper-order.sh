#!/bin/sh

usage() {
    echo "Generate and display YYYY-MM-DD-WholesalePaperOrder.pdf" >&2
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

SCRIPT=`readlink -f "$0"`
SCRIPTDIR=`dirname "${SCRIPT}"`

# Determine the output filename.
FILENAME="`date +%Y-%m-%d`-WholesaleOrder.pdf"
if [ "${OUTPUT_DIR}" ]; then
    FILENAME="${OUTPUT_DIR}/${FILENAME}"
fi
rm -f ${FILENAME}

# Generate the order form.  The "--exclude-sku" options indicate which
# products should be excluded.
#
# SKU 30001 = All other Earrings
"${SCRIPTDIR}/gen-wholesale-paper-order.py"\
    --category="Necklaces"\
    --category="Bags & Purses"\
    --category="Earrings"\
    --category="Bracelets"\
    --category="Baskets, Trivets & Bowls"\
    --category="Miscellaneous"\
    --exclude-sku=30001\
    --pdf-file="${FILENAME}"\
    --verbose\
    ${ARGS}

# Display the order form if it was successfully created.
if [ -f "${FILENAME}" ]; then
    # Enable monitor mode (job control).  If not set, evince will exit
    # when this shell exits.
    set -m
    evince "${FILENAME}" &
else
    echo "ERROR: '${FILENAME}' not created" >&2
    read -p "Press [Enter] to continue..." key
fi
