#!/bin/sh

usage() {
    echo "Generate and display YYYY-MM-DD-PriceListRetailTaxInc.pdf" >&2
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

SCRIPT=`readlink --canonicalize "$0"`
SCRIPTDIR=`dirname "${SCRIPT}"`

# Determine the output filename.
FILENAME="PriceListRetailTaxInc.pdf"
if [ "${OUTPUT_DIR}" ]; then
    FILENAME="${OUTPUT_DIR}/${FILENAME}"
fi
rm -f ${FILENAME}

# Generate the price list.
"${SCRIPTDIR}/gen-price-list.py"\
    --category="Necklaces"\
    --category="Bags & Purses"\
    --category="Bracelets"\
    --category="Baskets, Trivets & Bowls"\
    --category="Earrings"\
    --category="Miscellaneous"\
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
    ${ARGS}

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
