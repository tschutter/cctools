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

SCRIPT=`readlink -f "$0"`
SCRIPTDIR=`dirname "${SCRIPT}"`

# Determine the output filename.
FILENAME="PriceListRetailTaxInc.pdf"
if [ "${OUTPUT_DIR}" ]; then
    FILENAME="${OUTPUT_DIR}/${FILENAME}"
fi
rm -f ${FILENAME}

# Generate the price list.  The "--exclude-sku" options indicate which
# products should be excluded.
#
# 40012 = Rosary
# 40050 = Mary Basket
# 40051 = Mary Basket Small
# 40073 = Cheik Earrings
# 70074 = Jarara Earrings
# 70075 = Lapeta-Cheik Earrings
# 70076 = Lapeta-Tino Earrings
# 70077 = Tino Earrings
# 70078 = Tino-Cheik Earrings
# 70079 = Jarara-Cheik Earrings
# 70085 = Extra Marketing Materials
"${SCRIPTDIR}/gen-price-list.py"\
    --category="Necklaces"\
    --category="Bags & Purses"\
    --category="Earrings"\
    --category="Bracelets"\
    --category="Baskets, Trivets & Bowls"\
    --category="Miscellaneous"\
    --exclude-sku=40012\
    --exclude-sku=40050\
    --exclude-sku=40051\
    --exclude-sku=40073\
    --exclude-sku=70074\
    --exclude-sku=70075\
    --exclude-sku=70076\
    --exclude-sku=70077\
    --exclude-sku=70078\
    --exclude-sku=70079\
    --exclude-sku=70085\
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
    echo "ERROR: '${FILENAME}' not created" >&2
    read -p "Press [Enter] to continue..." key
fi
