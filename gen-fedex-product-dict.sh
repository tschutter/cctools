#!/bin/sh

usage() {
    echo "Generate and display YYYY-MM-DD-PurchaseOrder.xlsx" >&2
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
BASENAME="`date +%Y-%m-%d`-FedExProductDict"
FILENAME="${BASENAME}.xlsx"
if [ "${OUTPUT_DIR}" ]; then
    FILENAME="${OUTPUT_DIR}/${FILENAME}"
fi

# Generate the product dictionary.
# SKU 30001 = All other Earrings
"${SCRIPTDIR}/gen-fedex-product-dict.py"\
    --outfile="${FILENAME}"\
    --exclude-sku=30001\
    --verbose\
    ${ARGS}

# Display the product dictionary if it was successfully created.
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
