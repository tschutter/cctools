#!/bin/sh

usage() {
    echo "Generate and display YYYY-MM-DD-BASENAME.xlsx" >&2
    echo "" >&2
    echo "USAGE:" >&2
    echo "  $0 [options]" >&2
    echo "" >&2
    echo "OPTIONS:" >&2
    echo "  --basename=NAME = filename part after date prefix" >&2
    echo "  --dir=DIR = output directory" >&2
    echo "  other options are passed to gen-wholesale-line-sheet.py"
    echo "  useful options include:"
    echo "    --price-multiplier"
    echo "    --price-precision"
}

ARGS=""
BASENAME="RetailLineSheet"
OUTPUT_DIR=""
for ARG in "$@"; do
    case ${ARG} in
        --help)
            usage
            exit 1
            ;;
        --basename=*)
            BASENAME="${ARG#*=}"
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
FILENAME="`date +%Y-%m-%d`-${BASENAME}.xlsx"
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
    --wholesale-fraction=1\
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
