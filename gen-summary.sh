#!/bin/sh

usage() {
    echo "Generate and display a summary of interesting products" >&2
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
FILENAME="ProductSummary.txt"
if [ "${OUTPUT_DIR}" ]; then
    FILENAME="${OUTPUT_DIR}/${FILENAME}"
fi
rm -f ${FILENAME}

FIELDS="Category,Product Name,SKU,Avail,Discd,InvLvl,VInvLvl"

echo "Products that are viewable online and cannot be purchased online, but are available at shows:" >> "${FILENAME}"
"${SCRIPTDIR}/ccc" --notify-send list prod --fields "${FIELDS}" --filter Avail=Y --filter Discd=Y >> "${FILENAME}" ${ARGS}

echo "" >> "${FILENAME}"
echo "Products that are not viewable online, but are available at shows:" >> "${FILENAME}"
"${SCRIPTDIR}/ccc" --notify-send list prod --fields "${FIELDS}" --filter Avail=N --filter Discd=N >> "${FILENAME}" ${ARGS}

echo "" >> "${FILENAME}"
echo "Products that are not viewable online, and are not available at shows:" >> "${FILENAME}"
"${SCRIPTDIR}/ccc" --notify-send list prod --fields "${FIELDS}" --filter Avail=N --filter Discd=Y >> "${FILENAME}" ${ARGS}

echo "" >> "${FILENAME}"
date +"This report generated %Y-%m-%d %H:%M:%S." >> "${FILENAME}"

# Display the output file if it was successfully created.
if [ -f "${FILENAME}" ]; then
    # Enable monitor mode (job control).  If not set, evince will exit
    # when this shell exits.
    echo "Opening ${FILENAME}"
    set -m
    xdg-open "${FILENAME}" &
else
    echo "ERROR: '${FILENAME}' not created"
    read -p "Press [Enter] to continue..." key
fi
