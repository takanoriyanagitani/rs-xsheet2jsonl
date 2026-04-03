#!/bin/sh

wsm="./rs-xsheet2jsonl.wasm"

ifile="./sample.d/input.xlsx"
gname="/guest.d/read-write.d/input.xlsx"

geninput(){
    export ifile="${ifile}"
    mkdir -p sample.d
    python3 geninput.py
}

test -f "${ifile}" || geninput

test -f "${ifile}" || exec env iname="${ifile}" sh -c '
    echo input xlsx file "${iname}" missing.
    exit 1
'

wasmtime run \
    --dir="${PWD}/sample.d::/guest.d/read-write.d" \
    "${wsm}" \
    --xlsx-path="${gname}" |
    jq -c
