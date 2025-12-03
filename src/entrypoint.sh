#!/bin/sh

function parse_inputs {
    
    waf_file_or_dir="WAF"
    if [ "${INPUT_WAF_FILE_OR_DIR}" != "" ] || [ "${INPUT_WAF_FILE_OR_DIR}" != "." ]; then
        waf_file_or_dir="--root ${INPUT_WAF_FILE_OR_DIR}"
    fi
}

function main {

    scriptDir=$(dirname ${0})
    source ${scriptDir}/waf.sh
    parse_inputs
    
    waf
    
}

main "${*}"
