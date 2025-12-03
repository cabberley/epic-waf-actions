#!/bin/sh

function parse_inputs {
    
    waf_file_or_dir="WAF"
    if [ "${INPUT_WAF_FILE_OR_DIR}" != "" ] || [ "${INPUT_WAF_FILE_OR_DIR}" != "." ]; then
        waf_file_or_dir="--root ${INPUT_WAF_FILE_OR_DIR}"
    fi
    yamllint_comment=0
    if [[ "${INPUT_YAMLLINT_COMMENT}" == "0" || "${INPUT_YAMLLINT_COMMENT}" == "false" ]]; then
        yamllint_comment="0"
    fi

    if [[ "${INPUT_YAMLLINT_COMMENT}" == "1" || "${INPUT_YAMLLINT_COMMENT}" == "true" ]]; then
        yamllint_comment="1"
    fi
}

function main {

    scriptDir=$(dirname ${0})
    source ${scriptDir}/waf.sh
    parse_inputs
    
    waf
    
}

main "${*}"
