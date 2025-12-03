#!/bin/sh

function parse_inputs {
    
    waf_file_or_dir=""
    if [ "${INPUT_WAF_FILE_OR_DIR}" != "" ] || [ "${INPUT_WAF_FILE_OR_DIR}" != "." ]; then
        waf_file_or_dir="--root ${INPUT_WAF_FILE_OR_DIR}"
    fi
    waf_comment=0
    if [[ "${INPUT_WAF_COMMENT}" == "0" || "${INPUT_WAF_COMMENT}" == "false" ]]; then
        waf_comment="0"
    fi

    if [[ "${INPUT_WAF_COMMENT}" == "1" || "${INPUT_WAF_COMMENT}" == "true" ]]; then
        waf_comment="1"
    fi
    waf_run_validate_yaml_content=0
    if [[ "${INPUT_WAF_RUN_VALIDATE_YAML_CONTENT}" == "0" || "${INPUT_WAF_RUN_VALIDATE_YAML_CONTENT}" == "false" ]]; then
        waf_run_validate_yaml_content=0
    fi
    if [[ "${INPUT_WAF_RUN_VALIDATE_YAML_CONTENT}" == "1" || "${INPUT_WAF_RUN_VALIDATE_YAML_CONTENT}" == "true" ]]; then
        waf_run_validate_yaml_content=1
    fi
    waf_run_create_excel_file=0
    if [[ "${INPUT_WAF_RUN_CREATE_EXCEL_FILE}" == "0" || "${INPUT_WAF_RUN_CREATE_EXCEL_FILE}" == "false" ]]; then
        waf_run_create_excel_file=0
    fi
    if [[ "${INPUT_WAF_RUN_CREATE_EXCEL_FILE}" == "1" || "${INPUT_WAF_RUN_CREATE_EXCEL_FILE}" == "true" ]]; then
        waf_run_create_excel_file=1
    fi
}

function main {

    scriptDir=$(dirname ${0})
    #source ${scriptDir}/waf.sh

    if [ ${waf_run_validate_yaml_content} -eq 1]; then
        source ${scriptDir}/waf-validate.sh
        parse_inputs
        waf-validate 
    fi
    if [ ${waf_run_create_excel_file} -eq 1]; then
        source ${scriptDir}/waf-excel.sh
        parse_inputs
        waf-excel
    fi
    
}

main "${*}"
