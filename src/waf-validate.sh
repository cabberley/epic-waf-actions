#!/bin/sh

function waf-validate {

    # gather output
    echo "waf: info: waf validate on ${waf_file_or_dir}."
    #python /src/validate_waf.py ${waf_file_or_dir} > lint_result.txt
    python /src/validate_waf.py > waf_validate_result.txt
    waf_exit_code=${?}

    # exit code 0 - success
    if [ ${waf_exit_code} -eq 0 ];then
        waf_comment_status="Success"
        echo "waf: info: successful yaml validation on ${waf_file_or_dir}."
        cat waf_validate_result.txt
        echo
    fi

    # exit code !0 - failure
    if [ ${waf_exit_code} -ne 0 ]; then
        waf_comment_status="Failed"
        echo "waf: error: failed validation on ${waf_file_or_dir}."
        cat waf_validate_result.txt
        echo
    fi

    # comment if lint failed
    if [ "${GITHUB_EVENT_NAME}" == "pull_request" ] && [ "${waf_comment}" == "1" ] && [ ${waf_exit_code} -ne 0 ]; then
        waf_comment_wrapper="#### \`waf\` ${waf_comment_status}
<details><summary>Show Output</summary>

\`\`\`
$(cat waf_validate_result.txt)
 \`\`\`
</details>

*Workflow: \`${GITHUB_WORKFLOW}\`, Action: \`${GITHUB_ACTION}\`, waf: \`${waf_file_or_dir}\`*"
    
        echo "waf: info: creating json"
        waf_payload=$(echo "${waf_comment_wrapper}" | jq -R --slurp '{body: .}')
        waf_comment_url=$(cat ${GITHUB_EVENT_PATH} | jq -r .pull_request.comments_url)
        echo "waf: info: commenting on the pull request"
        echo "${waf_payload}" | curl -s -S -H "Authorization: token ${GITHUB_ACCESS_TOKEN}" --header "Content-Type: application/json" --data @- "${waf_comment_url}" > /dev/null
    fi

    echo "waf_validate_output<<EOF" >> "$GITHUB_OUTPUT"
    cat waf_validate_result.txt >> "$GITHUB_OUTPUT"
    echo "EOF" >> "$GITHUB_OUTPUT"
    exit ${waf_exit_code}
}
