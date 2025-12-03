FROM python:3.12-alpine

# hadolint ignore=DL3018
RUN apk add --update --no-cache bash ca-certificates curl git jq openssh

RUN pip install pyyaml

RUN ["bin/sh", "-c", "mkdir -p /src"]

COPY ["src", "/src/"]
RUN chmod -R +x /src

ENTRYPOINT ["/src/entrypoint.sh"]
