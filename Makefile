BASE := $(shell dirname $(realpath $(lastword $(MAKEFILE_LIST))))

SHELL := /bin/bash -o errexit -o pipefail

ifdef VERBOSE
	Q =
	E = @true 
	CAPTURE := 
else
	Q = @
	E = @echo 
	CAPTURE := $(shell command -v chronic 2>/dev/null) 
endif

all: help

help:
	$(Q)awk -F ':|##' \
		'/^[^\t].+?:.*?##/ {\
			printf "\033[36m%-30s\033[0m %s\n", $$1, $$NF \
		}' $(MAKEFILE_LIST)

clean: ## clean working dir
	$(E)Cleaning...
	$(Q)find $(BASE) -name '*.pyc' \
		-o -name '*.orig' \
		-o -name '*_BACKUP_*' \
		-o -name '*_BASE_*' \
		-o -name '*_LOCAL_*' \
		-o -name '*_REMOTE_*' | xargs rm 2>/dev/null || true
	$(Q)rm -rf "$(BASE)/dist"
	$(Q)rm -rf $(BASE)/htmlcov
	$(Q)rm -rf $(BASE)/.coverage

deps: ## install requirements
	$(E)Handling requirements...
	$(Q)$(CAPTURE)pip install -r "$(BASE)/requirements.txt"
	$(E)Installing test deps...
	$(Q)$(CAPTURE)pip install -r "$(BASE)/test-requirements.txt"

test: ## run tests for current version and all envs
	$(E)Testing all...
	$(Q)tox -c "$(BASE)/tox.ini"

.PHONY: all help clean deps test

.SUFFIXES:
