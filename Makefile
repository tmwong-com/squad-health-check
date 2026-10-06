PROJECT_NAME ?= Squad Health Check

# Development environment setup

.PHONY: env
env:
	@echo Installing @google/clasp...
	@npm install @google/clasp
	@echo Installed clasp v$(shell npx @google/clasp --version)

# Code sanity checks

.PHONY: format
format:
	npx prettier --write .

.PHONY: lint
lint:
	npm run format:check
	npm run lint

# Google Apps Script project management

.PHONY: login
login:
	npx @google/clasp login

.PHONY: project
project:
	npx @google/clasp create --title "$(PROJECT_NAME)" --type sheets

.PHONY: pull
pull:
	npx @google/clasp pull

.PHONY: push
push:
	npx @google/clasp push
