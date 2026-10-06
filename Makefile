PROJECT_NAME ?= Squad Health Check

env:
	@echo Installing @google/clasp...
	@npm install @google/clasp
	@echo Installed clasp v$(shell npx @google/clasp --version)

format:
	npx prettier --write .

login:
	npx @google/clasp login

project:
	npx @google/clasp create --title "$(PROJECT_NAME)" --type sheets

pull:
	npx @google/clasp pull

push:
	npx @google/clasp push
