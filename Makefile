install-dev:
	pip install -e .

.PHONY: install-dev

run-dev: install-dev
	python -m staffreport


install:
	pip install .

