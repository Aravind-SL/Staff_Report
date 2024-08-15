.PHONY: run

install-dev: $(UIPY)
	pip install -e .

run: 
	python -m staffreport

install:
	pip install .

