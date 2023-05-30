.PHONY: compile clean-build clean compile-force compile-profile install-local publish-pypi ruff flake8 cython-lint lint

.DEFAULT_GOAL := compile;

compile:
	@echo "Compiling...";
	python setup.py build_ext --inplace;
	@echo "Done.";

clean-build:
	@echo "Removing the build related directories...";
	rm -rf build dist kb_macropy.egg-info;
	rm -rf **/__pycache__;
	@echo "Done.";

clean: clean-build
	@echo "Removing the '.pyd', '.c' files, and the 'build' directory...";
	rm -rf src/cythonExtensions/**/*.pyd src/cythonExtensions/**/*.c;
	@echo "Done.";

compile-force:
	@echo "Forcing recompilation...";
	python setup.py build_ext --inplace --force;
	@echo "Done.";

compile-profile:
	@echo "Recompiling with profiling enabled...";
	python setup.py build_ext --inplace --force --profile;
	@echo "Done.";

install-local: clean-build
	@echo "Installing package from local...";
	pip uninstall kb_macropy -y;
	pip install .;
	@echo "Done.";

publish-pypi:
	@echo "Publishing to PyPI...";
	python setup.py sdist bdist_wheel;
	twine upload dist/*;
	@echo "Done.";

ruff:
	@echo "Linting Python files...";
	-ruff .;
	@echo "Done.";

flake8:
	@echo "Linting Python files...";
	-flake8 --color always;
	@echo "Done.";

cython-lint:
	@echo "Linting Cython files...";
	-cython-lint src/cythonExtensions/**/*.pyx --ignore W293,E501,E266,E265,E261,E221,E128,E127;
	@echo "Done.";

lint: ruff flake8 cython-lint;