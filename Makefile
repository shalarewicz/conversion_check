linux:
	# install or upgrade pip
	python3 -m pip install --user --upgrade pip
	# check version of pip
	python3 -m pip --version
	# install virtual environmnets
	python3 -m pip install --user virtualenv
	# create a new virtual environment
	python3 -m venv env
	# activate env from command line
	@echo
	@echo copy and run the following command
	@echo
	@echo source env/bin/activate

windows:
	# check version of pip
	py -m pip --version
	# install or upgrade pip
	py -m pip install --upgrade pip
	# install virtual environments
	py -m pip install --user virtualenv
	# create a new environment
	python3 -m venv env python==3.5
	# activate env from command line
	@echo
	@echo copy and run the following command
	@echo
	@echo py -m venv env

freeze:
	pip freeze | grep -v env > requirements.txt






