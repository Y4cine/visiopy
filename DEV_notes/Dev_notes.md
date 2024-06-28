Install the library in dev mode:
pip install -e .

  git config --global user.email "yacine.gacem@gmail.com"
  git config --global user.name "Yacine Gacem"


python setup.py sdist bdist_wheel

twine upload dist/*

Upload to TestPyPI:
twine upload --repository-url https://test.pypi.org/legacy/ dist/*

Ctrl + K followed by Ctrl + C
Ctrl + K followed by Ctrl + U

pip install --force-reinstall --no-cache-dir -e .
pip install --no-cache-dir package_name

Helper functions:

Creating rows in the user and prop section, with all necessary tests.

batch processes for fast processing

transform shapesheet cells in properties

SetGetStates

SmartShapeManager

LayerManager

PythonVBA translater

