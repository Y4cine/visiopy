# visiopy

A library to automate Visio operations using Python.  
Initially specialized as terminal in a Jupyter Notebook for fast batch editing Visio drawings.

## Installation

pip install visiopy

## Usage
### List Open Documents
from visiopy import loaded_docs
loaded_docs()

### Initialize Visio Application
from visiopy import vInit
vInit(0, globals_dict=globals())
print(c.visSectionUser)
