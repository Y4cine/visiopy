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

## Revision 2
The workflow is not straightforward enough.    
I often forget the name `loaded_docs`. `vDocs` is a better name. It aligns with `Init`.    
There should be no need to call `vDocs`. if `vInit` is called without arguments a dialog shall open, showing the loaded docs, additionally there should be a button to trigger a file picker to open a Visio file that is not yet loaded.    
As I often work with templates instead of drawings, handling vstx and vstm should be added to the scope.

