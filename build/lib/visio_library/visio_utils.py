"""
visio_utils.py

This module provides the base functionality to interact with Microsoft Visio, by providing handles on the application, the document, the page and the window.
It also provides the access to Visio constant names over the variable c.

Usage:
    from visio_utils import loaded_docs, vInit
    loaded_docs() # to get the list of open Visio drawings
    vInit(0, globals_dict=globals()) # to instantiate the variables vApp, vDoc, vPg, vWin and the visio constants
    # Use c for accessing the Visio constants
    print(c.visSectionUser)

For more information, visit the documentation at [link].
"""
import win32com.client
import pythoncom
import tkinter as tk
from tkinter import filedialog
import os

c = []  # to hold visio constants

# Function to list loaded documents


def loaded_docs(index=None):
    '''
    Prints the list of all open Visio drawings in all Visio instances and returns the list of the document objects.

    Parameter:
    - index: integer, optional makes the function return the document object with this index.
    '''
    global c
    visio_clsid = "{00021A21-0000-0000-C000-000000000046}"
    context = pythoncom.CreateBindCtx(0)
    rot = pythoncom.GetRunningObjectTable()
    docs = []

    for moniker in rot:
        try:
            name = moniker.GetDisplayName(context, None)
            if visio_clsid in name or name.endswith('.vsdx') or name.endswith('.vsdm'):
                visio_doc = moniker.BindToObject(
                    context, None, pythoncom.IID_IDispatch)
                visio_doc = win32com.client.Dispatch(visio_doc)
                docs.append(visio_doc)
                c = win32com.client.constants
        except Exception as e:
            print(f"Error processing moniker: {e}")

    if index is not None:
        if 0 <= index < len(docs):
            return docs[index]
        else:
            raise ValueError("Invalid document index.")

    for i, doc in enumerate(docs):
        print(f"{i}: {doc.FullName}")

    return docs

# Function to initialize Visio objects


def vInit(index=None, filename=None, new=False, template=None, globals_dict=None, suffix=None):
    """
    Initialize the Visio application and set global variables for vApp, vDoc, vPg, and vWin.

    Parameters:
    - index: Integer, index of the loaded docs.
    - filename: String, path to an existing Visio file to open.
    - new: Boolean, if True, create a new document (optionally with a template).
    - template: String, path to the template file for creating a new document.
    - suffix: String, for instantiating several documents. e.g.: vDoc1, vApp1, ...
    """
    if index is not None:
        doc = loaded_docs(index)
    elif filename is not None:
        doc = get_or_open_visio_file(filename)
    elif new:
        doc = create_new_document(template)
    else:
        print("Usage:")
        print("vInit(index=int) - Initialize with an existing document by index.")
        print("vInit(filename='path/to/file.vsdx') - Initialize with an existing file.")
        print("vInit(new=True, template='path/to/template.vstx') - Create a new document, optionally with a template.")
        return

    app = doc.Application
    page = list(doc.Pages)[0]
    window = app.ActiveWindow

    if not suffix:
        suffix = ''

    if globals_dict is not None:
        globals_dict[f'vApp{suffix}'] = app
        globals_dict[f'vDoc{suffix}'] = doc
        globals_dict[f'vPg{suffix}'] = page
        globals_dict[f'vWin{suffix}'] = window
        globals_dict['c'] = win32com.client.constants
        msg = f'''Instantiated the variables vApp{suffix}, vDoc{suffix}, vPg{suffix} and vWin{suffix} for the document {doc.Name}, 
as well as the variable c for the Visio constants'''
        print(msg)

    return app, doc, page, window


def get_or_open_visio_file(filename):
    """
    Check if a Visio file is already open. If not, open it.
    """
    docs = loaded_docs()
    for doc in docs:
        if doc.FullName.lower() == filename.lower():
            return doc
    return open_visio_file(filename)


def open_visio_file(file_path=None):
    """
    Open a Visio file.
    """
    global constants
    if file_path is None:
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(
            title="Select a Visio file",
            filetypes=[("Visio files", "*.vsd;*.vsdx")]
        )
        print(f"Selected file path: {file_path}")

    if file_path:
        file_path = os.path.normpath(file_path)
        print(f"Normalized file path: {file_path}")

        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"File does not exist: {file_path}")

        visio = win32com.client.Dispatch("Visio.Application")
        constants = win32com.client.constants
        try:
            doc = visio.Documents.Open(file_path)
            return doc
        except Exception as e:
            raise Exception(f"Error opening file: {e}")
    else:
        raise ValueError("No file selected.")


def create_new_document(template=None):
    """
    Create a new Visio document.
    """
    global constants
    visio = win32com.client.Dispatch("Visio.Application")
    constants = win32com.client.constants
    if template:
        doc = visio.Documents.Add(template)
    else:
        doc = visio.Documents.Add("")
    return doc
