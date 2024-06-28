import win32com.client
import pythoncom
import tkinter as tk
from tkinter import filedialog
import os

# Global variables to store Visio objects
_visio_apps = {}
_visio_docs = {}
_visio_pages = {}
_visio_windows = {}

constants = win32com.client.constants

help_text = '''Visio utils is a minimalistic Visio library.
'''
print(help_text)
def loaded_docs(index=None):
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


def vInit(index=None, template=None, filename=None):
    """
    Initialize the Visio application and set global variables for vApp, vDoc, vPg, and vWin.

    Parameters:
    - index: Integer, index of the loaded docs.
    - template: String, path to the template file for creating a new document.
    - filename: String, path to an existing Visio file to open.
    """
    global _visio_apps, _visio_docs, _visio_pages, _visio_windows

    if index is not None:
        doc = loaded_docs(index)
    elif filename is not None:
        doc = open_visio_file(filename)
    elif template is not None:
        doc = create_new_document(template)
    else:
        doc = ask_for_document()

    app = doc.Application
    page = list(doc.Pages)[0]
    window = app.ActiveWindow

    suffix = str(index) if index else ""

    _visio_apps[suffix] = app
    _visio_docs[suffix] = doc
    _visio_pages[suffix] = page
    _visio_windows[suffix] = window

    globals()[f'vApp{suffix}'] = app
    globals()[f'vDoc{suffix}'] = doc
    globals()[f'vPg{suffix}'] = page
    globals()[f'vWin{suffix}'] = window

    print(f'Instantiated vApp{suffix}, vDoc{suffix}, vPg{suffix}, vWin{suffix}')
    return app, doc, page, window


def open_visio_file(file_path=None):
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
        try:
            doc = visio.Documents.Open(file_path)
            return doc
        except Exception as e:
            raise Exception(f"Error opening file: {e}")
    else:
        raise ValueError("No file selected.")


def create_new_document(template=None):
    visio = win32com.client.Dispatch("Visio.Application")
    if template:
        doc = visio.Documents.Add(template)
    else:
        doc = visio.Documents.Add("")
    return doc


def ask_for_document():
    choice = input("Open existing file (e) or create new (n)? ").lower()
    if choice == 'e':
        return open_visio_file()
    elif choice == 'n':
        template = input("Enter template path (or press Enter for blank): ")
        return create_new_document(template if template else None)
    else:
        raise ValueError("Invalid choice.")
