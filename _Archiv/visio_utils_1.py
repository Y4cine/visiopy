import win32com.client
import pythoncom
import tkinter as tk
from tkinter import simpledialog, filedialog
import os

constants = []

# Function to list loaded documents


def loaded_docs(index=None):
    global constants
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
                constants = win32com.client.constants
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


def vInit(index=None, filename=None, new=False, template=None, globals_dict=None):
    """
    Initialize the Visio application and set global variables for vApp, vDoc, vPg, and vWin.

    Parameters:
    - index: Integer, index of the loaded docs.
    - filename: String, path to an existing Visio file to open.
    - new: Boolean, if True, create a new document (optionally with a template).
    - template: String, path to the template file for creating a new document.
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

    if globals_dict is not None:
        globals_dict['vApp'] = app
        globals_dict['vDoc'] = doc
        globals_dict['vPg'] = page
        globals_dict['vWin'] = window
        globals_dict['constants'] = win32com.client.constants
    else:
        globals()['vApp'] = app
        globals()['vDoc'] = doc
        globals()['vPg'] = page
        globals()['vWin'] = window

    return app, doc, page, window

    return app, doc, page, window

# Function for interactive initialization


def vInitGUI():
    """
    Interactive initialization of the Visio application using GUI dialogs.
    """
    root = tk.Tk()
    root.withdraw()

    choice = simpledialog.askstring(
        "Initialization", "Choose action: (loaded, file, new, template)")

    if choice == "loaded":
        loaded_docs()
        index = simpledialog.askinteger(
            "Select Document", "Enter document index:")
        if index is not None:
            return vInit(index=index)
    elif choice == "file":
        filename = filedialog.askopenfilename(
            title="Select a Visio file",
            filetypes=[("Visio files", "*.vsd;*.vsdx")]
        )
        if filename:
            return vInit(filename=filename)
    elif choice == "new":
        return vInit(new=True)
    elif choice == "template":
        template = filedialog.askopenfilename(
            title="Select a template file",
            filetypes=[("Visio template files", "*.vst;*.vstx")]
        )
        if template:
            return vInit(new=True, template=template)
    else:
        print("Invalid choice. Please choose 'loaded', 'file', 'new', or 'template'.")
        return

# Helper functions to manage Visio documents


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
