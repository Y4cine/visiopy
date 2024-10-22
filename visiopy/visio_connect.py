import win32com.client
import pythoncom
import os
import tkinter as tk
from tkinter import filedialog, messagebox

c = []  # to hold Visio constants


def get_visio_clsids():
    """
    Retrieve the CLSIDs for the installed Visio application.
    """
    try:
        visio_app = win32com.client.Dispatch("Visio.Application")
        clsid_main = f"{{{visio_app._oleobj_.GetTypeInfo().GetContainingTypeLib()[0].GetLibAttr()[0]}}}"
        # Common CLSID for unsaved documents
        clsid_unsaved = "{00021A20-0000-0000-C000-000000000046}"
        clsids = [clsid_main, clsid_unsaved]
        # print(f"Retrieved CLSIDs: {clsids}")
        return clsids
    except Exception as e:
        raise Exception(f"Error retrieving Visio CLSIDs: {e}")

def vDocs(index=None, silent=False):
    """
    Prints the list of all open Visio drawings in all Visio instances and returns the list of the document objects.

    Parameter:
    - index: integer, optional makes the function return the document object with this index.
    """
    global c

    visio_clsids = get_visio_clsids()
    context = pythoncom.CreateBindCtx(0)
    rot = pythoncom.GetRunningObjectTable()
    docs = []

    for moniker in rot:
        try:
            name = moniker.GetDisplayName(context, None)
            #print(f"Processing moniker: {name}")  # Debugging statement
            if any(clsid in name for clsid in visio_clsids) or name.endswith('.vsdx') or name.endswith('.vsdm') or name.endswith('.vstx') or name.endswith('.vstm'):
                try:
                    visio_doc = moniker.BindToObject(
                        context, None, pythoncom.IID_IDispatch)
                    visio_doc = win32com.client.Dispatch(
                        visio_doc.QueryInterface(pythoncom.IID_IDispatch))
                    docs.append(visio_doc)
                    c = win32com.client.constants
                except Exception as e:
                    if "{00021A20-0000-0000-C000-000000000046}" in name:
                        if not silent:
                            print(f"Unsaved document encountered: {name}. Please save the document.")
                    else:
                        print(f"Error processing document '{name}': {e}")
        except Exception as e:
            print(f"Error retrieving display name for document: {e}")

    if index is not None:
        if 0 <= index < len(docs):
            return docs[index]
        else:
            raise ValueError("Invalid document index.")
    if not silent:    
        print('-'*20)
        if not docs:
            print("No open Visio documents found.")
        else:
            for i, doc in enumerate(docs):
                print(f"{i}: {doc.FullName}")
    return docs

def vInit(index=None, filename=None, new=False, template=None, g=None, suffix=None):
    """
    Initializes the Visio application and sets global variables for vApp, vDoc, vPg, and vWin.

    The function can be used to open an existing Visio document, create a new one (optionally from a template), 
    or allow the user to select a file through a dialog if no parameters are provided.

    Parameters:
    ----------
    - index : int, optional
        Index of the loaded document (as listed by vDocs()). If provided, the document at this index is loaded.
    - filename : str, optional
        Path to an existing Visio file to open. If provided, the specified file is opened.
    - new : bool, optional
        If True, creates a new document. If a template path is also provided, the document is created from the template.
    - template : str, optional
        Path to the template file for creating a new document (used only if 'new=True').
    - suffix : str, optional
        A string suffix to append to the global variables (e.g., 'vApp1', 'vDoc1', ...).
    - g : dict, optional
        Pass `globals()` to automatically instantiate the global variables (vApp, vDoc, vPg, vWin, and Visio constants).

    Returns:
    --------
    - app : win32com.client.CDispatch
        The Visio application instance.
    - doc : win32com.client.CDispatch
        The Visio document instance.
    - page : win32com.client.CDispatch
        The first page of the Visio document.
    - window : win32com.client.CDispatch
        The active window in the Visio application.
    - c : win32com.client.constants
        The Visio constants (for easy access to Visio-specific enums like c.visSelect).
    
    Usage:
    ------
    1. To list open documents and load one by index:
        from visiopy import vInit, vDocs
        vDocs()  # lists open Visio documents
        vInit(index=1, g=globals())  # loads the second document in the list

    2. To open an existing Visio file by filename:
        vInit(filename="C:/path/to/file.vsdx", g=globals())

    3. To create a new document:
        vInit(new=True, g=globals())

    4. To create a new document from a template:
        vInit(new=True, template="C:/path/to/template.vstm", g=globals())

    5. Without arguments, a file dialog will prompt the user to select a document:
        vInit(g=globals())

    Notes:
    ------
    - If no parameters are provided, the function opens a file selection dialog to choose a Visio file.
    - The function modifies global variables if `g=globals()` is passed, enabling convenient access to 
      the Visio constants and objects (vApp, vDoc, vPg, vWin) directly in the global scope.

    """

    doc = None

    # If no parameters are passed, open the tkinter document manager
    if not any([index, filename, new, template]):
        result = document_manager()
        if 'doc' in result:
            doc = result['doc']
        elif 'filename' in result:
            filename = result['filename']
            doc = get_or_open_visio_file(filename)
        elif result.get('new', False):
            doc = create_new_document(result.get('template'))

    if index is not None:
        doc = vDocs(index, silent = True)
    elif filename is not None:
        doc = get_or_open_visio_file(filename)
    elif new:
        doc = create_new_document(template)

    if not doc:
        return

    app = doc.Application
    page = list(doc.Pages)[0]
    window = app.ActiveWindow

    if not suffix:
        suffix = ''

    if g is not None:
        g[f'vApp{suffix}'] = app
        g[f'vDoc{suffix}'] = doc
        g[f'vPg{suffix}'] = page
        g[f'vWin{suffix}'] = window
        g['c'] = win32com.client.constants
        msg = f'''Instantiated the variables vApp{suffix}, vDoc{suffix}, vPg{suffix} and vWin{suffix} for the document {doc.Name}, 
as well as the variable c for the Visio constants'''
        print(msg)

    return app, doc, page, window, c

def ask_for_visio_file(title="Select a Visio file", filetypes=[("Visio files", "*.vsd;*.vsdx;*.vsdm;*.vstx;*.vstm")]):
    """
    Opens a file dialog and returns the selected Visio file.
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    root.destroy()  # Destroy the tkinter window

    if file_path:
        # Normalize the path to the platform-specific format
        file_path = os.path.normpath(file_path)

        # Check if the file exists
        if not os.path.isfile(file_path):
            raise FileNotFoundError(f"File does not exist: {file_path}")
    return file_path

def get_or_open_visio_file(filename):
    """
    Check if a Visio file is already open. If not, open it.
    """
    docs = vDocs(silent = True)
    for doc in docs:
        if doc.FullName.lower() == filename.lower():
            return doc
    return open_visio_file(filename)


def open_visio_file(file_path=None):
    """
    Open a Visio file.
    """
    if file_path is None:
        file_path = ask_for_visio_file()
    if file_path:
        visio = win32com.client.Dispatch("Visio.Application")
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
    visio = win32com.client.Dispatch("Visio.Application")
    if template:
        print('create_new_document',template)
        doc = visio.Documents.Add(template)
    else:
        doc = visio.Documents.Add("")
    return doc

def document_manager():
    '''Function to open a tkinter form and return a result'''
    result = {}  # A dictionary to hold return values
    docs = []
    root = tk.Tk()
    root.title("Document Manager")
    root.geometry("400x400")
    root.resizable(True, True)

    def open_selected_doc():
        selected_doc_name = doc_listbox.get(tk.ACTIVE)
        selected_doc = None
        for doc in docs:
            if selected_doc_name == doc.Name:
                selected_doc = doc
                break
        if selected_doc:
            result['doc'] = selected_doc
            root.quit()  # Close the tkinter window after selecting the document
        else:
            messagebox.showwarning("No Selection", "Please select a document.")

    def pick_file_from_folder():
        file_path = ask_for_visio_file()
        if file_path:
            result['filename'] = file_path
            root.quit()  # Close the window after selecting a file

    def new_blank_document():
        result['new'] = True
        result['template'] = None  # For new blank document, no template
        root.quit()

    def new_document_from_template():
        template_file = ask_for_visio_file()
        if template_file:
            result['new'] = True
            result['template'] = template_file
            root.quit()

    def close_form():
        if not result:
            if messagebox.askyesno("Close", "No document selected. Are you sure you want to close?"):
                root.quit()

    # Create a frame for the listbox and scrollbar
    frame = tk.Frame(root)
    frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

    # Scrollable Listbox
    scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL)
    doc_listbox = tk.Listbox(frame, selectmode=tk.SINGLE, yscrollcommand=scrollbar.set)
    scrollbar.config(command=doc_listbox.yview)

    # Add scrollbar to the side
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    doc_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Populate the listbox with documents
    docs = vDocs(silent = True)
    for doc in docs:
        doc_listbox.insert(tk.END, doc.Name)

    # Create a frame for buttons to organize them in a grid
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    # Create buttons in a grid with padding
    tk.Button(button_frame, text="Open Selected Doc", command=open_selected_doc).grid(row=0, column=0, padx=5, pady=5)
    tk.Button(button_frame, text="Pick File from Folder", command=pick_file_from_folder).grid(row=0, column=1, padx=5, pady=5)
    tk.Button(button_frame, text="New Blank Document", command=new_blank_document).grid(row=1, column=0, padx=5, pady=5)
    tk.Button(button_frame, text="New Document from Template", command=new_document_from_template).grid(row=1, column=1, padx=5, pady=5)
    
    # Close button at the bottom
    tk.Button(root, text="Close", command=close_form).pack(pady=10)

    root.mainloop()
    root.destroy()
    return result