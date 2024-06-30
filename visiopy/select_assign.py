import win32com.client
import time
import threading

# Create Visio application object and get necessary handles
vApp = win32com.client.Dispatch("Visio.Application")
vDoc = vApp.ActiveDocument
vPg = vApp.ActivePage
vWin = vApp.ActiveWindow

# Define initial field and value
selected_field = "Field1"
selected_value = "Value1"
check_active = True


def get_property_set_fields():
    # Placeholder function to return a list of property set fields
    return ["Field1", "Field2", "Field3"]


def set_value(shape, field, value):
    try:
        if shape.CellExists(f"prop.{field}", False):
            shape.Cells(f"prop.{field}.Value").FormulaU = f'"{value}"'
    except Exception as e:
        print(f"Error setting value: {e}")


# Initialize field and value options (similar to UserForm_Initialize)
field_options = get_property_set_fields()
value_options = ["?", "A", "B", "D", "M", "P", "V", "VP", "VR", "VS", "VM", "S",
                 "SG", "SL", "SP", "ST", "SW", "FALSE", "TRUE", "0", "1", "2", "3", "4", "5"]

# Function to batch modify selected shapes


def batch_modify_shapes(field, value):
    selection = vWin.Selection
    selection.IterationMode = 64  # visSelModeSkipSuper
    if selection.Count > 0 and check_active:
        for i in range(1, selection.Count + 1):
            shape = selection.Item(i)
            set_value(shape, field, value)

# Monitor selection changes in a separate thread


def monitor_selection_changes():
    previous_selection = None
    while True:
        current_selection = vWin.Selection
        if current_selection is not previous_selection:
            batch_modify_shapes(selected_field, selected_value)
            previous_selection = current_selection
        time.sleep(1)  # Adjust the sleep time as necessary


# Start the monitoring in a separate thread
monitor_thread = threading.Thread(
    target=monitor_selection_changes, daemon=True)
monitor_thread.start()

# The main thread can continue with other tasks
print("Monitoring selection changes in Visio...")
