import sys
import os

# Determine the directory of the current script
current_dir = os.path.dirname(os.path.abspath(__file__))

# Construct the absolute path to the visioipy directory
visioipy_path = os.path.join(current_dir, '../visiopy')
visioipy_path = os.path.abspath(visioipy_path)  # Normalize the path
print(f"Inserting path to sys.path: {visioipy_path}")

# Ensure the visioipy module is imported from the correct path
sys.path.insert(0, visioipy_path)

# Debug: Print sys.path to verify the correct path is included
print("sys.path:", sys.path)

# Additional Debug: List contents of the visioipy directory
print("Contents of visiopy directory:", os.listdir(visioipy_path))

try:
    import visio_connect
except ModuleNotFoundError as e:
    print(f"Failed to import visio_connect: {e}")
    sys.exit(1)

def test_loaded_docs():
    try:
        docs = visio_connect.loaded_docs()
        assert isinstance(docs, list), "loaded_docs should return a list"
        print("Test passed!")
    except Exception as e:
        print(f"Test failed: {e}")

def test_vInit():
    try:
        visio_connect.vInit(new=True, globals_dict=globals())
        print(vDoc.Name)
    except Exception as e:
        print(f"Failed creating doc: {e}")

if __name__ == "__main__":
    test_loaded_docs()
    test_vInit()
