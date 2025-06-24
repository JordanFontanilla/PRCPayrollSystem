import sys
import os

def resource_path(relative_path):
    # Get absolute path to resource, works for dev and for PyInstaller bundle
    #recommended abosulute path, REMEMBER this is for resourcesss
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def data_path(relative_path):
    # Use a folder next to the executable or in the user's home directory
    if hasattr(sys, '_MEIPASS'):
        # When bundled, use the directory of the executable
        base_dir = os.path.dirname(sys.executable)
    else:
        # When running as script, use current directory
        base_dir = os.path.abspath(".")
    return os.path.join(base_dir, relative_path) 