import os
import sys
import ctypes


def check_dlls():
    print(f"Python: {sys.version}")
    try:
        import PySide6

        print(f"PySide6 location: {PySide6.__file__}")
        pyside_dir = os.path.dirname(PySide6.__file__)
        print(f"PySide6 directory: {pyside_dir}")

        # Try to find conflicting DLLs in PATH
        path_dirs = os.environ.get("PATH", "").split(os.pathsep)
        qt_dlls = ["Qt6Core.dll", "Qt6Widgets.dll", "Qt6Gui.dll"]

        print("\nChecking for Qt6 DLLs in PATH...")
        for d in path_dirs:
            for dll in qt_dlls:
                full_path = os.path.join(d, dll)
                if os.path.exists(full_path):
                    print(f"FOUND CONFLICT: {full_path}")

        print("\nAttempting to import QtWidgets...")
        from PySide6 import QtWidgets

        print("SUCCESS: QtWidgets imported successfully!")
    except ImportError as e:
        print(f"\nFAILED: {e}")
        # Try to use add_dll_directory as a fix
        if sys.platform == "win32" and "pyside_dir" in locals():
            print(f"\nAttempting fix: os.add_dll_directory('{pyside_dir}')")
            os.add_dll_directory(pyside_dir)
            try:
                from PySide6 import QtWidgets

                print("SUCCESS after fix!")
            except ImportError as e2:
                print(f"STILL FAILED: {e2}")


if __name__ == "__main__":
    check_dlls()
