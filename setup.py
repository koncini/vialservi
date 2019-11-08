from cx_Freeze import setup, Executable

setup(name="Ventana",
      version="0.1",
      description="Ventana",
      executables=[Executable("analizar.py", base="Win32GUI")],)
