# Streamlit Cloud arranca este archivo: solo ejecuta la versión nueva del ERP (app.py)
import runpy
from pathlib import Path
runpy.run_path(str(Path(__file__).with_name("app.py")), run_name="__main__")