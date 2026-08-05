import os
import sys

# Entry point alias for Streamlit Cloud (in case Streamlit Cloud deployment points to streamlit_app.py)
app_file = os.path.join(os.path.dirname(__file__), "app.py")
with open(app_file, encoding="utf-8") as f:
    code = f.read()

exec(code, globals())
