import streamlit.web.cli as stcli
import os
import sys

if __name__ == '__main__':
    # Trỏ đường dẫn đến file App.py
    script_path = os.path.join(os.path.dirname(__file__), 'App.py')
    sys.argv = ["streamlit", "run", script_path, "--global.developmentMode=false"]
    sys.exit(stcli.main())