
# Deployment Guide for PythonAnywhere

## Files included:
- app.py
- requirements.txt
- data.db (empty starter)
- README_deployment.txt

## Steps:
1. Upload all files to PythonAnywhere under `/home/yourusername/`
2. Open Bash console and run:
   pip install --user -r requirements.txt

3. Add a new web app → Manual configuration → Python 3.10

4. Edit WSGI file to:
   import os
   os.system("streamlit run /home/yourusername/app.py --server.port=$PORT --server.address=0.0.0.0")

5. Reload the web app.

Your app will now be live at:
https://yourusername.pythonanywhere.com
