## Init Python Environment
```sh
python -m venv .venv
.\\.venv\\Scripts\\activate
pip install -r requirements.txt
```

## How to run
1. Use Task Manager to close all Edge web browser and run `& 'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe' --remote-debugging-port=9222` in PowerShell
2. Disable xlsx file encrypt, and close Excel.
3. Run: `python teams_test.py <your_xlsx_file_path> <sheet_name> --save`.
4. Wait Teams website load success and go to the chat page which you want to test.
5. Back to terminal and press "Enter" to continue.
6. Wait it finish and check the xlsx file content.