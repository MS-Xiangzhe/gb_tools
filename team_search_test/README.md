## Init Python Environment
```sh
python -m venv .venv
.\\.venv\\Scripts\\activate
pip install -r requirements.txt
```

## How to run
1. Disable xlsx file encrypt, and close Excel.
2. Run: `python teams_test.py <your_xlsx_file_path> <sheet_name> --save`.
3. Wait Teams website load success and go to the chat page which you want to test.
4. Back to terminal and press "Enter" to continue.
5. Wait it finish and check the xlsx file content.