.\.venv\Scripts\Activate
pip install -r requirements.txt

python .\convert_excel.py $args

deactivate