# Microinvest Excel to XML converter

This python utility converts Chronological reports, generated from the Microinvest Delta Pro program and converts them to XML format, ready for import.

Create a virtual env for python to run:
```
pipenv shell
```

Install requirements:
```
pipenv install -r requirements.txt
```

Run the program:
```
python convert_xls.py -f path/to/xls -o path/to/output -r [start_row]
```
