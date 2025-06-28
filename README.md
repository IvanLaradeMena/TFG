# WCA Translator – README

## Description
Python tool for extracting component parameters from netlists (.net, .sxsch) and BoM (.bom, .csv), normalizing values, calculating worst-case deviations, and generating an Excel file ready for import into Mathcad Prime via COM.
This project is a Final Degree Project (TFG) at Universidad Rey Juan Carlos, carried out in collaboration with Thales Alenia Space (Spain).

## Requirements
Windows 10/11 with Mathcad Prime 4–10

Python 3.10+

Packages: openpyxl, comtypes, tkinter

## Installation
Clone:
git clone https://github.com/IvanLaradeMena/TFG

Go to the project and, optionally, create a virtual environment.

Install dependencies:
pip install -r requirements.txt

## GUI Usage
python main.py

Select file (.net/.bom/.csv)

Enter H(s) (optional)

Review or edit Excel

Confirm or choose Mathcad template

Data dumped to Mathcad Prime

## CLI Usage
python main.py --input example.net --hs "1/(R1*C1*s+1)" --no-gui

## Structure
wca-translator/

├── main.py

├── translator.py

├── auto_mathcad.py

├── requirements.txt

└── README.md

## Project Authorship
Author: Iván Lara de Mena

Tutors: Ángel Á. Sánchez, Ángel Otero R.
