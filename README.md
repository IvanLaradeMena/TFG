#WCA Translator – README
Description
Python tool to extract component parameters from netlists (.net, .sxsch) and BoMs (.bom, .csv), normalize values, compute Worst-Case deviations, and generate an Excel file ready for import into Mathcad Prime via COM.

##Requirements
Windows 10/11 with Mathcad Prime 4–10# Traductor WCA – README

## Descripción
Herramienta en Python para extraer parámetros de componentes desde netlists (.net, .sxsch) y BoM (.bom, .csv), normalizar valores, calcular desviaciones Worst-Case y generar un Excel listo para importar en Mathcad Prime vía COM.

## Requisitos
Windows 10/11 con Mathcad Prime 4–10

Python 3.10+

Paquetes: openpyxl, comtypes, tkinter

## Instalación
Clonar:
git clone https://github.com/IvanLaradeMena/TFG

Entrar al proyecto y, opcionalmente, crear entorno virtual.

Instalar dependencias:
pip install -r requirements.txt

## Uso GUI
python main.py

Seleccionar archivo (.net/.bom/.csv)

Introducir H(s) (opcional)

Revisar o editar Excel

Confirmar o elegir plantilla Mathcad

Datos volcados en Mathcad Prime

## Uso CLI
python main.py --input ejemplo.net --hs "1/(R1*C1*s+1)" --no-gui

## Estructura
traductor-wca/

├── main.py

├── traductor.py 

├── auto_mathcad.py

├── requirements.txt

└── README.md

## Autoría del proyecto
Autor: Iván Lara de Mena

Tutores: Ángel Á. Sánchez, Ángel Otero R.


Python 3.10+

Packages: openpyxl, comtypes, tkinter

##Installation
Clone:
git clone https://github.com/IvanLaradeMena/TFG

Enter the project directory and optionally create a virtual environment.

Install dependencies:
pip install -r requirements.txt

GUI Usage
python main.py

Select file (.net/.bom/.csv)

Enter H(s) (optional)

Review or edit Excel

Confirm or choose Mathcad template

Data transferred to Mathcad Prime

##CLI Usage
python main.py --input example.net --hs "1/(R1C1s+1)" --no-gui

##Structure
traductor-wca/

├── main.py
├── traductor.py
├── auto_mathcad.py
├── requirements.txt
└── README.md

##Project Authorship
Author: Iván Lara de Mena
Supervisors: Ángel Á. Sánchez, Ángel Otero R.

Supervisors: Ángel Á. Sánchez, Ángel Otero R.
