ChatGPT said:
Traductor WCA – README
Descripción
Herramienta en Python para extraer parámetros de componentes desde netlists (.net, .sxsch) y BoM (.bom, .csv), normalizar valores, calcular desviaciones Worst-Case y generar un Excel listo para importar en Mathcad Prime vía COM.

Requisitos
Windows 10/11 con Mathcad Prime 4–10

Python 3.10+

Paquetes: openpyxl, comtypes, tkinter

Instalación
Clonar:
git clone https://github.com/tu-usuario/traductor-wca.git

Entrar al proyecto y, opcionalmente, crear entorno virtual.

Instalar dependencias:
pip install -r requirements.txt

Uso GUI
python main.py

Seleccionar archivo (.net/.bom/.csv)

Introducir H(s) (opcional)

Revisar o editar Excel

Confirmar o elegir plantilla Mathcad

Datos volcados en Mathcad Prime

Uso CLI
python main.py --input ejemplo.net --hs "1/(R1*C1*s+1)" --no-gui

Estructura
main.py – interfaz Tkinter y CLI

traductor.py – parsing y Excel

auto_mathcad.py – automatización COM

requirements.txt – dependencias

Autor: Iván Lara de Mena
Tutores: Ángel Á. Sánchez, Ángel Otero R.
