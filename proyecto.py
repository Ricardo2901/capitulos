"""
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Instalacion de Python
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||

|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Hacer una REST API con Django
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    ============================================================
        Instalaciones con PIP
    ============================================================
    Para instalar Django, necesitamos tener Python instalado en nuestro sistema operativo. Una vez ya instalado hay que hacer lo siguiente:
    1. Abrir la terminal o consola de comandos.
    2. Instalar pip, que es necesario para instalar paquetes de Python.
    3. Instalar Django con pip.
        pip install django
    4. Instalar Django REST API:
        pip install djangorestframework
    5. 
    
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Hacer documentos en Word con Python
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    ============================================================
        Importaciones para que funcione bien el codigo
    ============================================================
        S= es el número de especies presentes

    ============================================================
        Variables
    ============================================================


    ============================================================
        Contenido
    ============================================================


    ============================================================
        Tablas y Celdas
    ============================================================
    #########################
    ### Título de la tabla del capítulo 13.1 ###
    #########################
    tituloTabla13b = doc.add_paragraph()
    dti13b = tituloTabla13b.add_run('\nTabla 9.1.- Periodo de ejecución por etapa.')
    dti13b_format = tituloTabla13b.paragraph_format
    dti13b_format.line_spacing = 1.15
    dti13b_format.space_after = 0

    dti13b.font.name = 'Bookman Old Style'
    dti13b.font.size = Pt(12)
    tituloTabla13b.alignment = WD_ALIGN_PARAGRAPH.CENTER

    #########################
    ### Tabla del capítulo 13.1 ###
    #########################
    filas = 1
    columnas = 2
    tabla13b = doc.add_table(rows=filas, cols=columnas, style='Table Grid')

    for cols in range(filas):
        cell = tabla13b.cell(0, cols)
        cell_background_color(cell, '0070C0')

        for rows in range(columnas):
            cell = tabla13b.cell(rows, cols)
            t13b = cell.paragraphs[0].add_run(' ')
            t13b.font.size = Pt(12)

    ============================================================
        Guardar Documento
    ============================================================


"""