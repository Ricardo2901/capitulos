from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def cell_background_color(cell, color_hex):
    # Obtener las propiedades de la celda
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()

    # Crear un nuevo elemento de color de fondo
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def entero_a_romano(numero):
    valores = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
    ]

    resultado = ""

    for (arabigo, romano) in valores:
        while numero >= arabigo:
            resultado += romano
            numero -= arabigo
    return resultado