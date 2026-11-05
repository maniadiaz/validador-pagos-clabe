import argparse
import re
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def error_clabe(valor):
    if valor is None:
        return "CLABE vacía"
    if type(valor) is int:
        valor = str(valor)
    elif isinstance(valor, str):
        valor = "".join(valor.split())
    else:
        return "CLABE debe contener solo dígitos enteros"

    if not valor:
        return "CLABE vacía"
    if not re.fullmatch(r"[0-9]+", valor):
        return "CLABE contiene caracteres distintos de los dígitos 0 a 9"
    if len(valor) != 18:
        return f"CLABE tiene {len(valor)} dígitos (debe tener 18)"
    return None


def error_importe(valor):
    if type(valor) not in (int, float):
        return "Importe debe ser numérico"
    importe = Decimal(str(valor))
    if not importe.is_finite():
        return "Importe debe ser un número finito"
    if importe <= 0:
        return "Importe debe ser mayor que cero"
    decimales = max(0, -importe.normalize().as_tuple().exponent)
    if decimales > 2:
        return f"Importe tiene {decimales} decimales (máximo 2)"
    return None


def validar(entrada, salida):
    if entrada.resolve() == salida.resolve():
        raise ValueError("La entrada y la salida deben ser archivos distintos")
    libro = load_workbook(entrada)
    hoja = libro.worksheets[0]
    encabezados = [celda.value for celda in hoja[1]]
    requeridos = ("Beneficiario", "CLABE", "Importe", "Concepto")
    for nombre in requeridos:
        if encabezados.count(nombre) != 1:
            raise ValueError(f"Se requiere exactamente una columna {nombre}")
    beneficiario, clabe, importe, _ = (
        encabezados.index(nombre) + 1 for nombre in requeridos
    )
    columna_estado = len(encabezados) + 1
    columna_motivo = columna_estado + 1
    ultima_fila = hoja.max_row
    hoja.cell(1, columna_estado, "estado")
    hoja.cell(1, columna_motivo, "motivo_error")
    hoja.column_dimensions[get_column_letter(columna_estado)].width = 14
    hoja.column_dimensions[get_column_letter(columna_motivo)].width = 65

    for fila in range(2, ultima_fila + 1):
        nombre = hoja.cell(fila, beneficiario).value
        errores = []
        if nombre is None or not str(nombre).strip():
            errores.append("Beneficiario vacío")
        for error in (
            error_clabe(hoja.cell(fila, clabe).value),
            error_importe(hoja.cell(fila, importe).value),
        ):
            if error:
                errores.append(error)
        hoja.cell(fila, columna_estado, "inválido" if errores else "válido")
        if errores:
            hoja.cell(fila, columna_motivo, " | ".join(errores))
    libro.save(salida)


def main():
    parser = argparse.ArgumentParser(description="Valida un lote de pagos en Excel.")
    parser.add_argument("entrada", type=Path, help="Archivo XLSX con los pagos")
    parser.add_argument("-o", "--salida", type=Path, default=Path("resultado.xlsx"))
    args = parser.parse_args()
    try:
        validar(args.entrada, args.salida)
    except (OSError, ValueError) as error:
        parser.exit(1, f"Error: {error}\n")
    print(f"Resultado generado en {args.salida}")


if __name__ == "__main__":
    main()
