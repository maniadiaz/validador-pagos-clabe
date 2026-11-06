# Validador de pagos CLABE

Con Python 3, instala la dependencia y ejecuta el validador con el Excel del reto.

```sh
python -m pip install -r requirements.txt
python validar_pagos.py "Datos de entrada simple.xlsx"
```

El script lee la primera hoja y genera `resultado.xlsx`. Usa `--salida` para indicar otra ruta. El resultado incluido contiene 9 filas válidas y 6 inválidas.

Las columnas requeridas son `Beneficiario`, `CLABE` e `Importe`. El importe debe estar guardado como número en Excel. Un valor como `"5000"` guardado como texto se marca como inválido. Las filas con los tres campos vacíos se dejan sin estado.

## ¿Qué fue lo más difícil del reto y cómo lo resolviste?

Distinguir los espacios de una pérdida de dígitos en una CLABE. Quito los espacios al validar y conservo el valor de entrada en el Excel. Si quedan menos de 18 dígitos, reporto cuántos llegaron. Esto evita dar por válida una cuenta completada con ceros que el archivo no permite confirmar.
