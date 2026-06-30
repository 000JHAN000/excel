# Respuestas - Análisis de Resultados

A continuación se responden las preguntas planteadas en la guía de aprendizaje, con base en los datos de los 10 empleados registrados en el archivo `nomina_empresa.xlsx`.

## 1. ¿Cuál fue el valor total pagado por la empresa en nómina?

**$29.203.291,67**

Este valor corresponde a la suma de la columna **Neto a Pagar** de todos los empleados, es decir, el total de dinero que la empresa debe consignar a su personal después de aplicar los descuentos de salud y pensión.

## 2. ¿Cuánto dinero se descontó por salud?

**$1.269.708,33**

El descuento por salud corresponde al **4 % del devengado** de cada empleado. Al sumar estos valores se obtiene el total destinado al aporte de salud.

## 3. ¿Cuánto dinero se descontó por pensión?

**$1.269.708,33**

Al igual que la salud, la pensión equivale al **4 % del devengado** de cada empleado. El valor total coincide con el de salud porque ambos porcentajes son iguales.

## 4. ¿Cuál empleado recibió el mayor pago?

**Valentina Morales** recibió el mayor pago, con un neto a pagar de **$3.496.000,00**.

Esto se debe a que es la **Project Manager** y tiene el salario base más alto de la empresa ($3.800.000), además de no tener horas extras que incrementen su base de descuento.

## 5. ¿Qué ventajas tiene automatizar la nómina mediante Python?

- **Reduce errores humanos:** elimina fallos de cálculo propios del proceso manual en Excel.
- **Ahorra tiempo:** genera toda la nómina y los reportes en segundos.
- **Es reproducible:** se puede ejecutar cada mes con los mismos datos o actualizarlos fácilmente.
- **Genera archivos profesionales:** crea libros de Excel con formatos, fórmulas y estilos automáticamente.
- **Facilita el análisis:** permite construir resúmenes gerenciales e indicadores para la toma de decisiones.
- **Es escalable:** se adapta fácilmente si la empresa crece y aumenta el número de empleados.

---

**Archivos generados:**
- `nomina_empresa.xlsx` — libro de Excel con la nómina y el resumen gerencial.
- `dist/nomina_sena.exe` — ejecutable que genera el archivo Excel al ejecutarse.
- `nomina_sena.py` — código fuente en Python.
