# Práctica Power BI para la clase de apoyo de Lucía
Armada sobre el material real de la cursada (Clases 1, 2, 4, 5 y Segmentadores/Filtros) y los archivos que ella tiene:
`bd_integrantes_Esp.xlsx` (hojas integrantes + departamentos) y `Datos de servicio.xlsx` (hoja atenciones).

Nota: DAX se escribe con nombres de función en inglés (SUM, CALCULATE...). Según la configuración regional, el separador de argumentos es coma o punto y coma. Si da error de sintaxis, probar el otro.

---

## 0. Mapa del curso → lo que probablemente evalúan
| Clase | Tema | Qué saber hacer |
|---|---|---|
| 1 | Introducción | Obtener datos desde Excel, conocer paneles (Visualizaciones, Campos, Informe) |
| 2 | Modelado | Relación uno a muchos entre tablas, tipos de datos, columna calculada, gráfico con datos de 2 tablas |
| 4 | DAX y KPI | Columna calculada vs medida (SUM, AVERAGE, MAX, MIN), medida de objetivo, DIVIDE, visual KPI |
| 5 | Dashboard | Relación por ID, IF en columna, CALCULATE + COUNTROWS, KPI (valor, objetivo, tendencia), segmentadores |
| Seg./Filtros | Slicers y filtros | Segmentador (visible) vs filtro de visual/página/informe |

Diferencias a tener claras (pregunta típica de examen):
- **Columna calculada:** se calcula fila por fila y se guarda en la tabla. **Medida:** se calcula al vuelo según los filtros del visual.
- **Segmentador:** visible, lo usa el lector. **Filtro:** del panel de filtros, lo configura el autor y puede aplicarse a visual, página o informe.

---

## EJERCICIO 1 — Integrantes (modelado + DAX + KPI). Cubre Clases 2, 4 y 5

### Parte A. Importar y limpiar (Transformar datos)
1. Obtener datos → Excel → `bd_integrantes_Esp.xlsx` → marcar **integrantes** y **departamentos**.
2. `departamentos`: hay una fila vacía al final → *Quitar filas → Quitar filas en blanco*. En `Gerente`, "Vinicius R. " tiene un espacio final → *Transformar → Formato → Recortar*.
3. `integrantes`: la columna **Departamento** guarda códigos (dpt1...), pero en `departamentos` también hay una columna llamada Departamento (el nombre). Renombrar la de `integrantes` a **cod_departamento** para no confundirse.
4. `integrantes`: **Salario-Cargo** viene junta ("5060-Técnico de Producción I"). *Dividir columna → Por delimitador "-" → en el primer delimitador*. Renombrar a **Salario** (tipo número entero) y **Cargo** (texto).
5. **Status do Integrante** tiene "Ativo" (portugués). *Reemplazar valores* "Ativo" → "Activo". **Fecha salida** está vacía en todas las filas (todos activos): se puede eliminar la columna.
6. Verificar tipos: fechas como Fecha, Edad y Salario como número, resto texto. Cerrar y aplicar.

Ojo: hay nombres repetidos (Samuel, Lívia...) pero son personas distintas con ID distinto. **No** quitar duplicados por nombre.

### Parte B. Modelo
Vista Modelo: relacionar `departamentos[cod_departamento]` (lado uno) con `integrantes[cod_departamento]` (lado muchos). Verificar "uno a muchos" y que el filtro vaya de departamentos hacia integrantes.

### Parte C. Columnas calculadas (Modelado → Nueva columna)
```
Antigüedad = DATEDIFF('integrantes'[Fecha de contratación], TODAY(), YEAR)

Rango Etario = IF('integrantes'[Edad] < 30, "Menos de 30", IF('integrantes'[Edad] < 40, "30 a 39", "40 o más"))

Desempeño OK = IF('integrantes'[Registro de desempeño] = "Abajo de lo esperado", "No", "Sí")
```

### Parte D. Medidas (Nueva medida)
```
Total Integrantes = COUNTROWS('integrantes')
Edad Promedio = AVERAGE('integrantes'[Edad])
Edad Máxima = MAX('integrantes'[Edad])
Edad Mínima = MIN('integrantes'[Edad])
Satisfacción Promedio = AVERAGE('integrantes'[Índice de satisfacción])

Integrantes Desempeño Alto =
CALCULATE(COUNTROWS('integrantes'), 'integrantes'[Registro de desempeño] = "Arriba")

% Desempeño Alto = DIVIDE([Integrantes Desempeño Alto], [Total Integrantes])

Objetivo Satisfacción = 4
Cumplimiento Satisfacción = DIVIDE([Satisfacción Promedio], [Objetivo Satisfacción])
```

### Parte E. Dashboard
- Tarjetas: Total Integrantes, Edad Promedio, Satisfacción Promedio.
- KPI: valor = Satisfacción Promedio, objetivo = Objetivo Satisfacción, eje de tendencia = Departamento (nombre).
- Barras: Total Integrantes por Departamento (nombre, de la tabla departamentos).
- Columnas: Registro de desempeño por Rango Etario.
- Segmentadores: Departamento y Sexo. Un filtro de página que deje fuera algún valor, para mostrar la diferencia.

### Resultados esperados (para validar)
| Medida | Valor |
|---|---|
| Total Integrantes | 207 |
| Edad promedio / mín / máx | 35,96 / 23 / 54 |
| Satisfacción promedio | 3,87 |
| Desempeño "Arriba" | 29 (14,0%) |
| Cumplimiento satisfacción (obj. 4) | 96,7% |
| Por departamento | Producción 125 · TI 47 · Ventas 27 · Administrativo 8 |
| Satisfacción por depto. | Producción 3,82 · Ventas 3,89 · Administrativo 3,50 · TI 4,04 |
| Por sexo | F 117 · M 90 |

---

## EJERCICIO 2 — Datos de servicio (KPI, CALCULATE, segmentadores). Cubre Clases 4 y 5

### Limpieza
1. Importar la hoja **atenciones** (1.009 llamadas, marzo 2023).
2. **Área:** "RECLAMOS" está en mayúsculas y el resto no → *Formato → Capitalizar cada palabra*.
3. **Respondido - Resuelto (S/N)** tiene "SÍ - S" / "SÍ - N". Dividir por " - " en dos columnas: **Respondido** y **Resuelto** (S/N).
4. **Duración de la llamada** viene como texto hh:mm:ss. Cambiar a tipo Hora/Duración, o usar directamente la columna de segundos.
5. Los Id de llamada tienen saltos (ID0005, ID0006 no están): no es un error a corregir, solo sirve para explicar que el conteo se hace con COUNTROWS y no con el último ID.

### Medidas
```
Total Llamadas = COUNTROWS('atenciones')
Satisfacción Promedio = AVERAGE('atenciones'[Índice de satisfacción])
Duración Promedio (seg) = AVERAGE('atenciones'[Duración con asistente en segundos])
Duración Máxima = MAX('atenciones'[Duración con asistente en segundos])
Duración Mínima = MIN('atenciones'[Duración con asistente en segundos])
Velocidad Respuesta Promedio = AVERAGE('atenciones'[Velocidad de respuesta])

Llamadas Resueltas = CALCULATE(COUNTROWS('atenciones'), 'atenciones'[Resuelto] = "S")
% Resolución = DIVIDE([Llamadas Resueltas], [Total Llamadas])

Llamadas Satisfechas = CALCULATE(COUNTROWS('atenciones'), 'atenciones'[Índice de satisfacción] >= 4)
% Satisfechas = DIVIDE([Llamadas Satisfechas], [Total Llamadas])

Objetivo Resolución = 0.9
Cumplimiento Resolución = DIVIDE([% Resolución], [Objetivo Resolución])
```

### Dashboard sugerido
- Tarjetas: Total Llamadas, Satisfacción Promedio, % Resolución.
- KPI: valor = % Resolución, objetivo = Objetivo Resolución, tendencia = Fecha.
- Barras: Total Llamadas por Área. Barras: Satisfacción Promedio por Asistente.
- Líneas o columnas: Total Llamadas por Fecha.
- Segmentadores: Área y Asistente.

### Resultados esperados
| Medida | Valor |
|---|---|
| Total llamadas | 1.009 |
| Satisfacción promedio | 3,40 |
| Duración promedio / máx / mín (seg) | 223,8 / 419 / 30 |
| Velocidad de respuesta promedio | 66,7 |
| Llamadas resueltas / % | 935 / 92,7% |
| Llamadas satisfechas (≥4) / % | 467 / 46,3% |
| Cumplimiento (obj. 90%) | 103% |
| Por área | Ventas 411 · Cancelaciones 208 · Intercambios 199 · Reclamos 191 |

Lectura para el tablero (sirve de "interpretación" en un examen): la resolución cumple el objetivo, pero la satisfacción ronda 3,4 sobre 5. Por asistente, **Juan D. tiene 2,82** frente a ~3,3 a 3,6 del resto, y es quien más llamadas atendió (171). Es el punto a investigar.

---

## Errores típicos de Lucía (para anticipar en la clase)
1. Relación al revés o sin relacionar: los gráficos con datos de 2 tablas muestran todo igual en cada categoría.
2. Confundir columna calculada con medida (usar "Nueva columna" cuando necesita "Nueva medida" y viceversa).
3. KPI con el objetivo mal configurado: si el objetivo es mayor que el valor, el porcentaje da negativo.
4. Tipos de dato mal asignados (fecha como texto, número como texto) → no suma ni promedia.
5. Agregar filas nuevas: hay que editar el Excel y apretar **Actualizar** en Power BI.
6. Mezclar segmentador con filtro.

## Propuesta de clase (2 h)
- 0:00–0:15 Diagnóstico: que ella muestre dónde se traba (con su audio y el modelo de examen).
- 0:15–0:45 Ejercicio 1, partes A y B (importar, limpiar, relacionar). Dibujar el flujo de datos en papel antes de tocar Power BI.
- 0:45–1:15 Columnas vs medidas, CALCULATE (partes C y D).
- 1:15–1:45 Dashboard con KPI y segmentadores (parte E o Ejercicio 2).
- 1:45–2:00 Repaso de los 6 errores y tarea: Ejercicio 2 completo.
