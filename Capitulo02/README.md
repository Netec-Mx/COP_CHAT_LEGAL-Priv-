# Cruce de Información entre Excel y Fuentes Externas — Extracción PDF, Conciliación Bancaria y Clasificación de Rubros Impositivos

## 1. Metadatos

| Atributo | Detalle |
|---|---|
| **Duración total** | 90 minutos (40 min instrucción guiada + 50 min práctica activa) |
| **Complejidad** | Alta (Hard) |
| **Nivel Bloom** | Aplicar (Apply) |
| **Módulo** | Módulo 2 — Temas 2.1 y 2.2 |
| **Sesión** | Sesión 2 |
| **Archivos requeridos** | `Lab02_Libro_Bancos.xlsx` · `Lab02_Estado_Cuenta_Bancario.pdf` |

---

## 2. Descripción General

Este laboratorio aborda uno de los desafíos más frecuentes en auditoría y contabilidad fiscal: la integración y validación de datos provenientes de fuentes heterogéneas. El participante cruzará el **libro auxiliar de bancos** (archivo Excel interno) con los movimientos de un **estado de cuenta bancario oficial en PDF**, utilizando Power Query para la extracción de datos y Microsoft 365 Copilot para la clasificación impositiva y la generación automatizada de la tabla de conciliación. Al concluir, el participante habrá producido un reporte de partidas conciliatorias con indicadores de riesgo fiscal, aplicando criterio profesional para validar y ajustar las sugerencias de Copilot.

> ⚠️ **Aviso de privacidad:** Todos los archivos utilizados son **ficticios**. No cargue información real de clientes, contribuyentes ni expedientes legales en Copilot durante esta práctica.

---

## 3. Objetivos de Aprendizaje

Al finalizar este laboratorio, el participante será capaz de:

- [ ] **Extraer** datos estructurados desde un archivo PDF de estado de cuenta bancario utilizando Power Query en Excel para su procesamiento automatizado.
- [ ] **Realizar** una conciliación bancaria automatizada cruzando el libro auxiliar de bancos con los movimientos del PDF, identificando partidas conciliatorias con asistencia de Copilot.
- [ ] **Clasificar** rubros impositivos (IVA, ISR, retenciones, deducciones) en transacciones importadas mediante prompts de Copilot basados en la descripción y naturaleza de cada operación.
- [ ] **Implementar** flujos de validación de datos cruzados entre múltiples fuentes usando Power Query y Copilot para garantizar integridad y consistencia en la información fiscal consolidada.

---

## 4. Prerrequisitos

### Conocimiento previo

| Área | Nivel requerido |
|---|---|
| Uso básico de Copilot en Excel (Lab 01) | Completado o equivalente |
| Power Query — importación y transformación básica | Básico (sabe conectar a fuente y cambiar tipos de datos) |
| Funciones Excel: `BUSCARX`/`XLOOKUP`, `SI.ERROR`, `COINCIDIR` | Básico |
| Marco impositivo local (IVA, ISR, retenciones) | Conceptual |
| Conceptos de conciliación bancaria (partidas en tránsito, depósitos no registrados) | Conceptual |

### Acceso y configuración previa

| Requisito | Estado esperado |
|---|---|
| Licencia **Microsoft 365 Copilot** asignada | ✅ Activa en tenant |
| Archivos de práctica en **OneDrive** | ✅ Subidos 24 h antes |
| Power Query habilitado | ✅ Verificar: **Datos → Obtener y transformar datos** |
| Copilot habilitado en Excel por el administrador de TI | ✅ Ícono visible en cinta de Excel |
| Microsoft Edge versión 120+ | ✅ Navegador predeterminado |

---

## 5. Entorno de Laboratorio

### Hardware mínimo

| Componente | Especificación |
|---|---|
| Procesador | Intel Core i5 8ª gen / AMD Ryzen 5 o superior |
| RAM | 8 GB mínimo (16 GB recomendado) |
| Almacenamiento libre | 10 GB |
| Resolución de pantalla | 1920 × 1080 |
| Conexión a internet | 10 Mbps mínimo (25 Mbps recomendado) |

### Software requerido

| Aplicación | Versión |
|---|---|
| Microsoft Excel | M365 versión 2401 o superior |
| Microsoft 365 Copilot | Versión de servicio actual (GPT-4 Turbo) |
| Microsoft OneDrive | Versión cloud actual |
| Adobe Acrobat Reader DC | Versión actual (para verificar el PDF fuente) |
| Microsoft Edge | Versión 120+ |

### Preparación del entorno (antes de iniciar)

Ejecute los siguientes pasos de verificación **antes** de comenzar las actividades:

**1. Verificar que los archivos están disponibles en OneDrive:**

```
1. Abrir Microsoft Edge
2. Navegar a: https://onedrive.live.com (o portal corporativo)
3. Confirmar que existen los archivos:
   - Lab02_Libro_Bancos.xlsx
   - Lab02_Estado_Cuenta_Bancario.pdf
4. Si no están presentes, solicitarlos al instructor ANTES de continuar
```

**2. Verificar que Power Query está habilitado en Excel:**

```
1. Abrir Excel (cualquier libro en blanco)
2. Hacer clic en la pestaña: Datos
3. Confirmar que el grupo "Obtener y transformar datos" es visible
4. Si no aparece: Archivo → Opciones → Complementos → Administrar: Complementos COM
   → Verificar que "Microsoft Power Query para Excel" está activo
```

**3. Verificar que Copilot está activo en Excel:**

```
1. Con Excel abierto, buscar el ícono de Copilot en la cinta (pestaña Inicio o Copilot)
2. Hacer clic en el ícono → debe abrirse el panel lateral de Copilot
3. Si aparece un mensaje de "sin licencia" o "no disponible":
   → Notificar al instructor inmediatamente
   → No continuar sin resolver el acceso
```

**4. Abrir los archivos desde OneDrive:**

```
1. Desde Excel: Archivo → Abrir → OneDrive
2. Abrir primero: Lab02_Libro_Bancos.xlsx
3. Mantener abierto durante toda la práctica
4. Tener a la mano (sin abrir en Excel): Lab02_Estado_Cuenta_Bancario.pdf
```

---

## 6. Instrucciones Paso a Paso

> 📋 **Estructura de la sesión:**
> - **Bloque 1 (Pasos 1–3):** Exploración del libro de bancos y preparación de datos — *20 minutos*
> - **Bloque 2 (Pasos 4–6):** Extracción del PDF con Power Query — *25 minutos*
> - **Bloque 3 (Pasos 7–9):** Conciliación bancaria con Copilot — *25 minutos*
> - **Bloque 4 (Paso 10):** Clasificación impositiva y reporte final — *20 minutos*

---

### Paso 1 — Exploración y Estructuración del Libro Auxiliar de Bancos

**Objetivo:** Comprender la estructura del archivo fuente interno y prepararlo como tabla formateada para que Copilot lo reconozca correctamente.

#### Instrucciones

1. Con `Lab02_Libro_Bancos.xlsx` abierto, ubique la hoja denominada **`Libro_Bancos`**.

2. Revise las columnas disponibles. El archivo debe contener al menos las siguientes columnas:

   | Columna esperada | Descripción |
   |---|---|
   | `Fecha` | Fecha del movimiento (formato DD/MM/AAAA) |
   | `Referencia` | Número de cheque, transferencia o referencia interna |
   | `Concepto` | Descripción del movimiento |
   | `Cargo` | Monto de salida (en pesos) |
   | `Abono` | Monto de entrada (en pesos) |
   | `Saldo` | Saldo acumulado |

3. Seleccione el rango de datos completo (incluyendo encabezados). Presione `Ctrl + T` para convertirlo en **Tabla de Excel**.

4. En el cuadro de diálogo, confirme que "La tabla tiene encabezados" está marcado. Haga clic en **Aceptar**.

5. En el campo **Nombre de tabla** (esquina superior izquierda de la cinta, cuando la tabla está seleccionada), escriba exactamente:

   ```
   Libro_Bancos
   ```

6. Verifique que los tipos de datos son correctos:
   - Columna `Fecha`: formato **Fecha** (no texto)
   - Columnas `Cargo`, `Abono`, `Saldo`: formato **Número** o **Moneda**
   - Columnas `Referencia`, `Concepto`: formato **Texto**

7. Si alguna columna de fecha tiene formato de texto, selecciónela y use: **Datos → Texto en columnas → Delimitado → Formato de fecha: DMA**.

8. Guarde el archivo con `Ctrl + S`.

#### Resultado esperado

La hoja `Libro_Bancos` muestra una tabla formateada con bandas de color, el nombre `Libro_Bancos` aparece en el cuadro de nombre de tabla, y todos los tipos de datos son correctos. El número de filas de datos (sin encabezado) debe ser visible en la barra de estado al seleccionar la columna `Fecha`.

#### Verificación

```
✅ La tabla tiene nombre: "Libro_Bancos"
✅ Los datos de fecha son reconocidos como fecha (no como texto)
✅ Las columnas numéricas no contienen texto ni celdas vacías con formato incorrecto
✅ Copilot puede "ver" la tabla: abrir panel de Copilot y escribir:
   "¿Cuántos registros tiene la tabla Libro_Bancos?"
   → Copilot debe responder con el número correcto de filas
```

---

### Paso 2 — Análisis Inicial con Copilot: Estadísticas del Libro de Bancos

**Objetivo:** Usar Copilot para obtener una visión rápida de los datos del libro de bancos antes de realizar el cruce, practicando el uso de prompts descriptivos.

#### Instrucciones

1. Con la tabla `Libro_Bancos` activa, abra el panel de Copilot haciendo clic en el ícono de Copilot en la cinta.

2. Ingrese el siguiente prompt en el panel de Copilot:

   ```
   Analiza la tabla Libro_Bancos y proporciona:
   1. El total de cargos y abonos del período
   2. El número de movimientos por mes
   3. Los 5 conceptos con mayor monto de cargo
   4. Identifica si hay fechas duplicadas con la misma referencia (posibles registros duplicados)
   ```

3. Revise la respuesta de Copilot. Copilot generará fórmulas y posiblemente una tabla resumen. Aplique las sugerencias haciendo clic en **"Insertar columna"** o **"Insertar hoja"** según corresponda.

4. Ingrese un segundo prompt para identificar posibles anomalías:

   ```
   En la tabla Libro_Bancos, identifica movimientos donde el campo
   "Concepto" esté vacío o donde el monto de Cargo o Abono sea igual
   a cero. Agrega una columna llamada "Alerta_Calidad" con el valor
   "REVISAR" para esos registros y "OK" para los demás.
   ```

5. Acepte la columna calculada propuesta por Copilot haciendo clic en **"Insertar columna"**.

6. Filtre la columna `Alerta_Calidad` para ver solo los registros marcados como `REVISAR`. Anote cuántos registros presentan problemas de calidad de datos.

#### Resultado esperado

- Una tabla resumen con totales de cargos y abonos por mes.
- La columna `Alerta_Calidad` agregada a la tabla `Libro_Bancos` con valores `OK` o `REVISAR`.
- Identificación de cualquier registro con datos incompletos que podría afectar la conciliación.

#### Verificación

```
✅ La columna "Alerta_Calidad" aparece en la tabla Libro_Bancos
✅ Los registros con concepto vacío o monto cero están marcados como "REVISAR"
✅ Copilot mostró en el panel el razonamiento utilizado para generar la fórmula
✅ El total de cargos y abonos es coherente con el saldo final del período
```

---

### Paso 3 — Preparación de una Hoja de Trabajo para la Conciliación

**Objetivo:** Crear la estructura base de la hoja de conciliación donde se consolidarán los datos del libro interno y del estado de cuenta bancario.

#### Instrucciones

1. En el mismo libro `Lab02_Libro_Bancos.xlsx`, inserte una nueva hoja haciendo clic derecho en la pestaña de hojas → **Insertar → Hoja de cálculo**. Nómbrela:

   ```
   Conciliacion_Bancaria
   ```

2. En la hoja `Conciliacion_Bancaria`, cree manualmente los siguientes encabezados en la fila 1 (celdas A1 a I1):

   ```
   A1: Fecha
   B1: Referencia
   C1: Concepto
   D1: Monto_Libro
   E1: Monto_Banco
   F1: Diferencia
   G1: Estado_Conciliacion
   H1: Rubro_Impositivo
   I1: Observaciones
   ```

3. Formatee estos encabezados como tabla: seleccione A1:I1, aplique negrita y color de fondo azul oscuro con texto blanco.

4. Deje las filas 2 en adelante vacías por ahora. Los datos se poblarán en pasos posteriores.

5. En la celda **A1 de una sección de notas** (por ejemplo, celda `K1`), escriba:

   ```
   Hoja preparada para conciliación bancaria - Lab02
   Período: Enero-Marzo [año fiscal de práctica]
   ```

6. Guarde el archivo con `Ctrl + S`.

#### Resultado esperado

La hoja `Conciliacion_Bancaria` está creada con los 9 encabezados definidos, lista para recibir los datos cruzados en pasos posteriores.

#### Verificación

```
✅ La hoja "Conciliacion_Bancaria" existe en el libro
✅ Los encabezados A1:I1 están escritos exactamente como se indica
✅ El archivo está guardado en OneDrive (verificar ícono de sincronización en barra de título)
```

---

### Paso 4 — Extracción de Datos del PDF con Power Query

**Objetivo:** Importar los movimientos del estado de cuenta bancario en PDF hacia Excel usando Power Query, transformando los datos para que sean comparables con el libro interno.

> ⏱️ **Nota de tiempo:** Este paso es el más técnico del laboratorio. El instructor puede demostrar los primeros 3 sub-pasos antes de que los participantes los ejecuten de forma independiente.

#### Instrucciones

1. En el libro `Lab02_Libro_Bancos.xlsx`, vaya a la pestaña **Datos** → **Obtener datos** → **Desde archivo** → **Desde PDF**.

2. En el explorador de archivos, navegue a su carpeta de OneDrive y seleccione el archivo:

   ```
   Lab02_Estado_Cuenta_Bancario.pdf
   ```

   Haga clic en **Importar**.

3. Se abrirá el **Navegador de Power Query**. En el panel izquierdo verá las páginas y tablas detectadas en el PDF. Busque la tabla que contenga los movimientos bancarios (generalmente aparece como `Table001` o `Página 1 - Tabla 1`). Selecciónela para ver la vista previa.

4. Verifique que la tabla detectada contiene columnas similares a:

   ```
   Fecha | Descripción | Referencia | Débito | Crédito | Saldo
   ```

   > ⚠️ **Advertencia:** Power Query puede detectar múltiples tablas en el PDF (encabezados, resúmenes, etc.). Seleccione **únicamente** la tabla de movimientos detallados. Si el PDF tiene múltiples páginas con movimientos, mantenga `Ctrl` presionado y seleccione todas las tablas de movimientos.

5. Haga clic en **Transformar datos** (NO en "Cargar"). Se abrirá el **Editor de Power Query**.

6. En el Editor de Power Query, realice las siguientes transformaciones:

   **a) Promover encabezados (si es necesario):**
   Si la primera fila contiene los nombres de columna como datos, vaya a:
   ```
   Inicio → Usar la primera fila como encabezado
   ```

   **b) Eliminar filas innecesarias:**
   Identifique y elimine filas que contengan texto de encabezado repetido, totales intermedios o filas vacías:
   ```
   Inicio → Quitar filas → Quitar filas en blanco
   ```
   Para filas con texto de encabezado repetido: haga clic derecho en la fila → **Quitar**.

   **c) Cambiar tipos de datos:**
   - Columna `Fecha`: clic derecho → **Cambiar tipo** → **Fecha**
   - Columnas `Débito`, `Crédito`, `Saldo`: clic derecho → **Cambiar tipo** → **Número decimal**
   - Columnas `Descripción`, `Referencia`: verificar que sean **Texto**

   **d) Renombrar columnas para estandarizar:**
   Haga doble clic en cada encabezado de columna y renómbrelas así:

   ```
   Fecha_Banco
   Descripcion_Banco
   Referencia_Banco
   Debito_Banco
   Credito_Banco
   Saldo_Banco
   ```

   **e) Reemplazar valores nulos en columnas numéricas:**
   Seleccione la columna `Debito_Banco` → **Transformar → Reemplazar valores**:
   ```
   Valor a buscar: null
   Reemplazar por: 0
   ```
   Repita para `Credito_Banco`.

7. Verifique en la barra inferior del Editor de Power Query que el número de filas cargadas corresponde al total de movimientos del período (3 meses).

8. Haga clic en **Inicio → Cerrar y cargar en...** (flecha desplegable). En el cuadro de diálogo:
   - Seleccione: **Tabla**
   - Ubicación: **Hoja de cálculo existente**
   - Celda: seleccione la hoja `Estado_Cuenta_PDF` (créela si no existe) → celda `$A$1`
   - Haga clic en **Aceptar**.

9. Power Query cargará los datos en una nueva hoja llamada `Estado_Cuenta_PDF`. Verifique que la tabla se llame `Estado_Cuenta_PDF` (ajuste el nombre en el cuadro de nombre de tabla si es necesario).

10. Guarde el archivo con `Ctrl + S`.

#### Resultado esperado

Una nueva hoja `Estado_Cuenta_PDF` con los movimientos bancarios extraídos del PDF, organizados en una tabla con las 6 columnas renombradas, tipos de datos correctos y sin filas vacías ni encabezados repetidos.

#### Verificación

```
✅ La hoja "Estado_Cuenta_PDF" existe con datos cargados desde el PDF
✅ La columna "Fecha_Banco" muestra fechas en formato de fecha (no texto)
✅ Las columnas "Debito_Banco" y "Credito_Banco" son numéricas (sin "$" ni comas como texto)
✅ No hay filas con encabezados repetidos dentro de los datos
✅ El panel "Consultas y conexiones" (Datos → Consultas y conexiones) muestra la consulta PDF activa
✅ Copilot puede leer la tabla: ingresar en Copilot:
   "¿Cuántos movimientos tiene la tabla Estado_Cuenta_PDF?"
   → Debe responder con el número correcto
```

---

### Paso 5 — Normalización de Datos para el Cruce con Copilot

**Objetivo:** Usar Copilot para identificar y resolver diferencias de formato entre las dos tablas antes de ejecutar el cruce, garantizando que los campos clave sean comparables.

#### Instrucciones

1. Abra el panel de Copilot y escriba el siguiente prompt de diagnóstico:

   ```
   Tengo dos tablas en este libro:
   - "Libro_Bancos" con columnas: Fecha, Referencia, Concepto, Cargo, Abono, Saldo
   - "Estado_Cuenta_PDF" con columnas: Fecha_Banco, Descripcion_Banco,
     Referencia_Banco, Debito_Banco, Credito_Banco, Saldo_Banco

   ¿Qué columnas debería usar como campos clave para cruzar ambas tablas?
   ¿Identificas algún posible problema de compatibilidad de datos entre ellas?
   Dame recomendaciones antes de ejecutar el cruce.
   ```

2. Lea cuidadosamente la respuesta de Copilot. Típicamente sugerirá usar `Fecha` + `Referencia` como campos clave compuestos, y puede alertar sobre diferencias de formato de fecha o espacios en blanco en referencias.

3. Basándose en las recomendaciones de Copilot, ingrese el siguiente prompt de limpieza:

   ```
   En la tabla Libro_Bancos, agrega una columna llamada "Referencia_Normalizada"
   que tome el valor de la columna "Referencia", elimine espacios al inicio y
   al final, y convierta todo a mayúsculas.

   En la tabla Estado_Cuenta_PDF, agrega una columna llamada "Referencia_Normalizada"
   con la misma transformación aplicada a "Referencia_Banco".
   ```

4. Aplique las columnas calculadas propuestas por Copilot haciendo clic en **"Insertar columna"** para cada tabla.

5. Haga una verificación manual rápida: compare visualmente 5 referencias de ambas tablas para confirmar que el formato es ahora idéntico (mismas mayúsculas, sin espacios extra).

6. Si Copilot detectó diferencias de formato en fechas, ingrese:

   ```
   Verifica que las fechas en la columna "Fecha" de la tabla Libro_Bancos
   y las fechas en "Fecha_Banco" de Estado_Cuenta_PDF estén en el mismo
   formato. Si hay diferencias, genera las fórmulas necesarias para
   estandarizarlas.
   ```

7. Guarde el archivo con `Ctrl + S`.

#### Resultado esperado

Ambas tablas tienen una columna `Referencia_Normalizada` con valores en mayúsculas y sin espacios, lista para ser usada como campo clave en el cruce. Los formatos de fecha son compatibles entre ambas fuentes.

#### Verificación

```
✅ La columna "Referencia_Normalizada" existe en AMBAS tablas
✅ Los valores en "Referencia_Normalizada" son idénticos para un registro
   que existe en ambas tablas (verificación manual de muestra)
✅ No hay diferencias de formato de fecha visibles entre las dos tablas
```

---

### Paso 6 — Cruce de Tablas con BUSCARX y Copilot

**Objetivo:** Ejecutar el cruce formal entre el libro de bancos y el estado de cuenta bancario, identificando partidas conciliatorias mediante fórmulas generadas por Copilot.

#### Instrucciones

1. Vaya a la hoja `Conciliacion_Bancaria`.

2. En el panel de Copilot, ingrese el siguiente prompt:

   ```
   Necesito poblar la hoja "Conciliacion_Bancaria" con datos cruzados.

   La tabla fuente principal es "Libro_Bancos". Para cada registro del
   Libro_Bancos, necesito:

   1. Buscar en "Estado_Cuenta_PDF" un registro que coincida en
      "Referencia_Normalizada" Y tenga una fecha dentro de ±3 días.

   2. Si se encuentra coincidencia: traer el valor de "Debito_Banco"
      o "Credito_Banco" según corresponda al campo "Monto_Banco".

   3. Calcular la diferencia: Monto_Libro menos Monto_Banco.

   4. Asignar "Estado_Conciliacion":
      - "CONCILIADO" si la diferencia es 0
      - "DIFERENCIA" si la diferencia es distinta de 0 pero hay coincidencia
      - "SOLO EN LIBRO" si no se encontró en el estado de cuenta

   Genera las fórmulas necesarias para implementar esto en la hoja
   Conciliacion_Bancaria.
   ```

3. Copilot generará un conjunto de fórmulas. Revíselas antes de aplicarlas. Las fórmulas típicamente incluirán combinaciones de `BUSCARX`/`XLOOKUP` con `SI.ERROR`:

   ```excel
   // Ejemplo de fórmula para Monto_Banco (columna E)
   // generada por Copilot — puede variar según la versión
   =IFERROR(
     XLOOKUP(
       [@Referencia_Normalizada],
       Estado_Cuenta_PDF[Referencia_Normalizada],
       Estado_Cuenta_PDF[Debito_Banco] + Estado_Cuenta_PDF[Credito_Banco]
     ),
     0
   )
   ```

   ```excel
   // Ejemplo de fórmula para Estado_Conciliacion (columna G)
   =IF(
     [@Monto_Banco] = 0,
     "SOLO EN LIBRO",
     IF(
       ABS([@Monto_Libro] - [@Monto_Banco]) < 0.01,
       "CONCILIADO",
       "DIFERENCIA"
     )
   )
   ```

4. Copie los datos del `Libro_Bancos` a la hoja `Conciliacion_Bancaria`:
   - Columna A (`Fecha`): desde `Libro_Bancos[Fecha]`
   - Columna B (`Referencia`): desde `Libro_Bancos[Referencia_Normalizada]`
   - Columna C (`Concepto`): desde `Libro_Bancos[Concepto]`
   - Columna D (`Monto_Libro`): desde `Libro_Bancos[Cargo]` o `Libro_Bancos[Abono]` según corresponda

   > **Sugerencia de prompt:** Si prefiere que Copilot haga esto automáticamente, ingrese:
   > ```
   > Crea una fórmula para poblar la columna A de la hoja Conciliacion_Bancaria
   > con las fechas de la tabla Libro_Bancos, y continúa con las demás columnas
   > de datos fuente.
   > ```

5. Aplique las fórmulas de cruce en las columnas E (`Monto_Banco`), F (`Diferencia`) y G (`Estado_Conciliacion`).

6. Ahora identifique los movimientos del estado de cuenta que **no están** en el libro de bancos. Ingrese en Copilot:

   ```
   Identifica los registros de la tabla "Estado_Cuenta_PDF" cuya
   "Referencia_Normalizada" NO aparece en la tabla "Libro_Bancos".
   Estos son movimientos bancarios no registrados contablemente.
   Agrega esos registros al final de la hoja "Conciliacion_Bancaria"
   con el Estado_Conciliacion = "SOLO EN BANCO".
   ```

7. Aplique las sugerencias de Copilot para agregar esas filas.

8. Guarde el archivo con `Ctrl + S`.

#### Resultado esperado

La hoja `Conciliacion_Bancaria` contiene todos los registros del libro de bancos con su estado de conciliación (`CONCILIADO`, `DIFERENCIA`, `SOLO EN LIBRO`), más los registros adicionales del estado de cuenta marcados como `SOLO EN BANCO`.

#### Verificación

```
✅ La columna G muestra los cuatro estados posibles:
   CONCILIADO | DIFERENCIA | SOLO EN LIBRO | SOLO EN BANCO
✅ Los registros "CONCILIADO" tienen diferencia = 0 (o < $0.01 por redondeo)
✅ Los registros "DIFERENCIA" tienen un monto distinto en Monto_Libro vs Monto_Banco
✅ Los registros "SOLO EN BANCO" provienen de Estado_Cuenta_PDF y tienen Monto_Libro = 0
✅ La suma de Monto_Libro de registros CONCILIADOS + DIFERENCIA es coherente
   con el total del Libro_Bancos
```

---

### Paso 7 — Generación del Resumen de Conciliación con Tabla Dinámica

**Objetivo:** Crear un resumen ejecutivo de la conciliación bancaria usando una tabla dinámica, asistida por Copilot para su configuración.

#### Instrucciones

1. Con la hoja `Conciliacion_Bancaria` activa, abra el panel de Copilot e ingrese:

   ```
   Crea una tabla dinámica en una nueva hoja llamada "Resumen_Conciliacion"
   que muestre:
   - Filas: Estado_Conciliacion
   - Valores: Conteo de registros, Suma de Monto_Libro, Suma de Monto_Banco,
     Suma de Diferencia
   - Formato de los valores monetarios con separador de miles y 2 decimales
   ```

2. Copilot puede guiarle con instrucciones o generar la tabla dinámica directamente. Si Copilot genera instrucciones, sígalas:
   - **Insertar → Tabla dinámica**
   - Seleccione la tabla `Conciliacion_Bancaria` como fuente
   - Ubíquela en una nueva hoja llamada `Resumen_Conciliacion`

3. Configure la tabla dinámica con los campos indicados por Copilot.

4. Agregue un segundo prompt para obtener el resumen narrativo:

   ```
   Basándote en los datos de la hoja Conciliacion_Bancaria, redacta un
   párrafo de resumen ejecutivo de la conciliación bancaria que incluya:
   - Total de movimientos analizados
   - Porcentaje conciliado
   - Monto total de diferencias identificadas
   - Número de partidas abiertas (SOLO EN LIBRO + SOLO EN BANCO)
   - Una recomendación de acción para las partidas con diferencias
   ```

5. Copie el texto generado por Copilot en la celda `A1` de la hoja `Resumen_Conciliacion`, por encima de la tabla dinámica.

6. Guarde el archivo con `Ctrl + S`.

#### Resultado esperado

Una hoja `Resumen_Conciliacion` con una tabla dinámica que muestra el conteo y montos por estado de conciliación, y un párrafo de resumen ejecutivo generado por Copilot.

#### Verificación

```
✅ La tabla dinámica muestra las 4 categorías de Estado_Conciliacion
✅ La suma de todos los registros en la tabla dinámica coincide con el total
   de filas en Conciliacion_Bancaria
✅ El resumen ejecutivo menciona cifras coherentes con los datos
✅ El porcentaje conciliado es calculable: (N° CONCILIADOS / Total) × 100
```

---

### Paso 8 — Clasificación Impositiva con Copilot

**Objetivo:** Usar Copilot para clasificar automáticamente cada transacción de la conciliación según su naturaleza impositiva, aplicando criterio profesional para validar las sugerencias.

#### Instrucciones

1. Regrese a la hoja `Conciliacion_Bancaria`. La columna H (`Rubro_Impositivo`) está vacía. Vamos a poblarla.

2. Ingrese el siguiente prompt en Copilot:

   ```
   Analiza la columna "Concepto" de la hoja Conciliacion_Bancaria y
   clasifica cada transacción en la columna "Rubro_Impositivo" usando
   las siguientes categorías:

   - "IVA_TRASLADADO": cobros a clientes que incluyen IVA
   - "IVA_ACREDITABLE": pagos a proveedores con IVA acreditable
   - "RETENCION_ISR": retenciones de ISR aplicadas
   - "RETENCION_IVA": retenciones de IVA aplicadas
   - "PAGO_IMPUESTO": pagos directos al fisco (SAT, autoridad tributaria)
   - "NOMINA": pagos de nómina y prestaciones laborales
   - "INVERSION_ACTIVO": adquisición de activos fijos o inversiones
   - "GASTO_DEDUCIBLE": gastos operativos deducibles del ISR
   - "GASTO_NO_DEDUCIBLE": gastos que no cumplen requisitos de deducción
   - "OPERACION_NO_CLASIFICADA": cuando no hay suficiente información

   Genera la fórmula o el proceso para clasificar automáticamente basándote
   en palabras clave en el campo Concepto.
   ```

3. Copilot generará una fórmula con `SI` anidados o `BUSCARV` con una tabla de palabras clave. Una respuesta típica puede incluir:

   ```excel
   // Ejemplo de clasificación por palabras clave (simplificado)
   // Copilot puede generar una versión más completa

   =IF(ISNUMBER(SEARCH("IVA COBRADO",[@Concepto])),"IVA_TRASLADADO",
    IF(ISNUMBER(SEARCH("RETENCION ISR",[@Concepto])),"RETENCION_ISR",
    IF(ISNUMBER(SEARCH("NOMINA",[@Concepto])),"NOMINA",
    IF(ISNUMBER(SEARCH("SAT",[@Concepto])),"PAGO_IMPUESTO",
    IF(ISNUMBER(SEARCH("PROVEEDOR",[@Concepto])),"IVA_ACREDITABLE",
    "OPERACION_NO_CLASIFICADA")))))
   ```

4. Aplique la fórmula a la columna H. Revise manualmente al menos **10 registros** para validar que la clasificación sea correcta según su criterio profesional.

5. Para los registros clasificados como `OPERACION_NO_CLASIFICADA`, ingrese un prompt de refinamiento:

   ```
   Hay registros en Conciliacion_Bancaria donde Rubro_Impositivo =
   "OPERACION_NO_CLASIFICADA". Muéstrame los conceptos únicos de esos
   registros y sugiere una clasificación impositiva para cada uno,
   explicando el razonamiento fiscal.
   ```

6. Basándose en las sugerencias de Copilot y su criterio profesional, actualice manualmente las clasificaciones que considere correctas. **No aplique automáticamente todas las sugerencias sin revisión.**

7. Ingrese un prompt final de validación:

   ```
   En la tabla Conciliacion_Bancaria, genera un resumen que muestre
   cuántos registros y qué monto total corresponde a cada categoría
   de Rubro_Impositivo. Identifica si el total de IVA_TRASLADADO
   menos IVA_ACREDITABLE y RETENCION_IVA da un resultado positivo
   o negativo (para determinar si hay IVA a pagar o a favor).
   ```

8. Guarde el archivo con `Ctrl + S`.

#### Resultado esperado

La columna `Rubro_Impositivo` de la hoja `Conciliacion_Bancaria` está poblada con clasificaciones impositivas para cada transacción. Un resumen muestra los totales por categoría y una indicación del posible saldo de IVA del período.

#### Verificación

```
✅ La columna "Rubro_Impositivo" no tiene celdas vacías (todas tienen alguna categoría)
✅ El porcentaje de "OPERACION_NO_CLASIFICADA" es menor al 15% del total
   (si es mayor, el instructor debe revisar la calidad del archivo PDF fuente)
✅ Las categorías impositivas asignadas son coherentes con los conceptos
   de las transacciones (validación manual de muestra de 10 registros)
✅ El resumen de IVA (trasladado vs acreditable) es calculable y tiene sentido fiscal
```

---

### Paso 9 — Identificación de Riesgos y Partidas de Alto Valor

**Objetivo:** Usar Copilot para detectar patrones de riesgo en las partidas no conciliadas y diferencias de alto valor que requieren atención prioritaria.

#### Instrucciones

1. En el panel de Copilot, ingrese el siguiente prompt de análisis de riesgo:

   ```
   En la hoja Conciliacion_Bancaria, identifica y marca con una columna
   adicional llamada "Nivel_Riesgo" los siguientes escenarios:

   - "ALTO": registros con Estado_Conciliacion = "DIFERENCIA" y
     monto de diferencia mayor a $10,000
   - "ALTO": registros con Estado_Conciliacion = "SOLO EN BANCO" y
     Rubro_Impositivo = "PAGO_IMPUESTO" (posibles pagos fiscales no registrados)
   - "MEDIO": registros con Estado_Conciliacion = "DIFERENCIA" y
     diferencia entre $1,000 y $10,000
   - "MEDIO": registros con Estado_Conciliacion = "SOLO EN LIBRO" y
     monto mayor a $5,000
   - "BAJO": todos los demás registros no conciliados
   - "N/A": registros con Estado_Conciliacion = "CONCILIADO"
   ```

2. Aplique la columna `Nivel_Riesgo` generada por Copilot.

3. Aplique formato condicional a la columna `Nivel_Riesgo`:
   - `ALTO` → fondo rojo, texto blanco
   - `MEDIO` → fondo amarillo, texto negro
   - `BAJO` → fondo naranja claro
   - `N/A` → fondo verde claro

   Puede solicitar esto a Copilot:
   ```
   Aplica formato condicional a la columna "Nivel_Riesgo" en la hoja
   Conciliacion_Bancaria: ALTO en rojo, MEDIO en amarillo, BAJO en
   naranja y N/A en verde.
   ```

4. Ingrese un prompt final de síntesis:

   ```
   Basándote en la hoja Conciliacion_Bancaria, redacta un listado de
   las 5 partidas de mayor riesgo que requieren investigación inmediata,
   indicando para cada una: fecha, referencia, monto, estado de
   conciliación y razón del riesgo. Usa un tono profesional apropiado
   para un informe de auditoría.
   ```

5. Copie el listado generado en la columna `I` (`Observaciones`) de los registros correspondientes, o en una celda de notas en la hoja `Resumen_Conciliacion`.

6. Guarde el archivo con `Ctrl + S`.

#### Resultado esperado

La hoja `Conciliacion_Bancaria` tiene la columna `Nivel_Riesgo` con formato condicional aplicado, y la hoja `Resumen_Conciliacion` incluye el listado de las 5 partidas prioritarias para investigación.

#### Verificación

```
✅ La columna "Nivel_Riesgo" existe y tiene valores en todas las filas
✅ El formato condicional es visible (celdas rojas, amarillas, naranjas y verdes)
✅ Los registros marcados como "ALTO" corresponden efectivamente a diferencias
   o partidas de mayor impacto económico
✅ El listado de las 5 partidas prioritarias es coherente con los datos
   y está redactado en tono profesional
```

---

### Paso 10 — Generación del Reporte Final de Conciliación

**Objetivo:** Consolidar todos los resultados en un reporte ejecutivo completo usando Copilot, que sirva como entregable profesional de la conciliación bancaria con clasificación impositiva.

#### Instrucciones

1. Inserte una nueva hoja y nómbrela `Reporte_Ejecutivo`.

2. En el panel de Copilot, ingrese el siguiente prompt de generación de reporte:

   ```
   Genera el contenido para un "Reporte Ejecutivo de Conciliación Bancaria
   y Análisis Impositivo" que incluya las siguientes secciones:

   1. RESUMEN EJECUTIVO: 3-4 oraciones con los hallazgos principales
   2. ESTADÍSTICAS DE CONCILIACIÓN: tabla con conteos y montos por estado
   3. ANÁLISIS IMPOSITIVO: distribución de transacciones por rubro
      impositivo y estimación de impacto fiscal (IVA neto, retenciones)
   4. PARTIDAS DE RIESGO: las 5 partidas prioritarias identificadas
   5. RECOMENDACIONES: 3-5 acciones concretas para el equipo fiscal
      y contable basadas en los hallazgos
   6. PRÓXIMOS PASOS: cronograma sugerido de seguimiento

   Usa los datos reales de las hojas Conciliacion_Bancaria y
   Resumen_Conciliacion de este libro para fundamentar cada sección.
   ```

3. Copie el contenido generado por Copilot en la hoja `Reporte_Ejecutivo`.

4. Dé formato profesional al reporte:
   - Título en fuente grande (18-20 pt), negrita
   - Secciones con encabezados en negrita y color corporativo
   - Tablas con bordes y encabezados sombreados

5. Agregue en la parte superior del reporte:

   ```
   Empresa: [Nombre ficticio del archivo de práctica]
   Período analizado: Enero - Marzo [año]
   Preparado por: [Su nombre]
   Fecha de elaboración: [Fecha actual]
   Herramientas utilizadas: Microsoft 365 Copilot + Power Query + Excel
   NOTA: Este documento es un ejercicio de práctica con datos ficticios
   ```

6. Realice una **revisión crítica** del reporte generado por Copilot:
   - ¿Las cifras mencionadas coinciden con los datos reales de las hojas?
   - ¿Las recomendaciones son aplicables al contexto fiscal de su país?
   - ¿Hay alguna afirmación que requiera verificación adicional?

   Anote al menos **2 observaciones o correcciones** que realizó al reporte de Copilot. Esto refuerza el principio de **criterio profesional sobre la IA**.

7. Guarde el archivo final con `Ctrl + S`.

8. Verifique que el archivo está sincronizado en OneDrive (el ícono en la barra de título debe mostrar la nube con palomita ✓).

#### Resultado esperado

La hoja `Reporte_Ejecutivo` contiene un reporte profesional completo con las 6 secciones solicitadas, cifras coherentes con los datos del libro, y al menos 2 correcciones o mejoras aplicadas con criterio propio del participante.

#### Verificación

```
✅ La hoja "Reporte_Ejecutivo" existe con las 6 secciones del reporte
✅ Las cifras del reporte son coherentes con los datos en Conciliacion_Bancaria
✅ El participante documentó al menos 2 correcciones al output de Copilot
✅ El archivo está guardado y sincronizado en OneDrive
✅ El libro final contiene las hojas:
   Libro_Bancos | Estado_Cuenta_PDF | Conciliacion_Bancaria |
   Resumen_Conciliacion | Reporte_Ejecutivo
```

---

## 7. Validación y Pruebas Finales

Al concluir todos los pasos, realice las siguientes verificaciones de integridad del trabajo completo:

### Lista de verificación final

| # | Verificación | Criterio de éxito |
|---|---|---|
| 1 | **Estructura del libro** | El libro contiene exactamente 5 hojas con los nombres especificados |
| 2 | **Extracción PDF** | La tabla `Estado_Cuenta_PDF` tiene datos con tipos correctos (fechas como fecha, montos como número) |
| 3 | **Cobertura de conciliación** | El total de registros en `Conciliacion_Bancaria` = registros en `Libro_Bancos` + registros "SOLO EN BANCO" |
| 4 | **Consistencia de montos** | La suma de `Monto_Libro` de registros CONCILIADOS + DIFERENCIA = Total de cargos/abonos del `Libro_Bancos` |
| 5 | **Clasificación impositiva** | Menos del 15% de registros clasificados como "OPERACION_NO_CLASIFICADA" |
| 6 | **Formato condicional** | La columna `Nivel_Riesgo` muestra colores correctos para cada categoría |
| 7 | **Criterio profesional** | El participante documentó al menos 2 correcciones al output de Copilot |
| 8 | **Sincronización OneDrive** | El archivo está guardado y sincronizado (sin ícono de error en barra de título) |

### Prueba de integridad del cruce

Para confirmar que el cruce es correcto, ejecute manualmente la siguiente verificación:

```excel
// En una celda vacía de cualquier hoja, ingrese:
// Esta fórmula debe dar 0 si la conciliación es íntegra

=COUNTIF(Conciliacion_Bancaria[Estado_Conciliacion],"CONCILIADO")
+COUNTIF(Conciliacion_Bancaria[Estado_Conciliacion],"DIFERENCIA")
+COUNTIF(Conciliacion_Bancaria[Estado_Conciliacion],"SOLO EN LIBRO")
-COUNTA(Libro_Bancos[Referencia])

// Resultado esperado: 0
// Si el resultado es diferente de 0, hay registros del Libro_Bancos
// que no fueron incluidos en la conciliación
```

---

## 8. Resolución de Problemas

### Problema 1: Power Query no detecta tablas en el PDF o extrae datos desordenados

**Síntomas:**
- El Navegador de Power Query muestra el PDF pero no detecta tablas estructuradas, o detecta tablas con datos mezclados, columnas desplazadas o filas fusionadas incorrectamente.
- Los datos importados tienen columnas con nombres como `Column1`, `Column2` en lugar de los encabezados del estado de cuenta.
- Algunos movimientos aparecen en una sola celda en lugar de distribuidos en columnas.

**Causa:**
El PDF del estado de cuenta fue generado como imagen escaneada (no como texto seleccionable), o el formato del PDF tiene tablas con celdas combinadas o diseño complejo que confunde al motor de extracción de Power Query. Power Query solo puede extraer texto seleccionable de PDFs; los PDFs escaneados requieren OCR.

**Solución:**
```
Paso 1 — Verificar si el PDF tiene texto seleccionable:
  1. Abrir el PDF en Adobe Acrobat Reader
  2. Intentar seleccionar texto con el cursor
  3. Si no se puede seleccionar texto → el PDF es imagen escaneada

Paso 2a — Si el PDF es imagen escaneada:
  1. Usar Adobe Acrobat Pro (si disponible) para OCR:
     Herramientas → Reconocer texto → En este archivo
  2. Guardar el PDF resultante y reintentar la importación en Power Query

Paso 2b — Si el PDF tiene texto pero las tablas están mal detectadas:
  1. En el Editor de Power Query, después de importar:
     → Inicio → Usar la primera fila como encabezado (si aplica)
     → Transformar → Detectar tipo de datos
  2. Si las columnas siguen desordenadas, use:
     → Transformar → Dividir columna → Por delimitador
     para separar datos que quedaron en una sola columna

Paso 3 — Alternativa manual si Power Query falla completamente:
  1. Copiar los datos del PDF seleccionando el texto en Acrobat Reader
  2. Pegarlo en Excel con Pegado especial → Solo texto
  3. Usar Datos → Texto en columnas para estructurar
  4. Dar formato de tabla manualmente con el nombre "Estado_Cuenta_PDF"
  5. Continuar el laboratorio desde el Paso 5

Paso 4 — Solicitar ayuda a Copilot:
  Ingrese en Copilot: "Los datos del PDF quedaron en una sola columna.
  Ayúdame a separar la columna A en las columnas:
  Fecha_Banco, Descripcion_Banco, Referencia_Banco, Debito_Banco,
  Credito_Banco, Saldo_Banco usando el patrón de los datos."
```

---

### Problema 2: Copilot no puede leer las tablas del libro o responde con información incorrecta

**Síntomas:**
- Copilot responde "No encontré datos en este libro" o "No tengo acceso a las tablas".
- Copilot genera fórmulas que hacen referencia a nombres de tablas o columnas incorrectos (por ejemplo, usa `Sheet1[Column1]` en lugar de `Libro_Bancos[Referencia]`).
- El panel de Copilot muestra un error de conexión o un mensaje de "servicio no disponible".
- Las fórmulas generadas por Copilot producen errores `#¿NOMBRE?` o `#REF!` al aplicarlas.

**Causa:**
Copilot en Excel requiere que los datos estén formateados como **Tabla de Excel** con nombre asignado. Si los datos son rangos simples (sin formato de tabla), Copilot no los reconoce correctamente. También puede ocurrir que la sesión de Copilot haya expirado o que el libro no esté guardado en OneDrive (Copilot requiere que el archivo esté en la nube para funcionar con todas sus capacidades).

**Solución:**
```
Paso 1 — Verificar que los datos son tablas formateadas:
  1. Hacer clic en cualquier celda del rango de datos
  2. Si NO aparece la pestaña "Diseño de tabla" en la cinta → el rango
     NO es una tabla
  3. Solución: seleccionar el rango completo → Ctrl+T → Aceptar
  4. Asignar el nombre correcto en "Nombre de tabla" (esquina superior izquierda)

Paso 2 — Verificar que el archivo está en OneDrive:
  1. Revisar la barra de título de Excel
  2. Si dice "Guardado en este equipo" → el archivo está local
  3. Solución: Archivo → Guardar una copia → OneDrive
  4. Cerrar el archivo local y abrir desde OneDrive

Paso 3 — Reiniciar la sesión de Copilot:
  1. Cerrar el panel de Copilot (X en la esquina del panel)
  2. Guardar el archivo (Ctrl+S)
  3. Esperar 30 segundos
  4. Reabrir Copilot desde la cinta

Paso 4 — Si Copilot usa nombres de tabla incorrectos en fórmulas:
  1. Antes de aplicar la fórmula, edítela manualmente en la barra de fórmulas
  2. Reemplazar los nombres incorrectos por los nombres correctos:
     "Libro_Bancos" y "Estado_Cuenta_PDF"
  3. Verificar que los nombres de columna entre corchetes coinciden
     exactamente con los encabezados de las tablas

Paso 5 — Verificar la licencia si el problema persiste:
  1. Ir a: https://office.com → iniciar sesión
  2. Verificar que aparece el ícono de Copilot en el portal
  3. Si no aparece → notificar al instructor para verificar la asignación
     de licencia Microsoft 365 Copilot en el tenant
```

---

## 9. Limpieza del Entorno

Al finalizar el laboratorio, realice los siguientes pasos para dejar el entorno ordenado:

1. **Guardar el archivo final:**
   ```
   Ctrl + S → Verificar que el ícono de OneDrive en la barra de título
   muestra sincronización completada (nube con palomita ✓)
   ```

2. **Eliminar columnas de trabajo temporales** (si las creó fuera de las tablas principales):
   ```
   Revisar si hay columnas auxiliares fuera de las tablas definidas
   (por ejemplo, fórmulas de prueba en columnas sueltas)
   Eliminar cualquier columna o celda temporal que no forme parte
   del entregable final
   ```

3. **Cerrar consultas de Power Query que ya no se necesiten:**
   ```
   Datos → Consultas y conexiones
   Si hay consultas de prueba o duplicadas, hacer clic derecho → Eliminar
   Conservar únicamente la consulta: "Lab02_Estado_Cuenta_Bancario"
   ```

4. **Renombrar el archivo final para entrega:**
   ```
   Archivo → Guardar como
   Nombre sugerido: Lab02_Conciliacion_[SuNombre]_[Fecha].xlsx
   Guardar en la carpeta de OneDrive designada por el instructor
   ```

5. **Cerrar el panel de Copilot** para liberar recursos:
   ```
   Hacer clic en la X del panel lateral de Copilot
   ```

6. **Verificar la lista de archivos en OneDrive:**
   ```
   Confirmar que el archivo final aparece en OneDrive con la fecha
   de modificación correcta (hoy)
   El archivo PDF fuente (Lab02_Estado_Cuenta_Bancario.pdf) puede
   conservarse en OneDrive para referencia futura
   ```

---

## 10. Resumen y Recursos Adicionales

### Resumen de lo aprendido

En este laboratorio aplicó un flujo de trabajo completo de conciliación bancaria asistida por IA, cubriendo cuatro competencias clave:

| Competencia | Herramienta principal | Resultado obtenido |
|---|---|---|
| **Extracción de datos PDF** | Power Query | Tabla `Estado_Cuenta_PDF` con movimientos estructurados |
| **Normalización y cruce de datos** | Copilot + BUSCARX | Columnas `Referencia_Normalizada` y `Estado_Conciliacion` |
| **Clasificación impositiva semántica** | Copilot (prompting) | Columna `Rubro_Impositivo` con 9 categorías fiscales |
| **Reporte ejecutivo automatizado** | Copilot + Tabla dinámica | Hoja `Reporte_Ejecutivo` con análisis de riesgo |

### Principios clave reforzados

- **Copilot como asistente, no como decisor:** Las clasificaciones impositivas y los estados de conciliación generados por Copilot deben ser **validados por el profesional** antes de usarse en declaraciones o informes oficiales.
- **La estructura de datos determina la calidad del resultado:** Tablas formateadas con nombres claros y tipos de datos correctos son el prerrequisito para que Copilot opere con máxima efectividad.
- **Power Query es el puente hacia fuentes heterogéneas:** Dominar la extracción y transformación básica de datos desde PDF, CSV y otras fuentes multiplica el valor de Copilot en contextos de auditoría y fiscalidad.
- **El prompting iterativo mejora los resultados:** Los mejores resultados se obtienen con prompts específicos, contextualizados y en secuencia lógica, no con una sola instrucción genérica.

### Recursos adicionales recomendados

| Recurso | Descripción | Enlace |
|---|---|---|
| Documentación oficial Copilot en Excel | Guía completa de capacidades y limitaciones | [Microsoft Learn — Copilot en Excel](https://learn.microsoft.com/es-es/copilot/microsoft-365/microsoft-365-copilot-excel) |
| Power Query — Importar desde PDF | Tutorial oficial de importación de PDF | [Microsoft Support — Importar PDF](https://support.microsoft.com/es-es/office/importar-datos-de-un-archivo-pdf-power-query-be4330b3-5356-486c-a168-b68e9e616f5a) |
| Introducción a Power Query | Guía de transformación de datos | [Microsoft Support — Power Query](https://support.microsoft.com/es-es/office/introducci%C3%B3n-a-microsoft-power-query-para-excel-6e92e2f4-2079-4e1f-bad5-89f6269cd605) |
| XLOOKUP / BUSCARX | Referencia de función avanzada | [Microsoft Support — XLOOKUP](https://support.microsoft.com/es-es/office/funci%C3%B3n-buscarx-b7fd680e-6d10-43e6-84f9-88eae8bf5929) |
| Mejores prácticas de prompting para finanzas | Guía de prompts para análisis financiero con Copilot | [Microsoft Copilot Adoption Hub](https://adoption.microsoft.com/es-es/copilot/) |

### Preparación para el Laboratorio 03

En el próximo laboratorio trabajará con **Microsoft Word y Copilot** para generar documentación fiscal y legal automatizada: respuestas a requerimientos de autoridades tributarias, minutas de reuniones con clientes y contratos con cláusulas de cumplimiento fiscal. Los archivos de práctica del Lab 03 deben estar en su OneDrive antes de la próxima sesión. Solicítelos al instructor al finalizar esta sesión.

---

> 💡 **Reflexión final del laboratorio:** La inteligencia artificial no reemplaza el criterio fiscal y contable del profesional — lo amplifica. Las herramientas que practicó hoy (Copilot, Power Query, BUSCARX) son tan poderosas como la calidad de los datos que las alimentan y la precisión del juicio profesional que valida sus resultados. El valor diferencial del auditor o asesor fiscal en la era de la IA es exactamente ese: saber **qué preguntar**, **cómo interpretar** y **cuándo cuestionar** lo que la máquina sugiere.

---
