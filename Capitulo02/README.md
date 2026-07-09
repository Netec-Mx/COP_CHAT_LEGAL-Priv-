# Cruce de Información entre Excel y Fuentes Externas — Extracción PDF, Conciliación Bancaria y Clasificación de Rubros Impositivos

# Práctica 2: Laboratorio de Extracción de Datos de PDF y Conciliación Fiscal Inteligente de Rubros Impositivos (90 min)

## 1. Metadatos del Laboratorio

| Atributo | Detalle |
| :--- | :--- |
| **Duración** | 90 minutos (Alta Densidad Operativa, Conciliación Cruzada y Gestión Fiscal) |
| **Complejidad** | Intermedia |
| **Audiencia** | Especialistas en Impuestos, Contadores Fiscales, Auditores de Cumplimiento y Consultores Tributarios |
| **Tecnologías** | Microsoft Copilot (M365 / Interfaz de Chat), Microsoft Excel y Microsoft Word |
| **Enfoque** | Generación de estructuras de datos financieras complejas, simulación de técnicas de extracción analítica de datos sin estructurar (PDF a Excel), reconciliación de flujos de efectivo contra auxiliares contables y clasificación automática de transacciones bajo códigos de rubros impositivos vigentes. |

---

## 2. Descripción Corta

Este laboratorio práctico de 90 minutos capacita a los profesionales contables y fiscales en el uso estratégico de Microsoft Copilot para automatizar el cruce de información entre registros de origen externo no estructurados (Estados de Cuenta en PDF) y los libros auxiliares internos en Microsoft Excel. Los estudiantes aprenderán a utilizar la IA para simular la estructura de un estado de cuenta bancario en PDF, extraer sus tablas transaccionales de forma limpia, aplicar reglas de mapeo en Excel para conciliar los flujos de efectivo y ejecutar una clasificación inteligente de gastos bajo los rubros impositivos del marco fiscal local.

---

## 3. Objetivos del Laboratorio

Al finalizar este laboratorio, el estudiante será capaz de:
* **Generar y estructurar datos sintéticos financieros** en formato PDF para utilizarlos como fuentes externas de auditoría.
* **Extraer tablas y registros transaccionales desde documentos PDF** hacia Excel de forma limpia utilizando instrucciones de Copilot.
* **Diseñar un modelo de conciliación cruzada** en Excel comparando depósitos bancarios contra facturación emitida (XML/Ingresos).
* **Clasificar transacciones de forma inteligente** asignando rubros impositivos (deducibles, no deducibles, retenciones) mediante prompts lógicos.
* **Actualizar de manera autónoma las reglas de clasificación impositiva** ante modificaciones en la miscelánea fiscal o legislación tributaria del país.

---

## 4. Prerrequisitos

* Cuenta activa de **Microsoft 365** con acceso a **Microsoft Copilot Chat**.
* Aplicación de **Microsoft Word** abierta con un documento en blanco guardado como `Estrategia_Conciliacion_Fiscal.docx`.
* Aplicación de **Microsoft Excel** abierta para el procesamiento final de las tablas.

---

## 5. Procedimiento Paso a Paso

### Fase A: Generación Sintética del Estado de Cuenta en PDF

Para simular una auditoría real, el estudiante debe contar con una fuente externa de información financiera. En esta fase, utilizaremos a Copilot para diseñar la estructura exacta de transacciones de un banco que posteriormente utilizaremos como nuestro archivo PDF de origen.

1. Abra el chat general de **Microsoft Copilot**.
2. Introduzca el siguiente prompt estratégico para generar la simulación del estado de cuenta externo:

```
Actúa como un Especialistas en Tecnología Financiera y Sistemas Bancarios. Necesito simular un Estado de Cuenta Corporativo detallado para la empresa "Logística Global", correspondiente al periodo del 01 al 15 de Mayo de 2026. 

Por favor, devuélveme un bloque de texto formal, tabular y perfectamente estructurado, diseñado para ser copiado en Word y guardado inmediatamente como PDF. El documento debe contener una tabla con exactamente 6 registros transaccionales utilizando las siguientes columnas fijas:
- Fecha (Entre el 01 y 15 de mayo de 2026)
- Referencia Bancaria (Código alfanumérico único)
- Descripción de la Operación (Incluye nombres de proveedores, conceptos de cobros de comisiones, pagos de clientes y transferencias)
- Cargo (Salidas de dinero)
- Abono (Entradas de dinero)
- Saldo Acumulado

Asegúrate de incluir escenarios fiscalmente interesantes: un pago de cliente con retención, un cobro de comisión bancaria con IVA indirecto, un gasto por viáticos y un pago a un proveedor extranjero. No agregues saludos conversacionales, entrégame el texto estructurado directamente.
```

3. Seleccione la tabla y datos de cabecera bancaria devueltos por Copilot, cópielos (`Ctrl+C`) y péguelos (`Ctrl+V`) en su archivo de Word (`Estrategia_Conciliacion_Fiscal.docx`).
4. En Word, vaya a **Archivo > Guardar como**, seleccione en el tipo de formato **PDF (*.pdf)** y asígnele el nombre `Estado_Cuenta_Origen_Mayo2026.pdf`. Cierre ese archivo PDF.

---

### Fase B: Extracción Analítica e Importación de Datos PDF a Excel con Copilot

En esta fase independiente, el estudiante asumirá el rol de un consultor fiscal que recibe el PDF externo de la Fase A y necesita transformarlo en datos tabulares estructurados dentro de Microsoft Excel sin recapturar la información de manera manual.

1. Abra la interfaz de **Microsoft Copilot Chat**.
2. Introduzca el siguiente prompt para obtener la estrategia de extracción y limpieza de datos del PDF:

```
Actúa como un Experto en Data Cleansing y Análisis Contable. Tengo un archivo llamado 'Estado_Cuenta_Origen_Mayo2026.pdf' que contiene transacciones bancarias tabuladas. 

Necesito que me indiques la guía paso a paso y las instrucciones exactas en lenguaje natural (Prompt de Extracción) que debo utilizar si decido subir este PDF al chat de Copilot (o usando la funcionalidad 'Obtener datos desde PDF' en Excel) para que la información se extraiga limpia, estructurada en columnas y lista para pegarse en Excel a partir de la celda A1. Explícame cómo evitar que los signos de pesos ($) o las comas de miles rompan el formato de número en Excel de las columnas Cargo y Abono.
```

3. Copie las instrucciones metodológicas brindadas por Copilot y péguelas en la sección "Bitácora de Extracción" de su documento maestro de Word.

---

### Fase C: Prueba de Ejecución – Conciliación Cruzada de Ingresos (Banco vs. Facturación)

Esta fase independiente permite al estudiante ejecutar el cruce de datos contables para verificar que los depósitos recibidos en el banco coincidan exactamente con las facturas fiscales (XML/Ingresos) emitidas por la compañía.

1. En la ventana de chat de Copilot, introduzca el siguiente caso práctico de conciliación para auditoría fiscal:

```
Actúa como Auditor Fiscal Principal. Necesito conciliar las siguientes dos bases de datos de ingresos de la empresa 'Logística Global' del periodo de mayo 2026:

Base A: Depósitos Registrados en el Estado de Cuenta Bancario (Extracción PDF)
- Registro B1: Fecha: 03/05/2026, Ref: TRSF-8902, Monto: $116,000.00 MXN, Descripción: Pago Cliente Alimentos S.A.
- Registro B2: Fecha: 10/05/2026, Ref: TRSF-9112, Monto: $58,000.00 MXN, Descripción: Pago Cliente Constructora Norte

Base B: Facturas Fiscales Electrónicas de Ingresos Emitidas (Registros Internos de Contabilidad)
- Factura F101: Cliente: Alimentos S.A., Subtotal: $100,000.00 MXN, IVA (16%): $16,000.00 MXN, Total: $116,000.00 MXN, Estado: Vigente
- Factura F102: Cliente: Constructora Norte, Subtotal: $50,000.00 MXN, IVA (16%): $8,000.00 MXN, Total: $58,000.00 MXN, Estado: Cancelada

Cruza ambas fuentes de información. Determina si existe alguna discrepancia u omisión fiscal grave, fundamenta el riesgo impositivo de los hallazgos y redacta la conclusión técnica para los papeles de trabajo de auditoría.
```

2. Evalúe analíticamente el resultado arrojó Copilot. La IA debe identificar de inmediato el riesgo del Registro B2: Se recibió y cobró un flujo de efectivo de $58,000 MXN ligado a una factura que contablemente aparece como "Cancelada", lo cual constituye una discrepancia fiscal de alto riesgo (omisión de ingresos acumulables ante la autoridad tributaria).

---

### Fase D: Prueba de Ejecución – Clasificación Inteligente de Rubros Impositivos

Esta fase se enfoca en auditar y etiquetar de forma masiva los egresos o cargos bancarios del estado de cuenta, asignándoles su tratamiento fiscal correspondiente bajo las leyes impositivas vigentes.

1. Introduzca el siguiente comando en el chat de Copilot para simular el motor de clasificación impositiva:

```
Actúa como un Consultor Senior de Impuestos Corporativos. Analiza los siguientes 3 egresos extraídos del Estado de Cuenta en PDF y clasifícalos de forma inteligente en una estructura de tabla.

Egresos a Evaluar:
1. Gasto 1: Cargo de $450.00 MXN por concepto de 'Comisión por Manejo de Cuenta Bancaria' más $72.00 MXN de IVA.
2. Gasto 2: Cargo de $2,300.00 MXN en establecimiento 'Restaurante El Cardenal', concepto: Almuerzo de negocios con clientes locales. No se cuenta con CFDI de la transacción, solo el ticket comercial.
3. Gasto 3: Cargo de $18,900.00 MXN a favor de 'DHL Express', concepto: Pago de flete internacional por importación de muestras de mercancía.

Para cada gasto, la tabla final debe contener:
- Concepto
- Tratamiento Fiscal (Deducible / No Deducible)
- Justificación Legal Base (Considerando leyes de impuesto sobre la renta y requisitos de las deducciones, como la obligatoriedad del comprobante fiscal digital o CFDI).
```

2. Revise la clasificación de Copilot. Valide que declare el Gasto 2 como "No Deducible" debido a la falta de un comprobante fiscal válido (CFDI), cumpliendo estrictamente con las formalidades del marco normativo fiscal.

---

### Fase E: Generación del Reporte de Conciliación y Ajustes Fiscales

Esta fase independiente permite consolidar las conciliaciones de ingresos y egresos en un reporte ejecutivo de ajustes contables que servirá de base para la declaración de impuestos.

1. Ingrese el siguiente prompt en Copilot para estructurar el reporte de cierre:

```
Necesito estructurar el Reporte Formal de Conciliación Bancaria y Ajustes Fiscales de Mayo 2026 para la Dirección de Finanzas de 'Logística Global'. 

Utilizando los hallazgos de las fases C y D, genera un informe técnico en prosa ejecutiva con los siguientes apartados específicos:
1. **Resultado de Conciliación de Ingresos:** Explicación técnica del impacto del depósito cobrado con factura cancelada ($58,000.00 MXN).
2. **Determinación de Egresos No Deducibles:** Resumen de los cargos bancarios rechazados para efectos fiscales y su impacto en la base gravable del impuesto.
3. **Propuesta de Asientos de Ajuste:** Redacta una recomendación de control contable para corregir el estatus de la facturación antes del cierre de la declaración provisional mensual.

Omite comentarios introductorios y entrégame el documento técnico directo para incorporar en la documentación corporativa.
```

2. Copie el texto generado por la herramienta y péguelo en su archivo de Word bajo el título `# Reporte de Ajustes y Conciliación Fiscal`.

---

### Fase F: Reto de Aplicación Autónoma – Ajuste por Reforma Normativa e Impuestos Retenidos

**Instrucciones para el estudiante:** La autoridad legislativa y tributaria del país ha publicado una reforma fiscal de última hora en el Diario Oficial. Esta reforma establece que a partir de este mes, todos los pagos por servicios de fletes internacionales de carga aérea (como el rubro de DHL auditado en la Fase D) están sujetos a una **Retención Obligatoria del 4% del Impuesto sobre la Renta (ISR)** que debe ser declarada de inmediato por la empresa que contrata el servicio.

#### El Desafío:
Redacte un prompt estratégico de forma totalmente autónoma en Copilot para actualizar los criterios del motor de auditoría fiscal conforme a la nueva ley:

1. **Nuevos Parámetros de Control:**
   - Cada vez que el sistema detecte un concepto ligado a "Fletes", "Transporte Internacional" o "Logística de Carga", debe calcular de forma automática una retención del 4% sobre el monto bruto.
   - Debe evaluar si el cargo del estado de cuenta en PDF ya incluye contablemente el desglose de dicha retención o si representa un riesgo de omisión de retenciones por pagar.
2. **Generación Autónoma (Excel):** Pídele a la IA que estructure la fórmula lógica o buscar en Excel que automatice este cálculo en una columna nueva llamada `Columna_Retencion_ISR_4%`.
3. **Prueba de Validación:** Introduzca a Copilot una línea de prueba simulada (ej. Pago de Flete por $10,000 MXN sin desglose de retención) y verifique que la IA emita una alerta de "Riesgo Fiscal: Omisión de Retención de ISR por $400.00 MXN", demostrando que las reglas del laboratorio asimilaron la reforma tributaria de manera autónoma.

---

## 6. Conceptos Clave para Recordar

* **Clasificación Inteligente de Rubros:** Proceso analítico asistido por inteligencia artificial que evalúa la descripción en texto de una transacción financiera y le asigna una categoría fiscal específica (deducible, exento, tasa cero, retención) basada en leyes tributarias preconfiguradas.
* **Cruce Contable-Fiscal (Data Matching):** Técnica de fiscalización masiva que contrasta los movimientos de efectivo reales reportados por las instituciones financieras externas contra los registros de comprobantes fiscales emitidos internamente por la empresa.
* **Discrepancia Fiscal Contable:** Inconsistencia detectada en auditoría donde los ingresos declarados o la facturación vigente no guardan simetría matemática con los flujos de dinero efectivamente depositados en las cuentas bancarias de la entidad.

---

## 7. Resultado Esperado del Estudiante

Para validar la correcta conclusión de esta práctica de 90 minutos, el estudiante consolidará los siguientes componentes:

1. **Archivo de Datos de Origen:** El PDF generado independientemente en la Fase A (`Estado_Cuenta_Origen_Mayo2026.pdf`) que sirvió como insumo externo para el laboratorio.
2. **Archivo `Estrategia_Conciliacion_Fiscal.docx` (Word):**
   * Documento técnico que resguarda las transacciones del estado de cuenta de la Fase A, las directrices de limpieza e importación de la Fase B y el **Reporte Formal de Ajustes Fiscales** finalizado en la Fase E.
3. **Evidencia del Reto Autónomo:**
   * La fórmula recalibrada para el cálculo automático de la retención de fletes del 4% de la Fase F anexada al final del archivo de Word, confirmando que domina la adaptación ágil de sistemas de IA ante reformas tributarias en vivo.
