"""
metodologia.py
==============
Genera documentos Excel descargables con la metodología / lógica
de cada reporte de la aplicación.
"""
from io import BytesIO
import streamlit as st


def _build_excel(title, sections):
    """
    Build a formatted Excel with methodology documentation.
    sections: list of (heading, body_text)
    """
    import xlsxwriter

    output = BytesIO()
    wb = xlsxwriter.Workbook(output)
    ws = wb.add_worksheet("Metodología")

    title_f = wb.add_format({"bold": True, "font_size": 16, "font_color": "#1F3864", "bottom": 2})
    heading_f = wb.add_format({"bold": True, "font_size": 12, "font_color": "#1F3864", "bg_color": "#D6DCE4", "border": 1, "text_wrap": True})
    body_f = wb.add_format({"font_size": 10, "text_wrap": True, "valign": "top", "border": 1})

    ws.set_column(0, 0, 80)
    row = 0
    ws.write(row, 0, title, title_f)
    ws.set_row(row, 30)
    row += 2

    for heading, body in sections:
        ws.write(row, 0, heading, heading_f)
        ws.set_row(row, 22)
        row += 1
        for line in body.strip().split("\n"):
            ws.write(row, 0, line.strip(), body_f)
            n_lines = max(1, len(line) // 90 + 1)
            ws.set_row(row, 15 * n_lines)
            row += 1
        row += 1

    ws.print_area(0, 0, row, 0)
    ws.fit_to_pages(1, 0)
    wb.close()
    return output.getvalue()


def render_download(page_key):
    """Render the methodology download button for a given page."""
    configs = {
        "segregaciones": {
            "title": "Metodología — Generar Segregaciones",
            "filename": "Metodologia_Segregaciones.xlsx",
            "sections": [
                ("Objetivo", """
Generar las muestras de ubicaciones a auditar en un inventario interno,
divididas en tres segregaciones: Aleatorio, Control Diferenciado y Material Valioso.
                """),
                ("Datos de entrada", """
1. Excel de Stock: extracto completo del almacén con ubicaciones, materiales, lotes, cantidades.
2. Excel de Valores Unitarios: valor unitario por material/lote para calcular el valor total por ubicación.
3. Excel de Control Diferenciado: listado de materiales sujetos a control especial.
                """),
                ("Cruce de valores", """
Se cruza el stock con los valores unitarios usando la siguiente prioridad:
  - Primero: cruce exacto por Material + Lote
  - Segundo: cruce por Lote solamente
  - Tercero: cruce por Material solamente
El valor total de cada línea = Cantidad × Valor Unitario.
                """),
                ("Material Valioso — Lógica", """
1. Se agrupan las ubicaciones por valor total y se ordenan de mayor a menor.
2. Se acumulan ubicaciones hasta cubrir el % de valor configurado (por defecto 80%).
3. Se incluye la primera ubicación que supera el umbral para no quedarse corto.
4. Si se supera el máximo de ubicaciones, se muestra error con el % cubierto.
Asunción: las ubicaciones de mayor valor son las más críticas para auditar.
                """),
                ("Control Diferenciado — Lógica", """
1. Se identifican ubicaciones que contienen materiales del listado de control.
2. Se filtran solo las que tienen al menos N líneas (configurable).
3. Se excluyen ubicaciones ya asignadas a Material Valioso (sin duplicados).
4. Se seleccionan aleatoriamente N ubicaciones de las elegibles.
                """),
                ("Aleatorio — Lógica", """
1. Se seleccionan ubicaciones con al menos N líneas (configurable).
2. Se excluyen las ya asignadas a Valioso o Control.
3. Se seleccionan aleatoriamente N ubicaciones.
Por defecto se usa una semilla aleatoria (diferente cada vez).
Opcionalmente se puede fijar la semilla para reproducibilidad.
                """),
                ("No repetir ubicaciones (historial)", """
Opción activable que excluye ubicaciones auditadas en inventarios anteriores:
  - Material Valioso: salta a las siguientes ubicaciones por valor.
  - Control Diferenciado: selecciona de entre las no auditadas previamente.
  - Aleatorio: no se ve afectado (ya es aleatorio por naturaleza).
El historial se guarda al pulsar "Guardar en historial" y se puede limpiar
para reiniciar el ciclo cuando se han cubierto todas las ubicaciones.
                """),
                ("Parámetros configurables", """
- Nº de ubicaciones aleatorias (por defecto 10)
- Nº de ubicaciones de control (por defecto 10)
- Mínimo de líneas por ubicación para cada tipo (por defecto 4)
- % del valor a cubrir en Valioso (por defecto 80%)
- Máximo de ubicaciones en Valioso (por defecto 30)
                """),
                ("Formato de salida", """
Excel con una pestaña por segregación. Columnas:
Fecha, Ref. centro, Ref. Almacén, Ubicación, Ref. Material, Descripción,
Nº Lote, Valor unitario (si aplica), Valor total (si aplica), Nº Serie,
Stock, Cant. Física (a rellenar), Descuadre (fórmula), Unidad Base,
Stock OK, Stock Bloqueado, Tipo Bloqueo, Observaciones.
Control Diferenciado añade: Fallo en el proceso, Obs. Inventario, Obs. Proceso.
                """),
            ],
        },
        "reporte_consolidado": {
            "title": "Metodología — Reporte Consolidado de Auditoría",
            "filename": "Metodologia_Reporte_Consolidado.xlsx",
            "sections": [
                ("Objetivo", """
Generar el reporte de KPIs de la auditoría interna a partir del Excel
rellenado por los operarios en campo.
                """),
                ("Datos de entrada", """
1. Excel de auditoría rellenado: el Excel generado en Segregaciones, con las
   columnas "Cant. Física" y "Observaciones" rellenadas por los operarios.
2. Stock original (opcional): para calcular % de cobertura sobre el total.
                """),
                ("Clasificación de pestañas", """
Se detectan automáticamente las pestañas por nombre:
  - "control" o "diferenc" → Control Diferenciado
  - "valios" → Material Valioso
  - "aleatori" → Aleatorio
                """),
                ("Fiabilidad del inventario", """
Fórmula: Fiabilidad = 1 − (Lotes erróneos / Lotes auditados)
Un lote se considera erróneo cuando: Cant. Física ≠ Stock (descuadre ≠ 0).
Solo se cuentan lotes donde el operario ha rellenado la Cant. Física.
                """),
                ("Cumplimiento del proceso (solo Control Diferenciado)", """
Fórmula: Cumplimiento = 1 − (Fallos proceso / Lotes revisados)
Un fallo de proceso se cuenta cuando la columna "Fallo en el proceso"
tiene un valor distinto de vacío, "no" o "n".
Es independiente de la fiabilidad: mide si las pegatinas/etiquetas
están correctas, no si la mercancía cuadra.
                """),
                ("Cobertura vs Stock total", """
Si se sube el stock original, se calcula:
  - % Ubicaciones = Ubic. auditadas / Ubic. totales del almacén
  - % Lotes = Lotes auditados / Lotes totales
  - % Referencias = Refs. únicas auditadas / Refs. totales
                """),
                ("Validación", """
Se detectan dos tipos de incidencia:
1. Descuadre sin justificar: Cant. Física ≠ Stock pero sin observación.
2. Fallo en proceso sin justificar: fallo marcado pero sin Obs. Proceso.
Estas incidencias se listan para que se completen antes de cerrar la auditoría.
                """),
                ("Formato de salida", """
Excel con dos pestañas:
1. "KPIs Consolidado Aud.Interna": resumen con totales, desglose por sección,
   tabla lateral de cumplimiento de proceso, y resumen de Material Valioso.
2. "Validación" (si hay incidencias): listado de descuadres/fallos sin justificar.
                """),
            ],
        },
        "bip_recepciones": {
            "title": "Metodología — Análisis de Recepciones BIP",
            "filename": "Metodologia_BIP_Recepciones.xlsx",
            "sections": [
                ("Objetivo", """
Analizar las recepciones del SGA para obtener KPIs mensuales de calidad,
Pareto de incidencias por proveedor, y desglose mensual.
                """),
                ("Datos de entrada", """
1. Extracto SGA: Excel con recepciones (Fecha, Proveedor, Incidencia Sí/No,
   Estado incidencia, etc.)
2. Aliases proveedores (opcional): Excel con 3 hojas para normalizar nombres:
   - Aliases: variantes de nombre → nombre canónico
   - Force Split: textos que contienen varios proveedores → separar
   - Paquetería: proveedores que son empresas de paquetería
                """),
                ("Normalización de proveedores", """
1. Se normalizan los nombres: quitar acentos, sufijos legales (S.L., S.A., GmbH).
2. Se aplican aliases para unificar variantes del mismo proveedor.
3. Se aplica force-split para textos que contienen varios proveedores.
4. Se tokeniza por separadores (+, /, ,, ;, &, "y").
5. El "owner" de una incidencia es el primer proveedor del texto original.
                """),
                ("KPIs mensuales (ResultadoFinal)", """
Por cada mes se calcula:
  - Total Receipts: recepciones con Incidencia = Sí o No (se excluyen vacías)
  - Total Receipts with incidents: recepciones con Incidencia = Sí
  - Supplier (with incidents): incidencias donde el proveedor NO es Nordex
  - Nordex (with incidents): incidencias donde el proveedor ES Nordex
  - Resolved / Unresolved: según Estado Incidencia = "Solucionada" o "Abierta"
  - % with incidents about the total: incidencias / total recepciones
                """),
                ("Pareto de incidencias", """
1. Las incidencias se asignan al "owner" (primer proveedor del texto).
2. Se ordenan de mayor a menor incidencias.
3. Se calcula el % acumulado para identificar el 80/20.
4. Se marca si el proveedor es paquetería.
Asunción: el primer proveedor mencionado es el responsable principal.
                """),
                ("Desglose mensual (split)", """
Cada recepción se explode por los proveedores detectados (split).
Se cuentan recepciones por proveedor × mes.
Se añaden observaciones de con quién comparte recepciones.
                """),
                ("Ajuste split", """
Muestra la diferencia entre filas originales y filas tras el split,
para que el usuario pueda ajustar las cifras si lo necesita.
                """),
                ("Formato de salida", """
4 ficheros Excel (descargables por separado o en ZIP):
1. ResultadoFinal.xlsx — KPIs mensuales
2. ParetoIncidenciasProveedores.xlsx — Pareto con resumen
3. Recepciones_por_mes_y_proveedor_SPLIT.xlsx — Desglose mensual
4. Ajuste_split_resumen.xlsx — Diferencias por el split
                """),
            ],
        },
        "absentismo": {
            "title": "Metodología — Análisis de Absentismo",
            "filename": "Metodologia_Absentismo.xlsx",
            "sections": [
                ("Objetivo", """
Analizar el absentismo por centro de trabajo a partir de los cuadrantes
mensuales de horas, generando KPIs y un reporte consolidado multi-centro.
                """),
                ("Datos de entrada", """
Un Excel por centro de trabajo con el cuadrante mensual:
  - Filas: empleados
  - Columnas: días del mes (1-31)
  - Celdas: código o número de horas
                """),
                ("Códigos reconocidos", """
  - Número (cualquier valor numérico): día trabajado
  - V: Vacaciones
  - B: Baja (incapacidad temporal)
  - AP: Asuntos Propios
  - P: Permiso retribuido
  - E: Excedencia
  - Celda vacía: se marca como advertencia (dato faltante)
  - Cualquier otro código: se marca como desconocido
                """),
                ("Días laborables", """
Se calculan automáticamente según el calendario:
  - Se cuentan los días de lunes a viernes del mes.
  - Se restan los festivos nacionales de España:
    Año Nuevo (1 ene), Epifanía (6 ene), Viernes Santo (variable),
    Día del Trabajo (1 may), Asunción (15 ago), Hispanidad (12 oct),
    Todos los Santos (1 nov), Constitución (6 dic),
    Inmaculada (8 dic), Navidad (25 dic).
Nota: los festivos autonómicos/locales no se incluyen automáticamente.
                """),
                ("Cálculo del absentismo", """
Días teóricos = Plantilla × Días laborables del mes
Horas teóricas = Días teóricos × 8

% Absentismo CON vacaciones = (V + B + AP + P + E) / Días teóricos × 100
% Absentismo SIN vacaciones = (B + AP + P + E) / Días teóricos × 100

La versión "sin vacaciones" permite ver el impacto real de ausencias
no planificadas (bajas, permisos, etc.) sin el efecto de las vacaciones.
                """),
                ("Asunciones", """
  - Cada día con un número (sin importar las horas) se cuenta como 1 día trabajado.
  - Se asumen jornadas de 8 horas para el cálculo de horas.
  - Un empleado aparece en el cuadrante = cuenta en la plantilla.
  - No se incluyen festivos autonómicos/locales (solo nacionales).
                """),
                ("Validaciones", """
  - Celdas vacías: se avisa con detalle de empleado y día afectado.
  - Códigos desconocidos: se listan para que el usuario los clasifique.
  - Si un empleado no tiene ningún dato registrado, no se incluye en el análisis.
                """),
                ("Formato de salida", """
Excel con:
1. Hoja "Resumen": tabla comparativa de todos los centros con métricas
   (plantilla, días, horas, desglose ausencias, % absentismo).
2. Una hoja por centro: detalle empleado a empleado.
                """),
            ],
        },
    }

    if page_key not in configs:
        return

    cfg = configs[page_key]

    with st.expander("Metodología / Lógica del reporte"):
        for heading, body in cfg["sections"]:
            st.markdown(f"**{heading}**")
            st.markdown(body.strip())
            st.markdown("---")
        st.download_button(
            "Descargar metodología (Excel)",
            _build_excel(cfg["title"], cfg["sections"]),
            file_name=cfg["filename"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
