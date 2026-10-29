"""
actualizar_cuentas.py — Boletas Enel (PDF) -> planilla Excel con metricas

Autor: Rodrigo Meza López
Contacto: rodrigo.meza.lopez@gmail.com
Versión: 0.1
Última actualización: 2026-10-29
Licencia: MIT (ver LICENSE) — uso y modificación libres citando al autor.
Repositorio: https://github.com/meriadox/enel-bill-tracker

Genera, DESDE CERO, una planilla nueva y limpia de consumo/costos de
electricidad a partir de las boletas PDF (Enel) guardadas en esta misma
carpeta. No modifica ningun .xlsx existente en la carpeta: cada corrida crea
(o reemplaza, con respaldo, si ya existia con el mismo nombre) un archivo
propio, nombrado segun el periodo que cubren las boletas encontradas.

Uso (ver README.md del repositorio para instalación y dependencias):
    python actualizar_cuentas.py            # genera la planilla (con respaldo si ya existia)
    python actualizar_cuentas.py --dry-run   # solo muestra que haria, no escribe nada
    python actualizar_cuentas.py --version   # muestra version y datos de contacto

No hay rutas ni nombres de archivo hardcodeados: el script toma como base la
carpeta donde el mismo se encuentra, procesa todos los .pdf/.PDF presentes,
y calcula el nombre de salida ("Resumen consumo electrico - AAAA-AAAA.xlsx")
a partir del primer y el ultimo año con boletas leidas. Cualquier otro .xlsx
que ya exista en la carpeta (por ejemplo una planilla historica con columnas
de calculo manuales o imagenes pegadas) se deja intacto: el script nunca lo
lee ni lo escribe.

Que hace:
  1. Extrae de cada boleta: N de boleta, periodo de lectura (inicio/fin),
     consumo (kWh), total a pagar, fecha de emision, fecha de vencimiento,
     tipo de tarifa, y el desglose de cargos detallado (ver CARGOS_DETALLADOS).
     Para boletas cuyo PDF tiene texto corrupto (fuente sin tabla Unicode),
     intenta un segundo metodo de lectura (pdftotext -raw) antes de darse
     por vencido.
  2. Boletas que no se pueden leer del todo (falta el N de boleta, el
     periodo o el total a pagar) NO entran a la planilla: se listan en
     "boletas_no_procesadas.txt" para que el usuario las revise/complete a
     mano.
  3. Escribe la hoja "Resumen Boletas" (una fila por boleta, orden mas
     reciente arriba) con las columnas base + info adicional + cargos
     detallados. Si un cargo no viene en una boleta puntual, la celda queda
     en "s/i" (nunca se adivina ni se calcula un numero). No incluye
     formulas ni columnas de calculo manual: esos resumenes viven en la
     hoja "Metricas".
  4. Genera la hoja "Metricas": resumen general (consumo y pago total,
     promedio mensual, maximos), y el detalle anual / trimestral /
     semestral / mensual, con graficos de barras nativos de Excel (con
     ejes rotulados).

Manejo de periodos de boleta que se solapan (ver tambien la nota dentro de
la propia hoja "Metricas"): el consumo (kWh) se prorratea dia por dia dentro
del periodo de lectura de cada boleta; si dos boletas reclaman el mismo dia
calendario (p.ej. una boleta "sin lectura" seguida de la boleta correctiva
que re-factura el mismo tramo), se usa el dato de la boleta emitida mas
tarde para esos dias, evitando contar el consumo dos veces. El gasto ($) se
asigna al año/mes/trimestre/semestre de la FECHA DE EMISION de cada boleta,
que nunca se solapa (cada boleta se paga una sola vez).

Limitaciones conocidas (quedan fuera a proposito, para no insertar datos
erroneos en la planilla):
  - El IVA casi nunca queda en la misma linea que su monto en el texto del
    PDF (cambia de layout entre años); la columna existe pero en la mayoria
    de los casos quedara en "s/i".
  - Los descuentos/compensaciones promocionales (PVC, SERNAC, bonos por
    temporales, etc.) se redactan como oraciones libres con montos
    incrustados en el texto, de formato muy variable año a año: no se
    extraen como columna para no arriesgar leer mal un numero.
  - Dos PDF de 2021 tienen el texto corrupto (fuente sin tabla Unicode); el
    N de boleta se puede recuperar con el metodo de respaldo, pero el resto
    de los campos no, por lo que sin lectura previa no entrarian a la
    planilla. (En la planilla historica del usuario ya estaban registrados
    a mano.)
"""

from __future__ import annotations

import argparse
import glob
import re
import shutil
import statistics
import subprocess
import sys
from collections import OrderedDict
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font

__author__ = "Rodrigo Meza López"
__email__ = "rodrigo.meza.lopez@gmail.com"
__version__ = "0.1"
__last_updated__ = "2026-10-29"
__license__ = "MIT"

BASE_DIR = Path(__file__).resolve().parent

MESES = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "oct": 10, "nov": 11, "dic": 12,
}
MESES_NOMBRE = ["", "Ene", "Feb", "Mar", "Abr", "May", "Jun",
                "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]

CAMPOS_CRITICOS = ("nboleta", "inicio", "final", "pagado")

# columnas base (siempre, en este orden) de la hoja "Resumen Boletas"
COLUMNAS_BASE = [
    ("ano", "Año"),
    ("inicio", "Inicio"),
    ("final", "Final"),
    ("consumokwh", "Consumo (kWh)"),
    ("nboleta", "N° Boleta"),
    ("pagado", "Pagado ($)"),
    ("fechaemision", "Fecha Emisión"),
    ("fechavencimiento", "Fecha Vencimiento"),
    ("tipotarifa", "Tipo Tarifa"),
]
CLAVES_INFO_EXTRA = {"fechaemision", "fechavencimiento", "tipotarifa"}

# Cargos detallados: clave interna -> (titulo de columna, regex que captura el monto).
# Cada regex exige que el "$" y el numero esten en la MISMA linea que la
# etiqueta (igual que ya funciona para "Total a pagar"); si no se encuentra
# en ese formato exacto, la celda queda en "s/i" en vez de arriesgar un
# numero mal leido.
CARGOS_DETALLADOS = OrderedDict([
    ("servicio_electrico", ("Servicio Eléctrico ($)", r"^\s*Servicio El[eé]ctrico\s+\$\s*(-?[\d.]+)")),
    ("electricidad_consumida_monto", ("Electricidad Consumida ($)", r"^\s*Electricidad Consumida(?:\s*\([^)]*\))?\s+\$\s*(-?[\d.]+)")),
    ("administracion_servicio", ("Administración del Servicio ($)", r"^\s*Administraci[oó]n del servicio\s+\$\s*(-?[\d.]+)")),
    ("transporte_electricidad", ("Transporte de Electricidad ($)", r"^\s*Transporte de electricidad\s+\$\s*(-?[\d.]+)")),
    ("cargo_fondo_estabilizacion", ("Cargo Fondo Estabilización Ley 21.472 ($)", r"^\s*Cargo Fondo de Estabilizaci[oó]n Ley 21\.472\s+\$\s*(-?[\d.]+)")),
    ("arriendo_medidor", ("Arriendo Medidor ($)", r"^\s*Arriendo Medidor\s+\$\s*(-?[\d.]+)")),
    ("otros_cargos", ("Otros Cargos ($)", r"^\s*Otros [Cc]argos?(?:\s*\([^)]*\))?\*?\s+\$\s*(-?[\d.]+)")),
    ("otros_cargos_abonos", ("Otros Cargos/Abonos ($)", r"^\s*Otros [Cc]argos?/abonos\s+\$\s*(-?[\d.]+)")),
    ("saldo_anterior", ("Saldo Anterior ($)", r"^\s*Saldo [Aa]nterior\s+\$\s*(-?[\d.]+)")),
    ("saldo_anterior_energia", ("Saldo Anterior Energía ($)", r"^\s*Saldo Anterior Energ[ií]a\s+\$\s*(-?[\d.]+)")),
    ("saldo_anterior_otros", ("Saldo Anterior Otros Cargos ($)", r"^\s*Saldo Anterior Otros Cargos\s+\$\s*(-?[\d.]+)")),
    ("monto_afecto_impuesto", ("Monto Afecto a Impuesto ($)", r"^\s*Monto afecto a impuesto\s+\$\s*(-?[\d.]+)")),
    ("iva", ("IVA ($)", r"^\s*IVA\s+\$\s*(-?[\d.]+)")),
    ("monto_exento_impuesto", ("Monto Exento de Impuesto ($)", r"^\s*Monto exento \w+ impuesto\s+\$\s*(-?[\d.]+)")),
    ("total_boleta", ("Total Boleta ($)", r"^\s*Total boleta\s+\$\s*(-?[\d.]+)")),
    ("devolucion_facturacion_provisoria", ("Devolución Facturación Provisoria ($)", r"^\s*Devoluci[oó]n por facturaci[oó]n provisoria\s+\$\s*(-?[\d.]+)")),
    ("descuento_interrupciones", ("Descuento Interrupciones Servicio ($)", r"^\s*Descuento por interrupciones del servicio\s+\$\s*(-?[\d.]+)")),
    ("intereses_mora", ("Intereses de Mora ($)", r"^\s*Intereses de Mora\s+\$\s*(-?[\d.]+)")),
    ("facturacion_provisoria_energia", ("Facturación Provisoria Energía ($)", r"^\s*Facturaci[oó]n provisoria de energ[ií]a(?:\s*\([^)]*\))?\s+\$\s*(-?[\d.]+)")),
    ("facturacion_provisoria_transporte", ("Facturación Provisoria Transporte ($)", r"^\s*Facturaci[oó]n provisoria de transporte de electricidad(?:\s*\([^)]*\))?\s+\$\s*(-?[\d.]+)")),
    ("credito_servicios", ("Crédito Servicios Eléctricos y Otros ($)", r"Cr[eé]dito servicios El[eé]ctricos y Otros\s*\$\s*(-?[\d.]+)")),
])


def parse_fecha_texto(s: str) -> date | None:
    """'23 Sep 2025' -> date(2025, 9, 23)"""
    m = re.match(r"(\d{1,2})\s+([A-Za-z]{3,4})\.?\s+(\d{4})", s)
    if not m:
        return None
    d, mes_txt, y = m.groups()
    mes = MESES.get(mes_txt.lower()[:3])
    if not mes:
        return None
    return date(int(y), mes, int(d))


def extraer_texto(path: Path, modo: str) -> str:
    r = subprocess.run(
        ["pdftotext", modo, "-enc", "UTF-8", str(path), "-"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
    )
    return r.stdout


def extraer_campos_criticos(text: str) -> dict:
    rec: dict = {}

    m = re.search(r"N[°ºª]\s*(\d{6,10})", text)
    rec["nboleta"] = int(m.group(1)) if m else None

    m = re.search(r"Fecha de emisi[oó]n:?\s*(\d{1,2}\s+[A-Za-z]{3,4}\.?\s+\d{4})", text, re.IGNORECASE)
    rec["fechaemision"] = parse_fecha_texto(m.group(1)) if m else None

    m = re.search(r"Fecha de vencimiento\s+(\d{1,2}\s+[A-Za-z]{3,4}\.?\s+\d{4})", text, re.IGNORECASE)
    rec["fechavencimiento"] = parse_fecha_texto(m.group(1)) if m else None

    m = re.search(
        r"Per[ií]odo de lectura:?\s*(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})",
        text, re.IGNORECASE,
    )
    if m:
        rec["inicio"] = datetime.strptime(m.group(1), "%d/%m/%Y").date()
        rec["final"] = datetime.strptime(m.group(2), "%d/%m/%Y").date()
    else:
        # boletas "sin lectura" (facturacion estimada) no traen Periodo de
        # lectura numerico, solo el rango en texto de "Monto del periodo"
        m = re.search(
            r"Monto del periodo\s+(\d{1,2}\s+[A-Za-z]{3,4}\.?\s+\d{4})\s*-\s*(\d{1,2}\s+[A-Za-z]{3,4}\.?\s+\d{4})",
            text, re.IGNORECASE,
        )
        if m:
            rec["inicio"] = parse_fecha_texto(m.group(1))
            rec["final"] = parse_fecha_texto(m.group(2))
        else:
            rec["inicio"] = rec["final"] = None

    m = re.search(r"Consumo total del (?:periodo|mes)\s*=\s*(-?\d+)\s*kWh", text, re.IGNORECASE)
    if m:
        rec["consumokwh"] = int(m.group(1))
    elif re.search(r"No se pudo tomar lectura|Sin Lectura", text, re.IGNORECASE):
        rec["consumokwh"] = "Sin lectura"
    else:
        rec["consumokwh"] = None

    m = re.search(r"Total a pagar:?\s*\$\s*(-?[\d\.]+)", text, re.IGNORECASE)
    if m:
        rec["pagado"] = int(m.group(1).replace(".", ""))
    elif re.search(r"Total a pagar:?\s*En Cr[eé]dito", text, re.IGNORECASE):
        rec["pagado"] = 0
    else:
        rec["pagado"] = None

    m = re.search(r"Tipo de tarifa contratada:\s*([A-Za-z0-9\-]+)", text, re.IGNORECASE)
    rec["tipotarifa"] = m.group(1) if m else None

    if rec["inicio"] is not None and rec["final"] is not None:
        if rec["inicio"].year == rec["final"].year:
            rec["ano"] = rec["final"].year
        else:
            rec["ano"] = f"{rec['inicio'].year}-{rec['final'].year}"
    else:
        rec["ano"] = None

    return rec


def extraer_cargos(text: str) -> dict:
    cargos = {}
    for clave, (_, patron) in CARGOS_DETALLADOS.items():
        m = re.search(patron, text, re.IGNORECASE | re.MULTILINE)
        cargos[clave] = int(m.group(1).replace(".", "")) if m else None
    return cargos


def parse_pdf(path: Path) -> dict:
    """Extrae los campos de una boleta Enel. Intenta -layout y, si faltan
    campos criticos, reintenta con -raw (algunos PDF antiguos tienen una
    fuente corrupta que -layout no logra ordenar, pero -raw a veces si
    conserva legible el encabezado)."""
    text_layout = extraer_texto(path, "-layout")
    rec = extraer_campos_criticos(text_layout)
    rec["archivo"] = path.name
    rec["metodo"] = "layout"

    if any(rec.get(c) is None for c in CAMPOS_CRITICOS):
        text_raw = extraer_texto(path, "-raw")
        rec_raw = extraer_campos_criticos(text_raw)
        recuperado_de_raw = False
        for campo in ("nboleta", "fechaemision", "fechavencimiento", "inicio",
                      "final", "consumokwh", "pagado", "tipotarifa"):
            if rec.get(campo) is None and rec_raw.get(campo) is not None:
                rec[campo] = rec_raw[campo]
                recuperado_de_raw = True
        if recuperado_de_raw:
            rec["metodo"] = "layout+raw"
            if rec["inicio"] is not None and rec["final"] is not None:
                if rec["inicio"].year == rec["final"].year:
                    rec["ano"] = rec["final"].year
                else:
                    rec["ano"] = f"{rec['inicio'].year}-{rec['final'].year}"

    rec["cargos"] = extraer_cargos(text_layout)
    return rec


def buscar_pdfs(base_dir: Path) -> list[Path]:
    vistos = {}
    for patron in ("*.pdf", "*.PDF"):
        for p in base_dir.glob(patron):
            vistos[p.name.lower()] = p
    return sorted(vistos.values(), key=lambda p: p.name)


def escribir_reporte_no_procesadas(base_dir: Path, items: list[dict]) -> Path | None:
    ruta = base_dir / "boletas_no_procesadas.txt"
    if not items:
        if ruta.exists():
            ruta.unlink()
        return None
    lineas = [
        "Boletas que NO se pudieron incluir en la planilla generada.",
        "Revisar/completar a mano estos PDF.",
        f"Generado: {datetime.now():%Y-%m-%d %H:%M}",
        "-" * 70,
    ]
    for it in items:
        lineas.append(f"\nArchivo: {it['archivo']}")
        lineas.append(f"  Motivo: {it['motivo']}")
        for campo in ("nboleta", "inicio", "final", "pagado"):
            valor = it.get(campo)
            lineas.append(f"  {campo}: {valor if valor is not None else '(no se pudo leer)'}")
    ruta.write_text("\n".join(lineas), encoding="utf-8")
    return ruta


LEEME_TEXTO = """\
# LEEME — Supuestos usados por actualizar_cuentas.py

Este archivo documenta los supuestos y reglas de calculo que usa el script
para transformar las boletas PDF en la planilla de consumo electrico, para
que cualquiera (incluyendo quien no lea el codigo) pueda juzgar si los
numeros son confiables para su uso.

## 1. Que boletas entran a la planilla

Cada PDF se procesa primero con `pdftotext -layout`. Si no logra leer el N°
de boleta, el periodo de lectura o el total a pagar, se reintenta con
`pdftotext -raw` (algunos PDF antiguos usan una fuente sin tabla Unicode que
-layout no logra ordenar, pero -raw a veces conserva legible el
encabezado). Una boleta se incluye en la planilla SOLO si, entre ambos
metodos, se logran leer sus 4 campos criticos: N° de boleta, inicio y fin
del periodo de lectura, y total a pagar. Si falta alguno, la boleta queda
afuera y se lista en `boletas_no_procesadas.txt` junto con lo que si se
pudo rescatar, para revision manual.

## 2. Cargos detallados ("s/i")

El desglose de cargos (servicio electrico, administracion, arriendo de
medidor, IVA, etc.) se busca linea por linea en el texto del PDF, exigiendo
que la etiqueta y el monto ($) esten en la MISMA linea. Si una boleta no
trae ese cargo en ese formato exacto, la celda queda en "s/i" — nunca se
calcula, interpola o adivina un numero. Por eso:
  - La columna IVA queda en "s/i" en la gran mayoria de las boletas: el
    monto de IVA casi nunca aparece pegado a su etiqueta en el texto
    (depende del layout, que cambio varias veces entre 2021 y 2026).
  - Los descuentos/compensaciones promocionales (PVC, SERNAC, bonos por
    temporales, etc.) NO se extraen como columna: se redactan como
    oraciones de texto libre con montos incrustados, de formato
    demasiado variable para leer con confianza.

## 3. Como se evita contar el consumo dos veces (boletas que se solapan)

El consumo (kWh) de cada boleta se reparte en una tasa diaria (kWh del
periodo / dias del periodo) y se asigna a cada dia calendario que cubre esa
boleta. Cuando dos boletas reclaman el mismo dia (tipicamente una boleta
"sin lectura"/estimada seguida, meses despues, de la boleta correctiva que
re-factura el mismo tramo con la lectura real), se recorren las boletas en
orden de FECHA DE EMISION ascendente y cada dia que cubre una boleta
posterior sobreescribe el valor que le hubiera puesto una boleta anterior.
Como la correccion siempre se emite despues que la estimacion, esto hace
que el dato mas confiable (o el mas reciente disponible) sea el que
prevalece, sin duplicar consumo.

Una consecuencia de este metodo: si una boleta "sin lectura" NO fue
corregida despues por otra boleta que cubra esos mismos dias, esos dias
quedan simplemente SIN dato de consumo (no se inventa un valor). Por eso
la suma de dias con dato de algunos años es menor a 365/366.

Las boletas "sin lectura" (sin kWh numerico) nunca aportan una tasa: no
pueden pisar un dato real ya existente ni rellenar un hueco por si solas.

## 4. Como se asigna el gasto ($) en el tiempo

A diferencia del consumo, el gasto de cada boleta se asigna integro al
año/mes/trimestre/semestre de su FECHA DE EMISION (no se prorratea por
dias). Esto es deliberado: cada boleta se paga una sola vez, en una fecha
real, sin importar que tramo de lectura facture — por lo que esta
asignacion nunca se solapa ni necesita ningun metodo de resolucion de
conflictos.

## 5. "Años completos" en los promedios anuales y el promedio mensual estacional

El primer y el ultimo año del historico casi nunca estan completos (los
datos empiezan o terminan a mitad de año, segun la fecha de la boleta mas
antigua/reciente disponible). Un promedio simple sobre todos los años
quedaria distorsionado por esos dos años parciales. Por eso:
  - "Consumo/Pago promedio anual" (en Resumen general) se calcula
    excluyendo el primer y el ultimo año del rango de datos.
  - El "Promedio mensual historico (Ene-Dic)" tambien promedia cada mes
    del calendario SOLO a traves de esos mismos años completos (y, dentro
    de ellos, solo usa los años que efectivamente tienen dato para ese mes
    puntual — ver seccion 3 sobre huecos por boletas sin lectura no
    corregidas).
  - Si hay menos de 3 años de datos en total, no hay forma de excluir un
    primer/ultimo año sin quedarse sin datos: esos promedios quedan en
    "s/i" con una nota aclaratoria.

Las tablas "Consumo mensual", "trimestral" y "semestral" (no el promedio
estacional) SI incluyen los años parciales tal cual, para no ocultar
informacion real — la exclusion solo aplica a los promedios resumen.

## 6. Archivo de salida

El nombre del archivo generado refleja el rango de años de las boletas
incluidas (ej. "Resumen consumo electrico - 2021-2026.xlsx") y se
recalcula en cada corrida. El script SIEMPRE reconstruye la planilla
completa desde los PDF (no la actualiza de forma incremental) y NUNCA lee
ni modifica otro archivo .xlsx que ya exista en la carpeta. Si ya existe un
archivo con el mismo nombre de salida, se respalda con marca de tiempo
antes de sobreescribirlo.

## 7. Zona horaria / redondeo

Las fechas se interpretan tal como aparecen impresas en la boleta (hora de
Chile, sin conversion). Los montos en $ son pesos chilenos sin decimales
(como los imprime Enel). Los kWh prorrateados por dia se redondean solo al
mostrarlos en las tablas; los calculos internos usan el valor con
decimales para no acumular error de redondeo.
"""


def escribir_leeme(base_dir: Path) -> Path:
    ruta = base_dir / "LEEME.md"
    ruta.write_text(LEEME_TEXTO, encoding="utf-8")
    return ruta


def nombre_archivo_salida(registros: list[dict]) -> str:
    anos = []
    for r in registros:
        for campo in ("inicio", "final"):
            if r.get(campo) is not None:
                anos.append(r[campo].year)
    if not anos:
        return "Resumen consumo electrico.xlsx"
    amin, amax = min(anos), max(anos)
    if amin == amax:
        return f"Resumen consumo electrico - {amin}.xlsx"
    return f"Resumen consumo electrico - {amin}-{amax}.xlsx"


# ---------------------------------------------------------------------------
# Hoja "Resumen Boletas"
# ---------------------------------------------------------------------------

def construir_hoja_resumen(wb, registros: list[dict]) -> None:
    ws = wb.active
    ws.title = "Resumen Boletas"

    columnas = list(COLUMNAS_BASE) + [(clave, titulo) for clave, (titulo, _) in CARGOS_DETALLADOS.items()]
    for col, (_, titulo) in enumerate(columnas, start=1):
        celda = ws.cell(row=1, column=col, value=titulo)
        celda.font = Font(bold=True)

    registros_ordenados = sorted(registros, key=lambda r: r["inicio"], reverse=True)

    for fila, rec in enumerate(registros_ordenados, start=2):
        for col, (clave, _) in enumerate(columnas, start=1):
            if clave in CARGOS_DETALLADOS:
                valor = rec.get("cargos", {}).get(clave)
            else:
                valor = rec.get(clave)
            celda = ws.cell(row=fila, column=col, value=valor if valor is not None else "s/i")
            if clave in ("inicio", "final", "fechaemision", "fechavencimiento") and isinstance(valor, date):
                celda.number_format = "d-mmm-yyyy"
            elif clave == "pagado" or clave in CARGOS_DETALLADOS:
                if isinstance(valor, (int, float)):
                    celda.number_format = '_-"$"\\ * #,##0_-;\\-"$"\\ * #,##0_-;_-"$"\\ * "-"_-;_-@_-'

    anchos = {"ano": 10, "inicio": 12, "final": 12, "consumokwh": 13, "nboleta": 12,
              "pagado": 12, "fechaemision": 13, "fechavencimiento": 15, "tipotarifa": 10}
    for i, (clave, _) in enumerate(columnas, start=1):
        letra = ws.cell(row=1, column=i).column_letter
        ws.column_dimensions[letra].width = anchos.get(clave, 24)

    ws.freeze_panes = "A2"


# ---------------------------------------------------------------------------
# Metricas
# ---------------------------------------------------------------------------

def construir_dias_consumo(registros: list[dict]) -> dict[date, float]:
    """
    Reparte el consumo (kWh) de cada boleta en una tasa diaria (kWh/dia) y la
    asigna a cada dia calendario que cubre esa boleta.

    Varias boletas pueden cubrir el MISMO dia calendario (boleta "sin
    lectura"/provisoria seguida de la boleta correctiva que re-factura el
    mismo tramo con la lectura real). Para no contar un mismo dia dos veces:
    se recorren las boletas en orden de FECHA DE EMISION ascendente, y cada
    dia que cubre una boleta posterior SOBRESCRIBE el valor que le hubiera
    puesto una boleta anterior. Como la boleta correctiva siempre se emite
    despues que la provisoria, esto hace que la lectura real (o la ultima
    version disponible) sea la que prevalece para los dias en comun.
    Las boletas "Sin lectura" (sin kWh numerico) no aportan tasa: no pueden
    pisar un dato real ya existente ni inventar uno donde no lo hay.
    """
    con_fecha = [r for r in registros if r.get("inicio") and r.get("final")]
    con_fecha.sort(key=lambda r: r.get("fechaemision") or r["inicio"])

    dias: dict[date, float] = {}
    for r in con_fecha:
        kwh = r.get("consumokwh")
        if not isinstance(kwh, (int, float)):
            continue
        inicio, final = r["inicio"], r["final"]
        n_dias = (final - inicio).days + 1
        if n_dias <= 0:
            continue
        tasa = kwh / n_dias
        d = inicio
        while d <= final:
            dias[d] = tasa
            d += timedelta(days=1)
    return dias


def _formatear_tabla(ws, titulo, encabezados, filas, formatos=None):
    ws.append([titulo])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True, size=12)
    ws.append(encabezados)
    for c in range(1, len(encabezados) + 1):
        ws.cell(row=ws.max_row, column=c).font = Font(bold=True)
    fila_inicio = ws.max_row + 1
    for fila in filas:
        ws.append(fila)
    fila_fin = ws.max_row
    ws.append([])
    return fila_inicio, fila_fin


def _ajustar_ejes(chart):
    """Evita que el titulo del eje se superponga con sus propios valores:
    en vez de un titulo de eje aparte (que Excel coloca pegado a las
    etiquetas y se solapa con ellas, sobre todo en el eje Y angosto o con
    categorias rotadas en el eje X), la unidad va en el titulo del grafico."""
    chart.x_axis.title = None
    chart.y_axis.title = None
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.x_axis.tickLblPos = "nextTo"
    chart.y_axis.tickLblPos = "nextTo"
    chart.x_axis.majorTickMark = "out"
    chart.y_axis.majorTickMark = "out"


def construir_metricas(wb, registros: list[dict]) -> dict | None:
    dias_kwh = construir_dias_consumo(registros)
    if "Metricas" in wb.sheetnames:
        del wb["Metricas"]
    ws = wb.create_sheet("Metricas")
    ws.column_dimensions["A"].width = 26
    for col in "BCDEF":
        ws.column_dimensions[col].width = 16

    if not dias_kwh:
        ws.append(["No hay suficientes boletas con periodo/consumo numerico para calcular metricas."])
        return None

    # --- gasto por boleta, agrupado por año/mes de emision ---
    gasto_anual: dict[int, int] = {}
    gasto_mensual: dict[tuple[int, int], int] = {}
    boleta_pago_max = max(
        (r for r in registros if isinstance(r.get("pagado"), (int, float))),
        key=lambda r: r["pagado"], default=None,
    )
    for r in registros:
        pagado = r.get("pagado")
        fecha = r.get("fechaemision") or r.get("inicio")
        if pagado is None or fecha is None:
            continue
        gasto_anual[fecha.year] = gasto_anual.get(fecha.year, 0) + pagado
        clave_mes = (fecha.year, fecha.month)
        gasto_mensual[clave_mes] = gasto_mensual.get(clave_mes, 0) + pagado

    anos = sorted({d.year for d in dias_kwh} | set(gasto_anual))

    consumo_anual: dict[int, float] = {a: 0.0 for a in anos}
    consumo_mensual: dict[tuple[int, int], float] = {}
    consumo_trim: dict[tuple[int, int], float] = {}
    consumo_sem: dict[tuple[int, int], float] = {}
    for d, tasa in dias_kwh.items():
        consumo_anual[d.year] += tasa
        consumo_mensual[(d.year, d.month)] = consumo_mensual.get((d.year, d.month), 0.0) + tasa
        trim = (d.year, (d.month - 1) // 3 + 1)
        consumo_trim[trim] = consumo_trim.get(trim, 0.0) + tasa
        sem = (d.year, 1 if d.month <= 6 else 2)
        consumo_sem[sem] = consumo_sem.get(sem, 0.0) + tasa

    boleta_kwh_max = max(
        (r for r in registros if isinstance(r.get("consumokwh"), (int, float))),
        key=lambda r: r["consumokwh"], default=None,
    )

    fecha_min = min(r["inicio"] for r in registros if r.get("inicio"))
    fecha_max = max(r["final"] for r in registros if r.get("final"))

    # El primer y el ultimo año del historico casi nunca estan completos
    # (los datos empiezan/terminan a mitad de año), asi que distorsionan un
    # promedio simple. Para "promedio anual" se usan solo los años
    # intermedios (ni el primero ni el ultimo del rango).
    if len(anos) >= 3:
        anos_completos = anos[1:-1]
        etiqueta_anos_completos = f"años completos {anos_completos[0]}–{anos_completos[-1]}"
    else:
        anos_completos = []
        etiqueta_anos_completos = "datos insuficientes para excluir años parciales"

    if anos_completos:
        promedio_anual_kwh = round(statistics.mean(consumo_anual[a] for a in anos_completos))
        promedio_anual_gasto = round(statistics.mean(gasto_anual.get(a, 0) for a in anos_completos))
    else:
        promedio_anual_kwh = promedio_anual_gasto = "s/i"

    # ---------------- Resumen general ----------------
    ws.append(["Resumen general"])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True, size=13)
    filas_resumen = [
        ["Período analizado", f"{fecha_min:%d-%m-%Y} a {fecha_max:%d-%m-%Y}"],
        ["N° de boletas incluidas", len(registros)],
        ["Consumo total histórico (kWh)", round(sum(consumo_anual.values()))],
        [f"Consumo promedio anual ({etiqueta_anos_completos}) (kWh)", promedio_anual_kwh],
        ["Consumo mensual promedio (kWh)", round(statistics.mean(consumo_mensual.values()), 1)],
        ["Consumo máximo en una boleta (kWh)",
         f"{boleta_kwh_max['consumokwh']} kWh  ({boleta_kwh_max['inicio']:%d-%m-%Y} a {boleta_kwh_max['final']:%d-%m-%Y}, boleta {boleta_kwh_max['nboleta']})"
         if boleta_kwh_max else "s/i"],
        ["Pago total histórico ($)", sum(gasto_anual.values())],
        [f"Pago promedio anual ({etiqueta_anos_completos}) ($)", promedio_anual_gasto],
        ["Pago mensual promedio ($)", round(statistics.mean(gasto_mensual.values()))],
        ["Pago máximo en una boleta ($)",
         (f"${boleta_pago_max['pagado']:,}".replace(",", ".") +
          f"  ({boleta_pago_max['inicio']:%d-%m-%Y} a {boleta_pago_max['final']:%d-%m-%Y}, boleta {boleta_pago_max['nboleta']})")
         if boleta_pago_max else "s/i"],
    ]
    for fila in filas_resumen:
        ws.append(fila)
    ws.append([])

    # ---------------- tabla anual ----------------
    filas_anual = []
    anterior_kwh = anterior_gasto = None
    for a in anos:
        kwh = round(consumo_anual.get(a, 0.0))
        gasto = gasto_anual.get(a, 0)
        costo_kwh = round(gasto / kwh, 1) if kwh else "s/i"
        var_kwh = round(100 * (kwh - anterior_kwh) / anterior_kwh, 1) if anterior_kwh else "s/i"
        var_gasto = round(100 * (gasto - anterior_gasto) / anterior_gasto, 1) if anterior_gasto else "s/i"
        filas_anual.append([a, kwh, gasto, costo_kwh, var_kwh, var_gasto])
        anterior_kwh, anterior_gasto = kwh, gasto
    fila_inicio_anual, fila_fin_anual = _formatear_tabla(
        ws, "Métricas anuales",
        ["Año", "Consumo (kWh)", "Gasto ($)", "Costo promedio ($/kWh)",
         "Var. consumo vs año anterior (%)", "Var. gasto vs año anterior (%)"],
        filas_anual,
    )

    # ---------------- tabla mensual ----------------
    filas_mensual = []
    for clave in sorted(consumo_mensual):
        a, m = clave
        filas_mensual.append([f"{a}-{MESES_NOMBRE[m]}", round(consumo_mensual[clave]),
                               gasto_mensual.get(clave, "s/i")])
    fila_inicio_mes, fila_fin_mes = _formatear_tabla(
        ws, "Consumo y gasto mensual",
        ["Año-Mes", "Consumo (kWh)", "Gasto ($)"],
        filas_mensual,
    )

    # ---------------- promedio mensual historico (estacional, Ene-Dic) ----------------
    # Un solo valor por mes del calendario (12 filas), promediando ese mes a
    # traves de los "años completos" (se excluyen el primer y el ultimo año
    # del historico por las mismas razones que en "Resumen general"). Si a
    # un año completo le falta un mes puntual (p.ej. por una boleta "sin
    # lectura" sin correccion que lo cubra), ese año simplemente no aporta
    # dato a ese mes -- el promedio se calcula solo con los años que si
    # tienen informacion para ese mes.
    filas_estacional = []
    for m in range(1, 13):
        valores = [consumo_mensual[(a, m)] for a in anos_completos if (a, m) in consumo_mensual]
        promedio = round(statistics.mean(valores)) if valores else "s/i"
        filas_estacional.append([MESES_NOMBRE[m], promedio])
    fila_inicio_estacional, fila_fin_estacional = _formatear_tabla(
        ws, f"Promedio mensual histórico ({etiqueta_anos_completos})",
        ["Mes", "Consumo promedio (kWh)"], filas_estacional,
    )

    # ---------------- tabla trimestral ----------------
    filas_trim = [[f"{t[0]}-T{t[1]}", round(consumo_trim[t])] for t in sorted(consumo_trim)]
    fila_inicio_trim, fila_fin_trim = _formatear_tabla(
        ws, "Consumo trimestral", ["Año-Trimestre", "Consumo (kWh)"], filas_trim,
    )

    # ---------------- tabla semestral ----------------
    filas_sem = [[f"{s[0]}-S{s[1]}", round(consumo_sem[s])] for s in sorted(consumo_sem)]
    fila_inicio_sem, fila_fin_sem = _formatear_tabla(
        ws, "Consumo semestral", ["Año-Semestre", "Consumo (kWh)"], filas_sem,
    )

    nota = (
        "Nota metodológica: el consumo (kWh) se prorratea día por día dentro del "
        "período de lectura de cada boleta. Cuando dos boletas cubren el mismo día "
        "(boleta 'sin lectura' + boleta correctiva posterior), se usa el dato de la "
        "boleta emitida más tarde para esos días, evitando contar el consumo dos "
        "veces. El gasto ($) se asigna al año/mes/trimestre/semestre de la fecha de "
        "emisión de cada boleta (no se solapa nunca, cada boleta se paga una sola vez)."
    )
    ws.cell(row=ws.max_row + 1, column=1, value=nota)
    ws.row_dimensions[ws.max_row].height = 30

    return {
        "anual": (fila_inicio_anual, fila_fin_anual),
        "mensual": (fila_inicio_mes, fila_fin_mes),
        "estacional": (fila_inicio_estacional, fila_fin_estacional),
        "trimestral": (fila_inicio_trim, fila_fin_trim),
        "semestral": (fila_inicio_sem, fila_fin_sem),
    }


def construir_graficos(wb, ws_metricas, rangos: dict) -> None:
    """Crea la hoja 'Gráficos' (separada de 'Metricas') con los graficos de
    barras, leyendo los datos de la hoja Metricas."""
    if "Gráficos" in wb.sheetnames:
        del wb["Gráficos"]
    ws = wb.create_sheet("Gráficos")
    ws["A1"] = "Gráficos"
    ws["A1"].font = Font(bold=True, size=14)

    fila_inicio_anual, fila_fin_anual = rangos["anual"]
    fila_inicio_mes, fila_fin_mes = rangos["mensual"]
    fila_inicio_trim, fila_fin_trim = rangos["trimestral"]

    categorias = Reference(ws_metricas, min_col=1, min_row=fila_inicio_anual, max_row=fila_fin_anual)

    chart_anual_kwh = BarChart()
    chart_anual_kwh.title = "Consumo anual (kWh)"
    datos = Reference(ws_metricas, min_col=2, min_row=fila_inicio_anual - 1, max_row=fila_fin_anual)
    chart_anual_kwh.add_data(datos, titles_from_data=True)
    chart_anual_kwh.set_categories(categorias)
    chart_anual_kwh.width, chart_anual_kwh.height = 18, 10
    _ajustar_ejes(chart_anual_kwh)
    ws.add_chart(chart_anual_kwh, "A3")

    chart_anual_gasto = BarChart()
    chart_anual_gasto.title = "Gasto anual ($)"
    datos = Reference(ws_metricas, min_col=3, min_row=fila_inicio_anual - 1, max_row=fila_fin_anual)
    chart_anual_gasto.add_data(datos, titles_from_data=True)
    chart_anual_gasto.set_categories(categorias)
    chart_anual_gasto.width, chart_anual_gasto.height = 18, 10
    _ajustar_ejes(chart_anual_gasto)
    ws.add_chart(chart_anual_gasto, "A22")

    chart_costo_kwh = BarChart()
    chart_costo_kwh.title = "Costo promedio por kWh ($/kWh)"
    datos = Reference(ws_metricas, min_col=4, min_row=fila_inicio_anual - 1, max_row=fila_fin_anual)
    chart_costo_kwh.add_data(datos, titles_from_data=True)
    chart_costo_kwh.set_categories(categorias)
    chart_costo_kwh.width, chart_costo_kwh.height = 18, 10
    _ajustar_ejes(chart_costo_kwh)
    ws.add_chart(chart_costo_kwh, "A41")

    chart_trim = BarChart()
    chart_trim.title = "Consumo trimestral (kWh)"
    datos = Reference(ws_metricas, min_col=2, min_row=fila_inicio_trim - 1, max_row=fila_fin_trim)
    categorias_trim = Reference(ws_metricas, min_col=1, min_row=fila_inicio_trim, max_row=fila_fin_trim)
    chart_trim.add_data(datos, titles_from_data=True)
    chart_trim.set_categories(categorias_trim)
    chart_trim.width, chart_trim.height = 26, 10
    _ajustar_ejes(chart_trim)
    ws.add_chart(chart_trim, "A60")

    chart_mensual = BarChart()
    chart_mensual.title = "Consumo mensual (kWh)"
    datos = Reference(ws_metricas, min_col=2, min_row=fila_inicio_mes - 1, max_row=fila_fin_mes)
    categorias_mes = Reference(ws_metricas, min_col=1, min_row=fila_inicio_mes, max_row=fila_fin_mes)
    chart_mensual.add_data(datos, titles_from_data=True)
    chart_mensual.set_categories(categorias_mes)
    chart_mensual.width, chart_mensual.height = 30, 10
    _ajustar_ejes(chart_mensual)
    ws.add_chart(chart_mensual, "A79")

    fila_inicio_estacional, fila_fin_estacional = rangos["estacional"]
    chart_estacional = BarChart()
    chart_estacional.title = "Promedio mensual histórico (kWh, Ene-Dic, años completos)"
    datos = Reference(ws_metricas, min_col=2, min_row=fila_inicio_estacional - 1, max_row=fila_fin_estacional)
    categorias_estacional = Reference(ws_metricas, min_col=1, min_row=fila_inicio_estacional, max_row=fila_fin_estacional)
    chart_estacional.add_data(datos, titles_from_data=True)
    chart_estacional.set_categories(categorias_estacional)
    chart_estacional.width, chart_estacional.height = 20, 10
    _ajustar_ejes(chart_estacional)
    ws.add_chart(chart_estacional, "A98")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dry-run", action="store_true", help="No escribe ningun archivo, solo informa.")
    parser.add_argument(
        "--version", action="version",
        version=f"actualizar_cuentas.py {__version__} ({__last_updated__}) — {__author__} <{__email__}>",
    )
    args = parser.parse_args()

    pdfs = buscar_pdfs(BASE_DIR)
    print(f"Carpeta: {BASE_DIR}")
    print(f"PDFs encontrados: {len(pdfs)}")

    registros_pdf = [parse_pdf(pdf) for pdf in pdfs]

    utilizables, no_procesadas = [], []
    vistos_nboleta: set[int] = set()
    for rec in registros_pdf:
        if rec.get("nboleta") is None:
            no_procesadas.append({**rec, "motivo": "No se pudo identificar el N° de boleta (texto del PDF corrupto)."})
            continue
        faltan = [c for c in CAMPOS_CRITICOS if rec.get(c) is None]
        if faltan:
            no_procesadas.append({**rec, "motivo": f"Faltan campos críticos: {faltan}"})
            continue
        if rec["nboleta"] in vistos_nboleta:
            no_procesadas.append({**rec, "motivo": f"N° de boleta duplicado entre PDFs (ya se uso de otro archivo)."})
            continue
        vistos_nboleta.add(rec["nboleta"])
        utilizables.append(rec)

    ruta_reporte = escribir_reporte_no_procesadas(BASE_DIR, no_procesadas)
    if ruta_reporte:
        print(f"\n{len(no_procesadas)} PDF no se pudieron incluir -> ver {ruta_reporte.name}")

    if not utilizables:
        print("\nNo hay boletas utilizables; no se genera ninguna planilla.")
        return

    nombre_salida = nombre_archivo_salida(utilizables)
    ruta_salida = BASE_DIR / nombre_salida
    print(f"\nBoletas incluidas: {len(utilizables)}")
    print(f"Archivo de salida: {nombre_salida}")

    if args.dry_run:
        print("\n[--dry-run] No se escribio ningun archivo.")
        return

    wb = openpyxl.Workbook()
    construir_hoja_resumen(wb, utilizables)
    rangos = construir_metricas(wb, utilizables)
    if rangos:
        construir_graficos(wb, wb["Metricas"], rangos)

    if ruta_salida.exists():
        respaldo = ruta_salida.with_name(
            f"{ruta_salida.stem}.respaldo_{datetime.now():%Y%m%d_%H%M%S}{ruta_salida.suffix}"
        )
        shutil.copy2(ruta_salida, respaldo)
        print(f"Ya existia un archivo con ese nombre; respaldo creado: {respaldo.name}")

    wb.save(ruta_salida)
    print(f"Planilla generada: {nombre_salida}")

    ruta_leeme = escribir_leeme(BASE_DIR)
    print(f"LEEME generado/actualizado: {ruta_leeme.name}")


if __name__ == "__main__":
    main()
