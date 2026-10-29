# enel-bill-tracker

[![Licencia: MIT](https://img.shields.io/badge/Licencia-MIT-yellow.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)

Convierte tus boletas eléctricas de **Enel (Chile)** en PDF a una planilla
Excel limpia, con el desglose de cargos de cada boleta y una hoja de
métricas (consumo y gasto anual/mensual/trimestral/semestral, promedio
mensual histórico, costo por kWh, y gráficos de barras nativos de Excel).

No requiere copiar manualmente ningún dato: apunta el script a una carpeta
con tus boletas PDF y genera la planilla automáticamente.

> Ver [`ejemplo/ejemplo_resumen_boletas.xlsx`](ejemplo/ejemplo_resumen_boletas.xlsx)
> para un resultado real (con el N° de boleta y otros identificadores
> eliminados) generado a partir de boletas de un usuario real: 3 hojas
> (`Resumen Boletas`, `Metricas`, `Gráficos`) con el desglose completo de
> cargos y las métricas descritas más abajo.

## ¿Para quién es esto?

Para cualquier cliente residencial de Enel en Chile que reciba sus boletas
en PDF y quiera llevar un registro histórico de consumo y gasto sin
tipearlo a mano. Fue desarrollado y probado contra boletas con tarifa
**BT1** (tarifa residencial estándar) emitidas entre 2021 y 2026; otras
tarifas o un rediseño futuro de la boleta pueden requerir ajustar las
expresiones regulares de extracción (ver `CARGOS_DETALLADOS` y las
funciones `extraer_campos_criticos` / `extraer_cargos` en el script).

## Qué hace

1. Lee todos los `.pdf`/`.PDF` de la carpeta donde se ejecuta, con
   `pdftotext` (con un segundo método de respaldo para PDFs con texto
   corrupto).
2. Extrae de cada boleta: N° de boleta, período de lectura, consumo (kWh),
   total a pagar, fecha de emisión/vencimiento, tipo de tarifa, y el
   desglose de cargos (servicio eléctrico, transporte, cargo fondo de
   estabilización, arriendo de medidor, IVA cuando está disponible, etc.).
3. Genera **desde cero** una planilla `.xlsx` nueva (nunca modifica un
   Excel que ya exista en la carpeta), nombrada según el rango de años de
   las boletas encontradas, con 3 hojas: `Resumen Boletas`, `Metricas` y
   `Gráficos`.
4. Boletas que no se puedan leer del todo quedan listadas en
   `boletas_no_procesadas.txt` para revisión manual, y nunca se insertan
   con datos adivinados — un cargo ausente se marca como `s/i`.
5. Genera `LEEME.md` junto a la planilla, documentando los supuestos de
   cálculo aplicados (ver también [`docs/SUPUESTOS_DE_CALCULO.md`](docs/SUPUESTOS_DE_CALCULO.md)).

## Instalación

Requiere **Python 3.10+** y el binario **`pdftotext`** (parte de
[Poppler](https://poppler.freedesktop.org/), no es un paquete de Python).

```bash
# Poppler / pdftotext
#   Windows:  choco install poppler   (o descarga poppler para Windows y agrega su /bin al PATH)
#   macOS:    brew install poppler
#   Linux:    sudo apt install poppler-utils   (Debian/Ubuntu)

pip install -r requirements.txt
```

## Uso

1. Copia `actualizar_cuentas.py` a la carpeta donde guardas tus boletas PDF
   de Enel (puede haber también un `.xlsx` antiguo tuyo en esa carpeta: el
   script nunca lo toca).
2. Desde esa carpeta:

```bash
python actualizar_cuentas.py            # genera la planilla (con respaldo si ya existía una con ese nombre)
python actualizar_cuentas.py --dry-run   # solo muestra qué haría, no escribe nada
python actualizar_cuentas.py --version   # muestra versión y datos de contacto
```

3. Se genera `Resumen consumo electrico - AAAA-AAAA.xlsx` (el rango de años
   se calcula solo) junto con `LEEME.md` y, si hubo boletas no procesables,
   `boletas_no_procesadas.txt`.

## Limitaciones conocidas

- El IVA casi nunca aparece en el mismo renglón que su monto en el texto
  del PDF (depende del layout, que cambió varias veces entre 2021 y 2026);
  esa columna queda en `s/i` en la mayoría de los casos.
- Los descuentos/compensaciones promocionales (PVC, SERNAC, bonos por
  temporales, etc.) no se extraen como columna: se redactan como oraciones
  de texto libre con montos incrustados, demasiado variables para leer con
  confianza.
- Algunos PDFs antiguos (observado en boletas de mediados de 2021) usan una
  fuente sin tabla Unicode y su texto no se puede extraer por completo; el
  script recupera al menos el N° de boleta con un método alternativo, pero
  si falta el período o el total a pagar, la boleta no se incluye.
- El desglose de cargos se extrae con expresiones regulares ancladas a
  etiquetas específicas en español, pensadas para el formato de Enel Chile.
  No funcionará tal cual con boletas de otras distribuidoras o países.

Más detalle metodológico (cómo se prorratea el consumo cuando las boletas
se solapan, cómo se definen los "años completos" para los promedios, etc.) en
[`docs/SUPUESTOS_DE_CALCULO.md`](docs/SUPUESTOS_DE_CALCULO.md).

¿Tu boleta no se lee bien? Abre un
[issue con la plantilla "Boleta no reconocida"](../../issues/new?template=boleta-no-reconocida.yml)
— sin adjuntar el PDF ni datos personales, solo el fragmento de texto plano
relevante.

## Autoría y licencia

Desarrollado por **Rodrigo Meza López** (rodrigo.meza.lopez@gmail.com), con
asistencia de Claude (Anthropic).

Publicado bajo licencia **[MIT](LICENSE)**: puedes usar, modificar y
redistribuir este código libremente, **conservando el aviso de autoría**
(el encabezado del archivo y/o el archivo `LICENSE`). Si quieres proponer
cambios, por favor hazlo en una rama/fork y abre un Pull Request — ver
[`CONTRIBUTING.md`](CONTRIBUTING.md).

## Aviso

Este proyecto no tiene ninguna afiliación con Enel. Es una herramienta
independiente para uso personal de clientes que reciben sus propias
boletas en PDF.
