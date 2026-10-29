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
