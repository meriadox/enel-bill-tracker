# Cómo contribuir

Gracias por el interés en mejorar `enel-bill-tracker`. El proyecto se
mantiene como un script de uso personal que se comparte públicamente, así
que el flujo de contribución es simple:

1. **Haz un fork** del repositorio (no se aceptan pushes directos a `main`).
2. Crea una rama descriptiva para tu cambio: `git checkout -b fix/iva-regex`.
3. Prueba tu cambio contra boletas reales propias (o describe cómo lo
   probaste en el Pull Request) **antes de subir datos**: revisa que no se
   incluya ningún PDF, Excel con datos personales, ni cualquier otro
   archivo con información identificable — `.gitignore` ya excluye los
   patrones más comunes, pero revísalo igual.
4. Abre un **Pull Request** contra `main` explicando qué cambia y por qué.
   Si agregas o modificas una columna de `CARGOS_DETALLADOS`, incluye en
   la descripción del PR un fragmento de texto de boleta (anonimizado) que
   muestre el formato exacto que intentas capturar.
5. Mantén el estilo del código: sin dependencias nuevas salvo que sean
   imprescindibles, comentarios solo donde el *por qué* no sea obvio, y
   nada de datos ni rutas hardcodeadas.

## Qué tipo de contribuciones son más útiles

- Soporte para otros formatos de boleta Enel (otras tarifas, versiones de
  layout más nuevas) o de otras distribuidoras eléctricas chilenas.
- Mejoras a la extracción de cargos (IVA, descuentos promocionales) que
  mantengan el criterio de "si no se puede leer con certeza, `s/i`, nunca
  un número adivinado".
- Nuevas métricas o gráficos en la hoja `Metricas`/`Gráficos`.
- Reportar (como *issue*) un formato de boleta que el script no reconoce,
  idealmente adjuntando el texto plano extraído con
  `pdftotext -layout tu_boleta.pdf -` (nunca el PDF original ni datos
  personales).

## Licencia de las contribuciones

Al enviar un Pull Request aceptas que tu contribución se publique bajo la
misma licencia MIT del proyecto (ver [`LICENSE`](LICENSE)).

## Atribución

Si reusas este código fuera de este repositorio (por ejemplo, para un
proyecto propio basado en él), conserva el aviso de autoría original
(Rodrigo Meza López) tal como exige la licencia MIT.
