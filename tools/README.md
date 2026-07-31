# Pipeline de datos MIP (Biofuturo) → `pest_data.json`

Convierte las carpetas de visita de Biofuturo (una por fecha) en un único JSON
normalizado, con la misma convención de catálogos + tabla larga que usan
`soil_data.json`, `foliar_data.json` y `agro_data.json`.

## Insumos por carpeta de visita

| Archivo | Aporta |
|---|---|
| `RESUMEN PROSPECCIÓN … .xlsx` | Cuantificación por plaga × cuartel (23 columnas `1.1`–`5.5`) |
| `*.kml` / `*.kmz` | Puntos georreferenciados con la plaga levantada en cada punto |
| `Manifold Post Visita … .pdf` | Informe cualitativo por plaga + observaciones generales |

## Scripts

### `build_pest_data.ps1`
Genera `../pest_data.json`. Requiere **Excel instalado** (lo lee por COM) y
PowerShell 5.1.

```powershell
powershell -ExecutionPolicy Bypass -File tools\build_pest_data.ps1
```

Parámetros útiles: `-Src` (carpeta con las visitas), `-OutFile`, `-Observations`.

Qué hace:

- **Planillas** → una lectura por `visita × plaga × cuartel`. Convierte la coma
  decimal, deja `value: null` y registra una incidencia cuando el valor no es
  numérico, y guarda los rangos textuales (`"1 a 5"`) en `rangeText`.
- **Cuarteles** → lee el literal `const GEOJSON` de `../index.html`, que es la
  fuente de verdad de los 23 polígonos. Así los cuarteles `1.1`–`5.5` de la
  planilla quedan con las mismas claves (`E1-S1` … `E5-S5`) que el resto del
  dashboard.
- **Puntos** → normaliza la etiqueta libre de cada punto (`acsors`, `acaros`,
  `arañita roja europea`, `FOCO PANONYCHUS ULMI`… ) a una plaga canónica más
  la especie, y hace *point-in-polygon* contra los cuarteles. Los puntos con
  coordenadas fuera del predio se **excluyen** de `points[]` (queda la razón en
  `issues[]`), y los que caen fuera de todo polígono quedan con `sector: null`
  más una incidencia `info` indicando el cuartel más cercano y su distancia.
- **`pointPests`** → catálogo de las plagas que efectivamente aparecen en
  `points[]`, con nombre, especie y conteo. Es un espacio de claves distinto al
  de `pests[]`: en terreno se marca "ácaros fitófagos móviles" sin separar
  perímetro e interior. El mapa filtra y arma su leyenda desde aquí.
- **Manifold** → toma el texto ya curado de `observations.json`.
- **Incidencias** → todo problema de calidad detectado queda en `issues[]` en
  vez de corregirse en silencio.

### `manifold_pdf_text.ps1`
Extractor de texto para los PDF de Manifold, sin dependencias externas
(descomprime los content streams y reconstruye las líneas por posición).
Se usa para volver a transcribir un informe nuevo a `observations.json`.

```powershell
powershell -ExecutionPolicy Bypass -File tools\manifold_pdf_text.ps1 `
  -Path "…\Manifold Post Visita … .pdf" -OutFile salida.txt
```

### `observations.json`
Parte cualitativa transcrita de los PDF, una entrada por visita: fenología,
fechas de cosecha, autor, un bloque por plaga (especie, estado de desarrollo,
ubicación, rango, enemigos naturales, comentarios) y las observaciones
generales. `pestRefs` enlaza cada bloque con las filas de la planilla y
`sectorRefs` con los cuarteles mencionados en el texto.

Esto es lo que alimenta el modal **"Ver informe de la visita"** de la pestaña
Plagas: los `comments` se muestran como párrafos, `generalObservations` como
lista numerada, cada `sectorRefs` como un chip que hace zoom a ese cuartel, y
`pestRefs` se cruza con `readings` para mostrar el rango realmente medido al
lado del rango que declara el informe. Los `pestRefs` se listan **uno por fila**
a propósito: un bloque suele citar métricas que no se pueden promediar juntas
(perímetro vs. interior, infestación vs. viabilidad).

## Agregar una visita nueva

1. Copiar la carpeta de la visita junto a las demás en `-Src`.
2. `manifold_pdf_text.ps1` sobre el PDF nuevo y agregar la entrada a
   `observations.json` (misma forma que las existentes).
3. Correr `build_pest_data.ps1`.
4. Revisar `issues[]` del JSON generado.
5. Actualizar la clave `pest` de `data-version.json` (rompe la caché del navegador).

Si aparece una plaga nueva en los KML, agregarla en cuatro lugares: `$LABEL_MAP`
(variantes de escritura → clave canónica), `$PEST_TAXA` (nombre y especie para
mostrar), `$TAXON_OF` (a qué taxón pertenece cada fila de la planilla) y
`PEST_COLORS` en `index.html` (color). Sin el color, la plaga cae al gris de
`_default`; sin `$TAXON_OF`, el generador lo avisa en `issues[]`.

## Paleta de plagas

`PEST_COLORS` en `index.html` es la única fuente del color, y `pests[].taxon`
del JSON es lo que permite que el mapa, la leyenda, los chips, el popup y las
tarjetas del informe usen el mismo color para la misma plaga.

Deliberadamente **no hay rojo ni verde**: sobre el mapa se leen como "algo está
mal" y "algo está bien", cuando lo que el color codifica es identidad, no estado.
El verde además se pierde contra la vegetación del satelital.

La paleta se validó con el validador de la guía de dataviz sobre el set real del
mapa —las 6 plagas con puntos más el gris neutro de "sin clasificar"— en modo
`--pairs all`, que es el que corresponde a una capa de puntos donde cualquier par
puede quedar contiguo:

| Compuerta | Resultado |
|---|---|
| Separación CVD (protan/deutan), all-pairs | ΔE **10,4** — objetivo ≥ 8 ✅ |
| Piso de visión normal, all-pairs | ΔE **15,9** — piso duro 15 ✅ |
| Banda de luminosidad + piso de croma | ✅ (el gris queda bajo el croma a propósito: es el rol neutro) |
| Contraste sobre superficie clara | 3 colores bajo 3:1 → exige etiquetas visibles, que están en leyenda, chips, popup y tarjetas |

Al cambiar la paleta, volver a correr el validador. Dos límites que ya se
midieron: **8 categorías cromáticas no pasan** las compuertas en all-pairs
(lo mejor encontrado fue CVD 5,7 / normal 14,0), y el gris tampoco se separa de
los magentas si se lo trata como una serie más. Por eso `trips` y
`chanchito-blanco` tienen color pero no aparecen en el mapa (no tienen puntos):
sólo encabezan una tarjeta del informe, nunca simultáneamente con las demás.

> Los scripts tienen BOM UTF-8 a propósito: sin él, PowerShell 5.1 lee los
> acentos de los literales como ANSI y los rompe.
