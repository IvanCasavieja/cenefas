# Generador de cenefas desde PPTX

Aplicacion web simple para tomar un template `.pptx` con placeholders como `{{COD}}`, cruzarlo con un archivo de datos `.xlsx` o `.csv`, y descargar un nuevo `.pptx` procesado.

## Que hace esta version

- Sube un template base `.pptx`.
- Sube un archivo de datos `.xlsx` o `.csv`.
- Detecta placeholders dentro del PowerPoint.
- Muestra un resumen con registros, columnas, placeholders, coincidencias y advertencias.
- Genera un `.pptx` final duplicando las slides template contiguas que contienen placeholders, una vez por cada fila del dataset.

## Stack

- Node.js + Express
- `exceljs` para leer `.xlsx`
- `csv-parse` para leer `.csv`
- `jszip` + Open XML para leer y reescribir el contenido del `.pptx`
- Frontend estatico con HTML, CSS y JavaScript

## Como correrlo

1. Instalar dependencias:

```powershell
npm.cmd install
```

2. Opcional: generar un template demo para pruebas rapidas:

```powershell
npm.cmd run make:example-template
```

3. Levantar la app:

```powershell
npm.cmd start
```

4. Abrir en el navegador:

```text
http://localhost:3000
```

## Archivos de ejemplo

- Datos de prueba: [examples/productos-demo.csv](examples/productos-demo.csv)
- Template demo opcional: `examples/template-demo.pptx` luego de correr `npm.cmd run make:example-template`

## Decisiones tecnicas

- El procesamiento se hace directamente sobre el `.pptx` como archivo Open XML. No se convierte a PDF ni a imagen.
- Para la primera version se priorizo robustez: la app repite slides template por fila de datos, en lugar de intentar acomodar varios productos dentro de la misma diapositiva.
- Cuando un placeholder esta repartido entre varios runs dentro del mismo parrafo, el motor intenta reconstruir el texto usando el formato del primer run disponible. Eso preserva posicion, caja, alineacion y la mayor parte del estilo, aunque puede aplanar formato mixto dentro de ese parrafo puntual.
- Si falta una columna para un placeholder, el placeholder queda visible en la salida y ademas se informa en advertencias.
- La estructura ya deja un punto unico de salida en `src/lib/output-service.js`, pensado para agregar una exportacion futura a PDF.

## Limitaciones conocidas

- La version actual no distribuye varias filas dentro de una misma slide; genera una copia por registro.
- Si un placeholder esta cortado entre distintos parrafos, esa situacion no se recompone automaticamente.
- Solo se repiten en bloque las slides contiguas que contienen placeholders. Si hay slides intermedias sin placeholders que tambien deberian repetirse, hoy conviene incluir al menos un placeholder en ellas o ajustar la logica en una segunda iteracion.
- No se implemento exportacion a PDF en esta version.

## Estructura principal

```text
index.html
app.js
styles.css
src/
  server.js
  lib/
    data-parser.js
    pptx-reader.js
    pptx-generator.js
    output-service.js
    placeholder-utils.js
    storage.js
```
