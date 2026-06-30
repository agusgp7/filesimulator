# Analizador de Consumos UTE

App estática para GitHub Pages que lee una planilla Excel con medidas cada 15 minutos y genera dashboard de energía activa AE y reactiva Q1.

## Archivos

- `index.html`
- `styles.css`
- `app.js`

## Cómo probar localmente

Abrir `index.html` en el navegador y cargar el Excel.

## Cómo subir a GitHub Pages

1. Crear un repositorio nuevo en GitHub.
2. Subir estos tres archivos a la raíz del repo.
3. Ir a Settings → Pages.
4. En Source elegir `Deploy from a branch`.
5. Elegir rama `main` y carpeta `/root`.
6. Abrir la URL publicada.

## Formato esperado del Excel

Debe tener columnas equivalentes a:

- `FECHA-HORA`
- `MAGNITUD`
- `VALOR`
- `INTERVALO`
- `RESULT VALIDACION`

La hoja puede llamarse `Measures`; si no existe, usa la primera hoja.
