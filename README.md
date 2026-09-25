# Calculadora de créditos SACAU

Página estática para cargar un Excel con el plan de estudios, calcular HTA, horas totales y CRE, y descargar los resultados en Excel y PDF. El archivo se procesa en el navegador.

## Formato de entrada

La primera hoja debe tener encabezados `Asignatura`, `Tipo`, `Regimen`, `Modalidad`, `Nivel`, `HIP` en la primera fila. Se acepta también `Régimen` con tilde. `Tipo` debe ser 1, 2 o 3; `Regimen`, 42 o 43; `HIP`, un número no negativo. La página permite descargar una plantilla vacía.

## Cálculo

Los índices de ponderación son los valores de ejemplo del script original y deben revisarse antes de un uso institucional:

| Régimen | Tipo 1 | Tipo 2 | Tipo 3 |
| --- | ---: | ---: | ---: |
| Art. 42 | 1,0 | 1,8 | 2,5 |
| Art. 43 | 0,5 | 1,2 | 2,0 |

`HTA = HIP × índice`, `Total = HIP + HTA`, `CRE = Total / 25`. Los cálculos conservan precisión completa; la interfaz y el PDF muestran hasta dos decimales.

## Publicar en GitHub Pages

1. Creá un repositorio público en GitHub (por ejemplo `calculadora-sacau`).
2. Subí `index.html`, `styles.css`, `app.js` y `README.md` a la raíz del repositorio. Desde la web de GitHub podés hacerlo con **Add file → Upload files → Commit changes**.
3. Abrí **Settings → Pages → Build and deployment**. En **Source** elegí **Deploy from a branch**; en **Branch**, `main` y `/(root)`; presioná **Save**.
4. En esa misma pantalla aparecerá la dirección `https://TU_USUARIO.github.io/calculadora-sacau/` cuando la publicación termine.

Las dependencias de Excel y PDF se descargan desde jsDelivr al abrir la página, por lo que el usuario necesita conexión a internet. No se necesita instalar Python ni configurar un servidor.
