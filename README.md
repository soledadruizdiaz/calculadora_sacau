# Calculadora de créditos SACAU

Página para la Unidad de Transformación Curricular de la Universidad Atlántida. Permite cargar un Excel, calcular HTA, Total y CRE, revisar los resultados y descargar un Excel y un PDF. El archivo se procesa en el navegador y no se guarda en ningún servidor.

## Formato de entrada

La primera hoja debe comenzar con los encabezados `Asignatura`, `Tipo`, `Regimen`, `Modalidad`, `Nivel`, `HIP`. Se admite también `Régimen` con tilde. Tipo 1 corresponde a Nivel Introductoria; Tipo 2, a Trayectoria; Tipo 3, a Integración/Finalización. `Regimen` debe ser 42 o 43; `Modalidad`, Presencial o A Distancia; `HIP`, un número no negativo. La página ofrece una plantilla vacía. Al seleccionar el Excel se habilita el botón **Procesar archivo**; los cálculos comienzan solo al pulsarlo.

Los índices son los **ejemplos del script original**: Art. 42: 1,0 / 1,8 / 2,5; Art. 43: 0,5 / 1,2 / 2,0. Verificarlos antes de un uso institucional. `HTA = HIP × índice`, `Total = HIP + HTA`, `CRE = redondear(Total / 25)`. Los créditos por nivel y los créditos generales se calculan redondeando el total de horas correspondiente, no sumando los créditos redondeados de las asignaturas. Por ejemplo, 120/25 → 5 CRE y 288/25 → 12 CRE, mientras que 408/25 → 16 CRE en el total; por eso puede haber diferencias entre la suma de filas y el total.

## Subir a GitHub Pages

1. Crear un repositorio público en GitHub, por ejemplo `calculadora-sacau`.
2. Subir los cinco archivos descomprimidos (`index.html`, `styles.css`, `app.js`, `UA_logo_sinfondo.png`, `README.md`) a la **raíz** del repositorio, mediante **Add file → Upload files → Commit changes**. No subir el ZIP como único archivo.
3. En **Settings → Pages**, elegir **Deploy from a branch**, rama `main`, carpeta `/(root)` y pulsar **Save**.
4. La página quedará en `https://TU_USUARIO.github.io/calculadora-sacau/` una vez finalizada la publicación.

Las bibliotecas de Excel y PDF se descargan desde jsDelivr, por lo que se necesita internet para procesar archivos. Al publicarse en GitHub Pages, el código y la página son públicos; no hay inicio de sesión.
