# Calculadora de créditos SACAU

Página para la Unidad de Transformación Curricular de la Universidad Atlántida. Lee automáticamente el catálogo de facultades y carreras que está en el repositorio, permite seleccionar modalidad y artículo 42 o 43, procesar un Excel de asignaturas, revisar resultados y descargar Excel y PDF. El archivo de materias se procesa en el navegador y no se guarda en ningún servidor.

## Formato de entrada

Hay **dos archivos independientes**:

- **Catálogo de facultades y carreras:** `catalogo_facultades_carreras.csv` queda en la raíz del repositorio y la página lo lee automáticamente al abrirse. Su primera fila debe ser exactamente `Facultad,Carrera`. Cada fila siguiente es una carrera asociada a una facultad. Si un nombre contiene coma, encerrarlo entre comillas dobles. Guardar como CSV UTF-8. Editar y confirmar este archivo en GitHub para que una nueva opción quede disponible para **todos**. El ZIP trae un ejemplo con cuatro carreras de la Facultad de Ingeniería. Elegir una facultad filtra sus carreras.
- **Plan de estudios:** Excel `.xls` o `.xlsx` con `Asignatura`, `Tipo`, `Nivel`, `HIP` en la primera fila y una fila por asignatura. No lleva Facultad ni Carrera. El botón **Descargar plantilla de Excel** genera este archivo.

La modalidad (Presencial o A Distancia) y el artículo (42 o 43) se seleccionan en la página. También es posible agregar manualmente una facultad y una carrera desde los combos, pero esa opción queda guardada **solo en ese navegador**. Para compartirla, agregar la fila al CSV del repositorio.

Las columnas adicionales de versiones anteriores se ignoran: los valores de Régimen y Modalidad utilizados serán siempre los seleccionados en la página. Tipo 1 corresponde a Nivel Introductoria; Tipo 2, a Trayectoria; Tipo 3, a Integración/Finalización (también se admite Integración o Finalización como abreviatura). `HIP` debe ser un número no negativo. La página ofrece una plantilla vacía. Al seleccionar el Excel se habilita el botón **Procesar archivo**; los cálculos comienzan solo al pulsarlo.

El Excel descargado incluye una hoja **Datos de la carrera** y otra con los resultados por asignatura. El PDF muestra los datos de la carrera encima de la tabla.

Los índices son los **ejemplos del script original**: Art. 42: 1,0 / 1,8 / 2,5; Art. 43: 0,5 / 1,2 / 2,0. Verificarlos antes de un uso institucional. `HTA = HIP × índice`, `Total = HIP + HTA`, `CRE = redondear(Total / 25)`. Los créditos por nivel y los créditos generales se calculan redondeando el total de horas correspondiente, no sumando los créditos redondeados de las asignaturas. Por ejemplo, 120/25 → 5 CRE y 288/25 → 12 CRE, mientras que 408/25 → 16 CRE en el total; por eso puede haber diferencias entre la suma de filas y el total.

## Subir a GitHub Pages

1. Crear un repositorio público en GitHub, por ejemplo `calculadora-sacau`.
2. Subir los seis archivos descomprimidos (`index.html`, `styles.css`, `app-sacau-v7.js`, `UA_logo_sinfondo.png`, `catalogo_facultades_carreras.csv`, `README.md`) a la **raíz** del repositorio, mediante **Add file → Upload files → Commit changes**. No subir el ZIP como único archivo. Si ya publicaste una versión anterior, reemplazá `index.html` y `styles.css` y agregá **tanto el JavaScript nuevo como el CSV**; el JavaScript anterior puede eliminarse.
3. En **Settings → Pages**, elegir **Deploy from a branch**, rama `main`, carpeta `/(root)` y pulsar **Save**.
4. La página quedará en `https://TU_USUARIO.github.io/calculadora-sacau/` una vez finalizada la publicación.

Las bibliotecas de Excel y PDF se descargan desde jsDelivr, por lo que se necesita internet para procesar archivos. Al publicarse en GitHub Pages, el código y la página son públicos; no hay inicio de sesión.
