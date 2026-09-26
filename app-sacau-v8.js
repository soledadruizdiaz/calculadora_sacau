const INDICES = {42: {1: 1, 2: 1.8, 3: 2.5}, 43: {1: .5, 2: 1.2, 3: 2}};
const COLUMNAS = ["Asignatura", "Tipo", "Nivel", "HIP"];
const NIVELES = {1: "Introductoria", 2: "Trayectoria", 3: "Integración/Finalización"};
let resultados = [];
let archivoSeleccionado = null;
let nombreArchivo = "plan_estudios";
let datosCarrera = null;
const catalogo = [];
const $ = id => document.getElementById(id);
const formato = n => new Intl.NumberFormat("es-AR", {maximumFractionDigits: 2}).format(n);
const normalizar = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase();

function estado(mensaje, tipo = "") { $("status").textContent = mensaje; $("status").className = "notice " + tipo; }
function limpiar() { resultados = []; $("results").hidden = true; $("empty-results").hidden = false; $("excel-button").disabled = true; $("pdf-button").disabled = true; }
async function cargarCarreras() {
  try {
    if (!window.XLSX) throw Error("No se pudo cargar la biblioteca de Excel.");
    const respuesta = await fetch(`catalogo_facultades_carreras.csv?v=${Date.now()}`, {cache: "no-store"});
    if (!respuesta.ok) throw Error("No se encontró catalogo_facultades_carreras.csv en el repositorio.");
    const workbook = XLSX.read(await respuesta.arrayBuffer(), {type: "array", codepage: 65001});
    const filas = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {header: 1, defval: "", blankrows: false});
    parsearCatalogo(filas).forEach(p => registrarPar(p.facultad, p.carrera));
    $("catalog-status").textContent = "";
  } catch (err) { $("catalog-status").textContent = `${err.message} Comprobá que el CSV esté en la raíz del repositorio.`; }
  mostrarCarreras();
}
function insertarOpcion(id, nombre) {
  const select = $(id);
  const existente = [...select.options].find(option => normalizar(option.value) === normalizar(nombre));
  if (existente) return existente.value;
  select.add(new Option(nombre, nombre));
  return nombre;
}
function insertarCarrera(nombre) { return insertarOpcion("career-select", nombre); }
function insertarFacultad(nombre) { return insertarOpcion("faculty-select", nombre); }
function registrarPar(facultad, carrera) {
  const f = insertarFacultad(facultad);
  if (!catalogo.some(p => normalizar(p.facultad) === normalizar(f) && normalizar(p.carrera) === normalizar(carrera)))
    catalogo.push({facultad: f, carrera});
}
function mostrarCarreras() {
  const select = $("career-select");
  const previo = select.value;
  select.replaceChildren(new Option("Seleccioná una carrera", ""));
  const facultad = $("faculty-select").value;
  catalogo.filter(p => normalizar(p.facultad) === normalizar(facultad)).forEach(p => insertarCarrera(p.carrera));
  select.value = [...select.options].some(o => o.value === previo) ? previo : "";
}
function parsearCatalogo(filas) {
  const cabecera = (filas[0] || []).map(normalizar);
  const f = cabecera.indexOf("facultad"), c = cabecera.indexOf("carrera");
  if (f < 0 || c < 0) throw Error("El catálogo debe tener las columnas Facultad y Carrera en la primera fila.");
  const pares = [];
  const errores = [];
  filas.slice(1).forEach((fila, i) => {
    if (fila.every(v => !String(v ?? "").trim())) return;
    const facultad = String(fila[f] ?? "").trim().replace(/\s+/g, " ");
    const carrera = String(fila[c] ?? "").trim().replace(/\s+/g, " ");
    if (!facultad || !carrera || facultad.length > 120 || carrera.length > 120) errores.push(i+2);
    else pares.push({facultad, carrera});
  });
  if (errores.length) throw Error("Completá Facultad y Carrera (hasta 120 caracteres) en las filas: " + errores.slice(0,10).join(", "));
  if (!pares.length) throw Error("El catálogo no contiene facultades y carreras.");
  return pares;
}
function leerDatosCarrera() {
  const facultad = $("faculty-select").value;
  const carrera = $("career-select").value;
  const modalidad = $("modality-select").value;
  const regimen = Number($("regime-select").value);
  if (!facultad) throw Error("Seleccioná una facultad del catálogo.");
  if (!carrera || !catalogo.some(p => normalizar(p.facultad) === normalizar(facultad) && normalizar(p.carrera) === normalizar(carrera))) throw Error("Seleccioná una carrera del catálogo para esa facultad.");
  if (!modalidad) throw Error("Seleccioná la modalidad de la carrera.");
  if (!Object.hasOwn(INDICES, regimen)) throw Error("Seleccioná el artículo 42 o 43.");
  return {facultad, carrera, modalidad, regimen};
}
function descargarPlantilla() {
  if (!window.XLSX) return estado("No se pudo cargar la biblioteca de Excel. Revisá tu conexión e intentá nuevamente.", "error");
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([COLUMNAS]), "Plan de estudios");
  XLSX.writeFile(wb, "plantilla_plan_estudios.xlsx");
}
function leerNumero(value) {
  if (typeof value === "number") return value;
  const s = String(value ?? "").trim().replace(/\s/g, "");
  if (!s) return NaN;
  const decimal = s.includes(",") ? s.replace(/\./g, "").replace(",", ".") : s;
  return Number(decimal);
}
function seleccionar(file) {
  limpiar();
  archivoSeleccionado = null;
  $("process-button").disabled = true;
  $("file-label").textContent = file?.name || "Arrastrá tu archivo acá";
  if (!file) return estado("Seleccioná un archivo para comenzar.");
  if (!/\.xlsx?$/i.test(file.name)) return estado("Seleccioná un archivo de Excel con extensión .xls o .xlsx.", "error");
  archivoSeleccionado = file;
  $("process-button").disabled = false;
  estado(`Archivo seleccionado: ${file.name}. Presioná «Procesar archivo» para calcular.`);
}
async function procesar(file) {
  limpiar();
  if (!file) return;
  if (!/\.xlsx?$/i.test(file.name)) return estado("Seleccioná un archivo de Excel con extensión .xls o .xlsx.", "error");
  if (!window.XLSX) return estado("No se pudo cargar la biblioteca de Excel. Revisá tu conexión e intentá nuevamente.", "error");
  estado("Leyendo el archivo…");
  try {
    const workbook = XLSX.read(await file.arrayBuffer(), {type: "array"});
    const seleccion = leerDatosCarrera();
    const planName = workbook.SheetNames.find(n => normalizar(n) === "plan de estudios") || workbook.SheetNames[0];
    const sheet = workbook.Sheets[planName];
    if (!sheet) throw Error("El archivo no tiene hojas con datos.");
    const filas = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: "", blankrows: false});
    const encabezados = (filas[0] || []).map(normalizar);
    const faltantes = COLUMNAS.filter(c => !encabezados.includes(normalizar(c)));
    if (faltantes.length) throw Error("Faltan columnas: " + faltantes.join(", ") + ". Descargá la plantilla para ver el formato esperado.");
    const indice = Object.fromEntries(COLUMNAS.map(c => [c, encabezados.indexOf(normalizar(c))]));
    const errores = [];
    const datos = [];
    filas.slice(1).forEach((fila, i) => {
      if (fila.every(v => String(v ?? "").trim() === "")) return;
      const r = Object.fromEntries(COLUMNAS.map(c => [c, fila[indice[c]] ?? ""]));
      const tipo = leerNumero(r.Tipo), regimen = seleccion.regimen, hip = leerNumero(r.HIP);
      const problemas = [];
      if (!String(r.Asignatura).trim()) problemas.push("Asignatura vacía");
      if (!Object.hasOwn(INDICES[42], tipo)) problemas.push("Tipo debe ser 1, 2 o 3");
      const nivel = normalizar(r.Nivel).replace(/\s*\/\s*/g, "/");
      const nivelesValidos = tipo === 3 ? ["integracion", "finalizacion", "integracion/finalizacion"] : [normalizar(NIVELES[tipo])];
      if (Object.hasOwn(NIVELES, tipo) && !nivelesValidos.includes(nivel))
        problemas.push(`Nivel debe ser ${NIVELES[tipo]} para Tipo ${tipo}`);
      if (!Number.isFinite(hip) || hip < 0) problemas.push("HIP debe ser un número no negativo");
      if (problemas.length) { errores.push(`Fila ${i + 2}: ${problemas.join("; ")}`); return; }
      const hta = hip * INDICES[regimen][tipo];
      datos.push({Asignatura: String(r.Asignatura).trim(), Tipo: tipo, Nivel: NIVELES[tipo], HIP: hip, HTA: hta, Total: hip + hta, CRE: Math.round((hip + hta) / 25)});
    });
    if (errores.length) throw Error("Corregí el archivo y volvé a cargarlo:\n" + errores.slice(0, 8).join("\n") + (errores.length > 8 ? `\n…y ${errores.length - 8} fila(s) más.` : ""));
    if (!datos.length) throw Error("El archivo no contiene asignaturas debajo de los encabezados.");
    resultados = datos;
    datosCarrera = seleccion;
    nombreArchivo = file.name.replace(/\.xlsx?$/i, "").replace(/[^\w-]/g, "_") || "plan_estudios";
    mostrar();
    estado(`Listo: ${datos.length} asignatura(s) procesada(s) de ${file.name}.`, "success");
  } catch (err) { limpiar(); estado(err.message || "No se pudo leer el archivo.", "error"); }
}
function niveles() {
  const mapa = new Map();
  resultados.forEach(r => {if (!mapa.has(r.Nivel)) mapa.set(r.Nivel, []); mapa.get(r.Nivel).push(r);});
  return [...mapa.entries()].map(([nivel, filas]) => ({Nivel: nivel, ...totales(filas)}));
}
function totales(filas) {
  const HIP = filas.reduce((s,r) => s + r.HIP, 0);
  const HTA = filas.reduce((s,r) => s + r.HTA, 0);
  const Total = HIP + HTA;
  return {HIP, HTA, Total, CRE: Math.round(Total / 25)};
}
function agregarFilaTotales(id, etiqueta, valores, columnasPrevias) {
  const pie = $(id); pie.replaceChildren();
  const tr = document.createElement("tr");
  const th = document.createElement("th"); th.scope = "row"; th.colSpan = columnasPrevias; th.textContent = etiqueta; tr.append(th);
  [valores.HIP, valores.HTA, valores.Total, valores.CRE].forEach(n => {
    const td = document.createElement("td"); td.className = "numeric"; td.textContent = formato(n); tr.append(td);
  });
  pie.append(tr);
}
function mostrar() {
  $("empty-results").hidden = true; $("results").hidden = false;
  $("career-summary").textContent = `${datosCarrera.facultad} · ${datosCarrera.carrera} · ${datosCarrera.modalidad} · Artículo ${datosCarrera.regimen}`;
  const general = totales(resultados);
  $("count").textContent = formato(resultados.length);
  $("hours").textContent = formato(general.Total);
  $("credits").textContent = formato(general.CRE);
  const body = $("results-body"); body.replaceChildren();
  resultados.forEach(r => {
    const tr = document.createElement("tr");
    [r.Asignatura, r.Tipo, r.Nivel, formato(r.HIP), formato(r.HTA), formato(r.Total), formato(r.CRE)].forEach((value,i) => {
      const td = document.createElement("td"); td.textContent = value; if (i > 2) td.className = "numeric"; tr.append(td);
    }); body.append(tr);
  });
  agregarFilaTotales("results-total", "Totales", general, 3);
  const summary = $("level-summary"); summary.replaceChildren();
  niveles().forEach(r => {
    const tr = document.createElement("tr");
    [r.Nivel,r.HIP,r.HTA,r.Total,r.CRE].forEach((value,i) => {
      const td = document.createElement("td"); td.textContent = i ? formato(value) : value; if(i) td.className="numeric"; tr.append(td);
    }); summary.append(tr);
  });
  agregarFilaTotales("level-total", "Total general", general, 1);
  $("excel-button").disabled = false; $("pdf-button").disabled = false;
}
function descargarExcel() {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Facultad", "Carrera"], [datosCarrera.facultad, datosCarrera.carrera],
    [], ["Modalidad", datosCarrera.modalidad], ["Artículo de la Ley de Educación Superior", datosCarrera.regimen],
    [], ["Créditos totales", Math.round(resultados.reduce((s,r) => s + r.Total, 0) / 25)]
  ]), "Datos de la carrera");
  const sheet = XLSX.utils.json_to_sheet(resultados, {header: [...COLUMNAS, "HTA", "Total", "CRE"]});
  const general = totales(resultados);
  XLSX.utils.sheet_add_aoa(sheet, [["Totales", "", "", general.HIP, general.HTA, general.Total, general.CRE]], {origin: -1});
  sheet["!cols"] = [{wch: 40}, {wch: 9}, {wch: 25}, {wch: 12}, {wch: 12}, {wch: 12}, {wch: 12}];
  XLSX.utils.book_append_sheet(wb, sheet, "Créditos SACAU");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ["Nivel", "HIP", "HTA", "Total", "CRE"],
    ...niveles().map(r => [r.Nivel,r.HIP,r.HTA,r.Total,r.CRE]),
    ["Total general",general.HIP,general.HTA,general.Total,general.CRE]
  ]), "Resumen por nivel");
  XLSX.writeFile(wb, `${nombreArchivo}_sacau.xlsx`);
}
async function descargarPDF() {
  if (!window.jspdf?.jsPDF) return estado("No se pudo cargar la biblioteca PDF. Revisá tu conexión e intentá nuevamente.", "error");
  try {
    const response = await fetch("UA_logo_sinfondo.png");
    if (!response.ok) throw Error("No se pudo cargar el logo institucional para el PDF.");
    const bytes = new Uint8Array(await response.arrayBuffer());
    let binary = ""; bytes.forEach(b => { binary += String.fromCharCode(b); });
    const logo = "data:image/png;base64," + btoa(binary);
    const doc = new window.jspdf.jsPDF({orientation: "landscape"});
    if (!doc.autoTable) throw Error("No se pudo cargar el generador de tablas PDF.");
    doc.addImage(logo, "PNG", 14, 9, 63, 17.6);
    doc.setFontSize(16); doc.text("Plan de estudios con créditos SACAU", 84, 19);
    doc.setFontSize(10);
    doc.text(doc.splitTextToSize(`Facultad: ${datosCarrera.facultad} · Carrera: ${datosCarrera.carrera} · Modalidad: ${datosCarrera.modalidad} · Artículo ${datosCarrera.regimen}`, 265), 14, 33);
    const general = totales(resultados);
    doc.autoTable({startY: 48, head: [["Asignatura", "Tipo", "Nivel", "HIP", "HTA", "Total", "CRE"]],
      body: resultados.map(r => [r.Asignatura, r.Tipo, r.Nivel, ...[r.HIP,r.HTA,r.Total,r.CRE].map(formato)]),
      foot: [["Totales", "", "", ...[general.HIP,general.HTA,general.Total,general.CRE].map(formato)]],
      theme: "grid", showFoot: "lastPage", headStyles: {fillColor: [21, 75, 122]}, footStyles: {fillColor: [225, 237, 250], textColor: [20, 49, 88], fontStyle: "bold"},
      styles: {fontSize: 8, cellPadding: 2, overflow: "linebreak"},
      columnStyles: {0: {cellWidth: 95}, 1: {cellWidth: 17}, 2: {cellWidth: 50}}});
    let y = doc.lastAutoTable.finalY + 14;
    if (y > doc.internal.pageSize.getHeight() - 45) {doc.addPage(); y = 20;}
    doc.setFontSize(12); doc.text("Resumen por nivel", 14, y);
    doc.autoTable({startY: y + 4, head: [["Nivel", "HIP", "HTA", "Total", "CRE"]],
      body: niveles().map(r => [r.Nivel, ...[r.HIP,r.HTA,r.Total,r.CRE].map(formato)]),
      foot: [["Total general", ...[general.HIP,general.HTA,general.Total,general.CRE].map(formato)]],
      theme: "striped", headStyles: {fillColor: [21, 75, 122]}, footStyles: {fillColor: [225, 237, 250], textColor: [20, 49, 88], fontStyle: "bold"}, styles: {fontSize: 9}, tableWidth: 180});
    doc.save(`${nombreArchivo}_sacau.pdf`);
  } catch (err) {estado(err.message || "No se pudo generar el PDF.", "error");}
}
$("file-input").addEventListener("change", e => seleccionar(e.target.files[0]));
$("process-button").addEventListener("click", () => procesar(archivoSeleccionado));
const dropzone = $("dropzone");
["dragenter", "dragover"].forEach(event => dropzone.addEventListener(event, e => {e.preventDefault(); dropzone.classList.add("dragover");}));
["dragleave", "drop"].forEach(event => dropzone.addEventListener(event, e => {e.preventDefault(); dropzone.classList.remove("dragover");}));
dropzone.addEventListener("drop", e => seleccionar(e.dataTransfer.files[0]));
$("template-button").addEventListener("click", descargarPlantilla);
$("excel-button").addEventListener("click", descargarExcel);
$("pdf-button").addEventListener("click", descargarPDF);
$("career-select").addEventListener("change", limpiar);
$("faculty-select").addEventListener("change", () => {mostrarCarreras(); limpiar();});
$("modality-select").addEventListener("change", limpiar);
$("regime-select").addEventListener("change", limpiar);
cargarCarreras();
