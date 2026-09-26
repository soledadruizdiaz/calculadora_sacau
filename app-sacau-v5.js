const INDICES = {42: {1: 1, 2: 1.8, 3: 2.5}, 43: {1: .5, 2: 1.2, 3: 2}};
const COLUMNAS = ["Asignatura", "Tipo", "Nivel", "HIP"];
const NIVELES = {1: "Introductoria", 2: "Trayectoria", 3: "Integración/Finalización"};
let resultados = [];
let archivoSeleccionado = null;
let nombreArchivo = "plan_estudios";
let datosCarrera = null;
const $ = id => document.getElementById(id);
const formato = n => new Intl.NumberFormat("es-AR", {maximumFractionDigits: 2}).format(n);
const normalizar = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase();

function estado(mensaje, tipo = "") { $("status").textContent = mensaje; $("status").className = "notice " + tipo; }
function limpiar() { resultados = []; $("results").hidden = true; $("empty-results").hidden = false; $("excel-button").disabled = true; $("pdf-button").disabled = true; }
function cargarCarreras() {
  try {
    const guardadas = JSON.parse(localStorage.getItem("sacau_carreras") || "[]");
    if (Array.isArray(guardadas)) guardadas.filter(x => typeof x === "string").forEach(insertarCarrera);
    const facultades = JSON.parse(localStorage.getItem("sacau_facultades") || "[]");
    if (Array.isArray(facultades)) facultades.filter(x => typeof x === "string").forEach(insertarFacultad);
  } catch { /* El almacenamiento local puede estar deshabilitado. */ }
}
function insertarOpcion(id, nombre) {
  const select = $(id);
  const existente = [...select.options].find(option => normalizar(option.value) === normalizar(nombre));
  if (existente) return existente.value;
  select.add(new Option(nombre, nombre), select.querySelector('[value="__other__"]'));
  return nombre;
}
function insertarCarrera(nombre) {
  return insertarOpcion("career-select", nombre);
}
function insertarFacultad(nombre) {
  return insertarOpcion("faculty-select", nombre);
}
function persistir(nombre, clave) {
  try {
    const guardadas = JSON.parse(localStorage.getItem(clave) || "[]");
    if (Array.isArray(guardadas) && !guardadas.some(x => normalizar(x) === normalizar(nombre)))
      localStorage.setItem(clave, JSON.stringify([...guardadas, nombre]));
    return true;
  } catch { return false; }
}
function guardarCarrera() {
  const nombre = $("new-career").value.trim().replace(/\s+/g, " ");
  if (!nombre) { $("career-status").textContent = "Escribí el nombre de la carrera antes de agregarla."; return; }
  if (nombre.length > 120) { $("career-status").textContent = "El nombre no puede superar los 120 caracteres."; return; }
  const valor = insertarCarrera(nombre);
  $("career-select").value = valor;
  $("new-career-wrap").hidden = true;
  $("new-career").value = "";
  $("career-status").textContent = "Carrera agregada. Quedará disponible en este navegador.";
  if (!persistir(valor, "sacau_carreras")) $("career-status").textContent = "Carrera agregada para esta sesión. El navegador no permitió guardarla para la próxima visita.";
  limpiar();
}
function guardarFacultad() {
  const nombre = $("new-faculty").value.trim().replace(/\s+/g, " ");
  if (!nombre || nombre.length > 120) {$("career-status").textContent = "Escribí un nombre de facultad de hasta 120 caracteres."; return;}
  const valor = insertarFacultad(nombre);
  $("faculty-select").value = valor;
  $("new-faculty-wrap").hidden = true;
  $("new-faculty").value = "";
  $("career-status").textContent = persistir(valor, "sacau_facultades") ? "Facultad agregada." : "Facultad agregada para esta sesión.";
  limpiar();
}
function leerDatosExcel(workbook) {
  const nombreHoja = workbook.SheetNames.find(n => normalizar(n) === "datos de la carrera");
  if (!nombreHoja) return;
  const filas = XLSX.utils.sheet_to_json(workbook.Sheets[nombreHoja], {header: 1, defval: ""});
  const cabecera = (filas[0] || []).map(normalizar);
  const f = cabecera.indexOf("facultad"), c = cabecera.indexOf("carrera");
  if (f < 0 || c < 0) throw Error("La hoja «Datos de la carrera» debe tener las columnas Facultad y Carrera en la primera fila.");
  const registro = filas.slice(1).find(fila => String(fila[f] ?? "").trim() || String(fila[c] ?? "").trim());
  if (!registro) return;
  const facultad = String(registro[f] ?? "").trim().replace(/\s+/g," ");
  const carrera = String(registro[c] ?? "").trim().replace(/\s+/g," ");
  if (!facultad || !carrera || facultad.length > 120 || carrera.length > 120)
    throw Error("Completá Facultad y Carrera en la segunda fila de «Datos de la carrera» (hasta 120 caracteres cada una).");
  $("faculty-select").value = insertarFacultad(facultad);
  $("career-select").value = insertarCarrera(carrera);
  $("new-faculty-wrap").hidden = true;
  $("new-career-wrap").hidden = true;
  persistir(facultad, "sacau_facultades");
  persistir(carrera, "sacau_carreras");
  $("career-status").textContent = "Facultad y carrera cargadas desde el Excel.";
}
function leerDatosCarrera() {
  const facultad = $("faculty-select").value;
  const carrera = $("career-select").value;
  const modalidad = $("modality-select").value;
  const regimen = Number($("regime-select").value);
  if (!facultad || facultad === "__other__") throw Error("Seleccioná una facultad o agregá una nueva.");
  if (!carrera || carrera === "__other__") throw Error("Seleccioná una carrera o agregá una nueva.");
  if (!modalidad) throw Error("Seleccioná la modalidad de la carrera.");
  if (!Object.hasOwn(INDICES, regimen)) throw Error("Seleccioná el artículo 42 o 43.");
  return {facultad, carrera, modalidad, regimen};
}
function descargarPlantilla() {
  if (!window.XLSX) return estado("No se pudo cargar la biblioteca de Excel. Revisá tu conexión e intentá nuevamente.", "error");
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Facultad", "Carrera"], ["", ""]]), "Datos de la carrera");
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
    leerDatosExcel(workbook);
    const seleccion = leerDatosCarrera();
    const planName = workbook.SheetNames.find(n => normalizar(n) === "plan de estudios") || workbook.SheetNames.find(n => normalizar(n) !== "datos de la carrera");
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
  resultados.forEach(r => mapa.set(r.Nivel, (mapa.get(r.Nivel) || 0) + r.Total));
  return [...mapa.entries()].map(([nivel, horas]) => [nivel, Math.round(horas / 25)]);
}
function mostrar() {
  $("empty-results").hidden = true; $("results").hidden = false;
  $("career-summary").textContent = `${datosCarrera.facultad} · ${datosCarrera.carrera} · ${datosCarrera.modalidad} · Artículo ${datosCarrera.regimen}`;
  $("count").textContent = formato(resultados.length);
  $("hours").textContent = formato(resultados.reduce((s,r) => s + r.Total, 0));
  $("credits").textContent = formato(Math.round(resultados.reduce((s,r) => s + r.Total, 0) / 25));
  const body = $("results-body"); body.replaceChildren();
  resultados.forEach(r => {
    const tr = document.createElement("tr");
    [r.Asignatura, r.Tipo, r.Nivel, formato(r.HIP), formato(r.HTA), formato(r.Total), formato(r.CRE)].forEach((value,i) => {
      const td = document.createElement("td"); td.textContent = value; if (i > 2) td.className = "numeric"; tr.append(td);
    }); body.append(tr);
  });
  const summary = $("level-summary"); summary.replaceChildren();
  niveles().forEach(([nivel, cre]) => {const el = document.createElement("span"); const b = document.createElement("b"); b.textContent = nivel + ": "; el.append(b, `${formato(cre)} CRE`); summary.append(el);});
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
  sheet["!cols"] = [{wch: 40}, {wch: 9}, {wch: 25}, {wch: 12}, {wch: 12}, {wch: 12}, {wch: 12}];
  XLSX.utils.book_append_sheet(wb, sheet, "Créditos SACAU");
  XLSX.writeFile(wb, `${nombreArchivo}_sacau.xlsx`);
}
function descargarPDF() {
  if (!window.jspdf?.jsPDF) return estado("No se pudo cargar la biblioteca PDF. Revisá tu conexión e intentá nuevamente.", "error");
  const doc = new jspdf.jsPDF({orientation: "landscape"});
  if (!doc.autoTable) return estado("No se pudo cargar el generador de tablas PDF. Revisá tu conexión.", "error");
  doc.setFontSize(16); doc.text("Plan de estudios con créditos SACAU", 14, 16);
  doc.setFontSize(10);
  doc.text(doc.splitTextToSize(`Facultad: ${datosCarrera.facultad} · Carrera: ${datosCarrera.carrera} · Modalidad: ${datosCarrera.modalidad} · Artículo ${datosCarrera.regimen}`, 265), 14, 23);
  doc.autoTable({startY: 42, head: [["Asignatura", "Tipo", "Nivel", "HIP", "HTA", "Total", "CRE"]],
    body: resultados.map(r => [r.Asignatura, r.Tipo, r.Nivel, ...[r.HIP,r.HTA,r.Total,r.CRE].map(formato)]),
    theme: "grid", headStyles: {fillColor: [21, 75, 122]}, styles: {fontSize: 8, cellPadding: 2, overflow: "linebreak"},
    columnStyles: {0: {cellWidth: 95}, 1: {cellWidth: 17}, 2: {cellWidth: 50}}});
  let y = doc.lastAutoTable.finalY + 14;
  if (y > doc.internal.pageSize.getHeight() - 27) {doc.addPage(); y = 20;}
  doc.setFontSize(12); doc.text("Resumen de créditos por nivel", 14, y);
  doc.autoTable({startY: y + 4, head: [["Nivel", "CRE"]], body: niveles().map(([nivel, cre]) => [nivel, formato(cre)]),
    theme: "striped", headStyles: {fillColor: [21, 75, 122]}, styles: {fontSize: 9}, tableWidth: 100});
  doc.save(`${nombreArchivo}_sacau.pdf`);
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
$("career-select").addEventListener("change", () => {$("new-career-wrap").hidden = $("career-select").value !== "__other__"; $("career-status").textContent = ""; limpiar();});
$("faculty-select").addEventListener("change", () => {$("new-faculty-wrap").hidden = $("faculty-select").value !== "__other__"; $("career-status").textContent = ""; limpiar();});
$("modality-select").addEventListener("change", limpiar);
$("regime-select").addEventListener("change", limpiar);
$("save-career").addEventListener("click", guardarCarrera);
$("save-faculty").addEventListener("click", guardarFacultad);
cargarCarreras();
