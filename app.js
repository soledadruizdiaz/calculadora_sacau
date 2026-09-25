const INDICES = {42: {1: 1, 2: 1.8, 3: 2.5}, 43: {1: .5, 2: 1.2, 3: 2}};
const COLUMNAS = ["Asignatura", "Tipo", "Regimen", "Modalidad", "Nivel", "HIP"];
let resultados = [];
let nombreArchivo = "plan_estudios";
const $ = id => document.getElementById(id);
const formato = n => new Intl.NumberFormat("es-AR", {maximumFractionDigits: 2}).format(n);
const normalizar = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase();

function estado(mensaje, tipo = "") { $("status").textContent = mensaje; $("status").className = "notice " + tipo; }
function limpiar() { resultados = []; $("results").hidden = true; $("empty-results").hidden = false; $("excel-button").disabled = true; $("pdf-button").disabled = true; }
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
async function procesar(file) {
  limpiar();
  if (!file) return;
  if (!/\.xlsx?$/i.test(file.name)) return estado("Seleccioná un archivo de Excel con extensión .xls o .xlsx.", "error");
  if (!window.XLSX) return estado("No se pudo cargar la biblioteca de Excel. Revisá tu conexión e intentá nuevamente.", "error");
  estado("Leyendo el archivo…");
  try {
    const workbook = XLSX.read(await file.arrayBuffer(), {type: "array"});
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
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
      const tipo = leerNumero(r.Tipo), regimen = leerNumero(r.Regimen), hip = leerNumero(r.HIP);
      const problemas = [];
      if (!String(r.Asignatura).trim()) problemas.push("Asignatura vacía");
      if (!Object.hasOwn(INDICES[42], tipo)) problemas.push("Tipo debe ser 1, 2 o 3");
      if (!Object.hasOwn(INDICES, regimen)) problemas.push("Regimen debe ser 42 o 43");
      if (!String(r.Nivel).trim()) problemas.push("Nivel vacío");
      if (!Number.isFinite(hip) || hip < 0) problemas.push("HIP debe ser un número no negativo");
      if (problemas.length) { errores.push(`Fila ${i + 2}: ${problemas.join("; ")}`); return; }
      const hta = hip * INDICES[regimen][tipo];
      datos.push({Asignatura: String(r.Asignatura).trim(), Tipo: tipo, Regimen: regimen, Modalidad: String(r.Modalidad).trim(), Nivel: String(r.Nivel).trim(), HIP: hip, HTA: hta, Total: hip + hta, CRE: (hip + hta) / 25});
    });
    if (errores.length) throw Error("Corregí el archivo y volvé a cargarlo:\n" + errores.slice(0, 8).join("\n") + (errores.length > 8 ? `\n…y ${errores.length - 8} fila(s) más.` : ""));
    if (!datos.length) throw Error("El archivo no contiene asignaturas debajo de los encabezados.");
    resultados = datos;
    nombreArchivo = file.name.replace(/\.xlsx?$/i, "").replace(/[^\w-]/g, "_") || "plan_estudios";
    mostrar();
    estado(`Listo: ${datos.length} asignatura(s) procesada(s) de ${file.name}.`, "success");
  } catch (err) { limpiar(); estado(err.message || "No se pudo leer el archivo.", "error"); }
}
function niveles() {
  const mapa = new Map();
  resultados.forEach(r => mapa.set(r.Nivel, (mapa.get(r.Nivel) || 0) + r.CRE));
  return [...mapa.entries()];
}
function mostrar() {
  $("empty-results").hidden = true; $("results").hidden = false;
  $("count").textContent = formato(resultados.length);
  $("hours").textContent = formato(resultados.reduce((s,r) => s + r.Total, 0));
  $("credits").textContent = formato(resultados.reduce((s,r) => s + r.CRE, 0));
  const body = $("results-body"); body.replaceChildren();
  resultados.forEach(r => {
    const tr = document.createElement("tr");
    [r.Asignatura, r.Tipo, r.Regimen, r.Modalidad, r.Nivel, formato(r.HIP), formato(r.HTA), formato(r.Total), formato(r.CRE)].forEach((value,i) => {
      const td = document.createElement("td"); td.textContent = value; if (i > 4) td.className = "numeric"; tr.append(td);
    }); body.append(tr);
  });
  const summary = $("level-summary"); summary.replaceChildren();
  niveles().forEach(([nivel, cre]) => {const el = document.createElement("span"); const b = document.createElement("b"); b.textContent = nivel + ": "; el.append(b, `${formato(cre)} CRE`); summary.append(el);});
  $("excel-button").disabled = false; $("pdf-button").disabled = false;
}
function descargarExcel() {
  const wb = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(resultados, {header: [...COLUMNAS, "HTA", "Total", "CRE"]});
  sheet["!cols"] = [{wch: 40}, {wch: 9}, {wch: 12}, {wch: 20}, {wch: 16}, {wch: 12}, {wch: 12}, {wch: 12}, {wch: 12}];
  XLSX.utils.book_append_sheet(wb, sheet, "Créditos SACAU");
  XLSX.writeFile(wb, `${nombreArchivo}_sacau.xlsx`);
}
function descargarPDF() {
  if (!window.jspdf?.jsPDF) return estado("No se pudo cargar la biblioteca PDF. Revisá tu conexión e intentá nuevamente.", "error");
  const doc = new jspdf.jsPDF({orientation: "landscape"});
  if (!doc.autoTable) return estado("No se pudo cargar el generador de tablas PDF. Revisá tu conexión.", "error");
  doc.setFontSize(16); doc.text("Plan de estudios con créditos SACAU", 14, 16);
  doc.autoTable({startY: 23, head: [["Asignatura", "Tipo", "Régimen", "Modalidad", "Nivel", "HIP", "HTA", "Total", "CRE"]],
    body: resultados.map(r => [r.Asignatura, r.Tipo, r.Regimen, r.Modalidad, r.Nivel, ...[r.HIP,r.HTA,r.Total,r.CRE].map(formato)]),
    theme: "grid", headStyles: {fillColor: [21, 75, 122]}, styles: {fontSize: 8, cellPadding: 2, overflow: "linebreak"},
    columnStyles: {0: {cellWidth: 58}, 1: {cellWidth: 13}, 2: {cellWidth: 19}, 3: {cellWidth: 32}, 4: {cellWidth: 28}}});
  let y = doc.lastAutoTable.finalY + 14;
  if (y > doc.internal.pageSize.getHeight() - 27) {doc.addPage(); y = 20;}
  doc.setFontSize(12); doc.text("Resumen de créditos por nivel", 14, y);
  doc.autoTable({startY: y + 4, head: [["Nivel", "CRE"]], body: niveles().map(([nivel, cre]) => [nivel, formato(cre)]),
    theme: "striped", headStyles: {fillColor: [21, 75, 122]}, styles: {fontSize: 9}, tableWidth: 100});
  doc.save(`${nombreArchivo}_sacau.pdf`);
}
$("file-input").addEventListener("change", e => procesar(e.target.files[0]));
const dropzone = $("dropzone");
["dragenter", "dragover"].forEach(event => dropzone.addEventListener(event, e => {e.preventDefault(); dropzone.classList.add("dragover");}));
["dragleave", "drop"].forEach(event => dropzone.addEventListener(event, e => {e.preventDefault(); dropzone.classList.remove("dragover");}));
dropzone.addEventListener("drop", e => procesar(e.dataTransfer.files[0]));
$("template-button").addEventListener("click", descargarPlantilla);
$("excel-button").addEventListener("click", descargarExcel);
$("pdf-button").addEventListener("click", descargarPDF);
