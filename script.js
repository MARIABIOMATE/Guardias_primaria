// ============================================================
// BLOQUE 1 - CONFIGURACIÓN Y CARGA DE DATOS DESDE EXCEL
// ============================================================

const URL_EXCEL = "https://raw.githubusercontent.com/MARIABIOMATE/Guardias_primaria/main/HORARIOS_GUARDIAS.xlsx";

let datosHorario = [];
let datosGuardias = [];
let datosPastoral = [];
let profesores = [];

window.addEventListener("DOMContentLoaded", async () => {
  await cargarExcel();
  cargarAusencias();
  document.getElementById("add-ausencia").addEventListener("click", registrarAusencia);
  document.getElementById("generar-cuadrante").addEventListener("click", generarCuadrante);
  document.getElementById("descargar-excel").addEventListener("click", descargarExcel);
  document.getElementById("volver").addEventListener("click", mostrarPanelAusencias);
});

async function cargarExcel() {
  try {
    const response = await fetch(URL_EXCEL);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    datosHorario = XLSX.utils.sheet_to_json(workbook.Sheets["HORARIO"]);
    datosGuardias = XLSX.utils.sheet_to_json(workbook.Sheets["GUARDIAS"]);
    datosPastoral = XLSX.utils.sheet_to_json(workbook.Sheets["PASTORAL_CONVIVENCIA"]);

    const profes = new Set();
    datosHorario.forEach(f => {
      if (f["Profesor de referencia"]) profes.add(f["Profesor de referencia"]);
      if (f["Profesor de apoyo 1"]) profes.add(f["Profesor de apoyo 1"]);
      if (f["Profesor de apoyo 2"]) profes.add(f["Profesor de apoyo 2"]);
    });
    profesores = [...profes].sort();

    const select = document.getElementById("profesor");
    profesores.forEach(p => {
      const opt = document.createElement("option");
      opt.value = p;
      opt.textContent = p;
      select.appendChild(opt);
    });

    console.log("✅ Datos cargados correctamente:", { horario: datosHorario.length });
  } catch (error) {
    console.error("❌ Error al cargar Excel:", error);
    alert("Error al cargar los datos desde el Excel.");
  }
}

// ============================================================
// BLOQUE 2 - AUSENCIAS
// ============================================================

let ausencias = [];

function registrarAusencia() {
  const profesor = document.getElementById("profesor").value;
  const desdeISO = document.getElementById("fecha-desde").value;
  const hastaISO = document.getElementById("fecha-hasta").value || desdeISO;
  const horasSel = Array.from(document.getElementById("horas").selectedOptions).map(o => o.value);

  if (!profesor || !desdeISO) {
    alert("Debes seleccionar profesor y fecha.");
    return;
  }

  const desde = formatearFechaInput(desdeISO);
  const hasta = formatearFechaInput(hastaISO);
  const nueva = { profesor, desde, hasta, horas: horasSel };

  ausencias.push(nueva);
  guardarAusencias();
  renderAusencias();
}

function guardarAusencias() {
  localStorage.setItem("ausencias", JSON.stringify(ausencias));
}

function cargarAusencias() {
  const data = localStorage.getItem("ausencias");
  if (data) {
    ausencias = JSON.parse(data);
    renderAusencias();
  }
}

function renderAusencias() {
  const lista = document.getElementById("lista-ausencias");
  lista.innerHTML = "";
  ausencias.forEach((a, i) => {
    const li = document.createElement("li");
    li.textContent = `${a.profesor} → ${a.desde}${a.hasta !== a.desde ? " - " + a.hasta : ""} (${a.horas.join(", ")})`;
    const del = document.createElement("button");
    del.textContent = "🗑";
    del.onclick = () => { ausencias.splice(i, 1); guardarAusencias(); renderAusencias(); };
    li.appendChild(del);
    lista.appendChild(li);
  });
}

function formatearFechaInput(f) {
  const [y, m, d] = f.split("-");
  return `${d}/${m}/${y}`;
}

// ============================================================
// BLOQUE 3 - GENERACIÓN DE CUADRANTE
// ============================================================

let cuadranteGenerado = [];
let computoGlobal = {};
let rotacionApoyos = ["Belén", "María Teresa", "Cris", "Pamela"];
let indiceRotacion = 0;

function generarCuadrante() {
  if (ausencias.length === 0) {
    alert("No hay ausencias registradas.");
    return;
  }

  const fechaDesde = prompt("Introduce la fecha DESDE (dd/mm/aaaa):");
  const fechaHasta = prompt("Introduce la fecha HASTA (dd/mm/aaaa):");
  if (!fechaDesde || !fechaHasta) return;

  cuadranteGenerado = [];
  computoGlobal = {};

  const desdeISO = convertirAISO(fechaDesde);
  const hastaISO = convertirAISO(fechaHasta);

  const filtradas = ausencias.filter(a =>
    convertirAISO(a.desde) <= hastaISO && convertirAISO(a.hasta) >= desdeISO
  );

  filtradas.forEach(a => {
    const rango = obtenerRangoFechas(convertirAISO(a.desde), convertirAISO(a.hasta));
    rango.forEach(fISO => {
      const dia = obtenerDiaSemana(fISO);
      const horas = a.horas.includes("dia_completo")
        ? ["9:00-10:00", "10:00-11:00", "11:30-12:30", "14:30-15:15", "15:15-16:00"]
        : a.horas;
      horas.forEach(h => {
        const clase = buscarClaseDelProfesor(a.profesor, dia, h);
        const sustituto = buscarSustituto(a.profesor, dia, h, clase);
        registrarCuadrante(fISO, a.profesor, clase, h, sustituto);
      });
    });
  });

  mostrarCuadrante();
}

function registrarCuadrante(fecha, prof, clase, hora, sustituto) {
  let bloque = cuadranteGenerado.find(d => d.fecha === fecha && d.profesorAusente === prof);
  if (!bloque) {
    bloque = { fecha, profesorAusente: prof, horas: [] };
    cuadranteGenerado.push(bloque);
  }
  bloque.horas.push({ hora, clase, sustituye: sustituto });
  if (sustituto && sustituto !== "—") computoGlobal[sustituto] = (computoGlobal[sustituto] || 0) + 1;
}

function buscarSustituto(prof, dia, hora, clase) {
  const guardia = datosGuardias.find(f => f["Día"]?.toLowerCase() === dia.toLowerCase() && f["Hora"] === hora);
  const horarioDia = datosHorario.filter(f => f["Día"]?.toLowerCase() === dia.toLowerCase() && f["Hora"] === hora);

  if (guardia) {
    const turno = Math.random() < 0.5 ? "Profesor de guardia" : "Profesor de guardia 2";
    if (guardia[turno]) return guardia[turno];
  }

  const fila = horarioDia.find(f => f["Clase"] === clase);
  if (fila) {
    const apoyos = [fila["Profesor de apoyo 1"], fila["Profesor de apoyo 2"]].filter(Boolean);
    if (apoyos.length > 0) {
      const pamela = apoyos.find(p => p.toLowerCase().includes("pamela"));
      if (pamela) return pamela;
      return elegirPorRotacion(apoyos);
    }
  }

  const externos = horarioDia.flatMap(f => [f["Profesor de apoyo 1"], f["Profesor de apoyo 2"]]).filter(Boolean);
  if (externos.length > 0) {
    const pamela = externos.find(p => p.toLowerCase().includes("pamela"));
    if (pamela) return pamela;
    return elegirPorRotacion(externos);
  }

  const pastoral = datosPastoral.find(f =>
    f["Día"]?.toLowerCase() === dia.toLowerCase() && f["Hora"] === hora && f["Disponible"] === "Sí"
  );
  if (pastoral) return pastoral["Profesor"];

  return "—";
}

function buscarClaseDelProfesor(nombre, dia, hora) {
  const fila = datosHorario.find(f =>
    f["Día"]?.toLowerCase() === dia.toLowerCase() &&
    f["Hora"]?.trim() === hora.trim() &&
    (f["Profesor de referencia"] === nombre ||
     f["Profesor de apoyo 1"] === nombre ||
     f["Profesor de apoyo 2"] === nombre)
  );
  return fila ? fila["Clase"] : "-";
}

function elegirPorRotacion(lista) {
  if (lista.length === 0) return "—";
  const candidato = rotacionApoyos[indiceRotacion % rotacionApoyos.length];
  indiceRotacion++;
  return lista.includes(candidato) ? candidato : lista[0];
}

function convertirAISO(f) {
  const [d, m, y] = f.split("/");
  return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
}
function obtenerDiaSemana(fISO) {
  const dias = ["Domingo","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado"];
  return dias[new Date(fISO).getDay()];
}
function obtenerRangoFechas(d, h) {
  const arr = [];
  let cur = new Date(d), fin = new Date(h);
  while (cur <= fin) { arr.push(cur.toISOString().slice(0,10)); cur.setDate(cur.getDate()+1); }
  return arr;
}

// ============================================================
// BLOQUE 4 - MOSTRAR Y EXPORTAR
// ============================================================

function mostrarCuadrante() {
  const seccion = document.getElementById("cuadrante");
  seccion.innerHTML = "";
  if (cuadranteGenerado.length === 0) {
    seccion.innerHTML = "<p>No hay sustituciones.</p>";
    return;
  }

  cuadranteGenerado.sort((a, b) => a.fecha.localeCompare(b.fecha)).forEach(b => {
    const fecha = new Date(b.fecha).toLocaleDateString("es-ES");
    const div = document.createElement("div");
    div.innerHTML = `
      <h3>🗓️ Sustituciones del ${fecha}</h3>
      <h4>Profesor ausente: <strong>${b.profesorAusente}</strong></h4>
      <table class="tabla-cuadrante">
        <thead><tr><th>Hora</th><th>Clase</th><th>Ausente</th><th>Sustituye</th></tr></thead>
        <tbody>
          ${b.horas.map(h =>
            `<tr><td>${h.hora}</td><td>${h.clase}</td><td>${b.profesorAusente}</td><td>${h.sustituye}</td></tr>`
          ).join("")}
        </tbody>
      </table>`;
    seccion.appendChild(div);
  });
}

function mostrarPanelAusencias() {
  document.getElementById("panel-ausencias").style.display = "block";
  document.getElementById("panel-cuadrante").style.display = "none";
}

function descargarExcel() {
  if (cuadranteGenerado.length === 0) {
    alert("No hay cuadrante generado.");
    return;
  }
  const wb = XLSX.utils.book_new();
  const hoja1 = [];
  cuadranteGenerado.forEach(b => {
    b.horas.forEach(h => hoja1.push({
      Fecha: new Date(b.fecha).toLocaleDateString("es-ES"),
      Clase: h.clase,
      "Profesor ausente": b.profesorAusente,
      "Profesor que cubre": h.sustituye
    }));
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(hoja1), "CUADRANTE_GENERADO");

  const hoja2 = Object.entries(computoGlobal).map(([prof, horas]) => ({ Profesor: prof, "Horas cubiertas": horas }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(hoja2), "CÓMPUTO_GLOBAL");

  XLSX.writeFile(wb, "Guardias_Primaria.xlsx");
}
