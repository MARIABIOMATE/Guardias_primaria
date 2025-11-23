// ============================================================
// BLOQUE 1 - CONFIGURACIÓN Y CARGA DE DATOS DESDE EXCEL (CORREGIDO)
// ============================================================

const URL_EXCEL = "https://raw.githubusercontent.com/MARIABIOMATE/Guardias_primaria/main/HORARIOS_GUARDIAS.xlsx";

let datosHorario = [];
let datosGuardias = [];
let datosPastoral = [];
let profesores = [];

// Función principal al cargar la página
window.addEventListener("DOMContentLoaded", async () => {
  await cargarExcel();
  cargarAusencias();
  document.getElementById("add-ausencia").addEventListener("click", registrarAusencia);
  document.getElementById("generar-cuadrante").addEventListener("click", generarCuadrante);
  document.getElementById("descargar-excel").addEventListener("click", descargarExcel);
  document.getElementById("volver").addEventListener("click", mostrarPanelAusencias);
});

// Cargar y procesar las hojas del Excel
async function cargarExcel() {
  try {
    const response = await fetch(URL_EXCEL);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    // Hojas esperadas
    const hojaHorario = workbook.Sheets["HORARIO"];
    const hojaGuardias = workbook.Sheets["GUARDIAS"];
    const hojaPastoral = workbook.Sheets["PASTORAL_CONVIVENCIA"];

    // Convertir a JSON
    datosHorario = XLSX.utils.sheet_to_json(hojaHorario);
    datosGuardias = XLSX.utils.sheet_to_json(hojaGuardias);
    datosPastoral = XLSX.utils.sheet_to_json(hojaPastoral);

    // Extraer lista única de profesores
    const profes = new Set();
    datosHorario.forEach(f => {
      if (f["Profesor de referencia"]) profes.add(f["Profesor de referencia"]);
      if (f["Profesor de apoyo 1"]) profes.add(f["Profesor de apoyo 1"]);
      if (f["Profesor de apoyo 2"]) profes.add(f["Profesor de apoyo 2"]);
    });
    profesores = [...profes].sort();

    // Llenar el selector de profesores
    const select = document.getElementById("profesor");
    profesores.forEach(p => {
      const opt = document.createElement("option");
      opt.value = p;
      opt.textContent = p;
      select.appendChild(opt);
    });

    console.log("✅ Datos cargados correctamente:", {
      horario: datosHorario.length,
      guardias: datosGuardias.length,
      pastoral: datosPastoral.length
    });

  } catch (error) {
    console.error("❌ Error al cargar el Excel:", error);
    alert("No se pudieron cargar los datos desde el archivo Excel.");
  }
}

// ============================================================
// BLOQUE 2 - GESTIÓN DE AUSENCIAS (REGISTRO Y LOCALSTORAGE)
// ============================================================

let ausencias = [];

// Registrar una nueva ausencia
function registrarAusencia() {
  const profesor = document.getElementById("profesor").value;
  const desdeISO = document.getElementById("fecha-desde").value;
  const hastaISO = document.getElementById("fecha-hasta").value || desdeISO;
  const horasSel = Array.from(document.getElementById("horas").selectedOptions).map(o => o.value);

  if (!profesor || !desdeISO) {
    alert("Debes seleccionar al menos un profesor y una fecha.");
    return;
  }

  const desde = formatearFechaInput(desdeISO);
  const hasta = formatearFechaInput(hastaISO);

  const nuevaAusencia = { profesor, desde, hasta, horas: horasSel };
  ausencias.push(nuevaAusencia);
  guardarAusencias();
  renderAusencias();
}

// Guardar en localStorage
function guardarAusencias() {
  localStorage.setItem("ausencias", JSON.stringify(ausencias));
}

// Cargar ausencias almacenadas
function cargarAusencias() {
  const data = localStorage.getItem("ausencias");
  if (data) {
    ausencias = JSON.parse(data);
    renderAusencias();
  }
}

// Mostrar la lista de ausencias en pantalla
function renderAusencias() {
  const lista = document.getElementById("lista-ausencias");
  lista.innerHTML = "";

  ausencias.forEach((a, i) => {
    const li = document.createElement("li");
    const texto = `${a.profesor} → ${a.desde}${a.hasta !== a.desde ? " - " + a.hasta : ""} (${a.horas.join(", ")})`;
    li.textContent = texto;

    const btnBorrar = document.createElement("button");
    btnBorrar.textContent = "🗑";
    btnBorrar.onclick = () => {
      ausencias.splice(i, 1);
      guardarAusencias();
      renderAusencias();
    };

    li.appendChild(btnBorrar);
    lista.appendChild(li);
  });
}

// ============================================================
// FUNCIÓN AUXILIAR PARA FORMATEAR FECHAS ISO → dd/mm/aaaa
// ============================================================

function formatearFechaInput(fechaISO) {
  if (!fechaISO) return "";
  const [y, m, d] = fechaISO.split("-");
  return `${d}/${m}/${y}`;
}
// ============================================================
// BLOQUE 3 - GENERACIÓN DE CUADRANTE Y ASIGNACIÓN DE SUSTITUTOS
// ============================================================

let cuadranteGenerado = [];
let computoGlobal = {};
let rotacionApoyos = ["Belén", "María Teresa", "Cris", "Pamela"];
let indiceRotacion = 0;

// Generar cuadrante (filtrado por rango de fechas dd/mm/aaaa)
function generarCuadrante() {
  if (ausencias.length === 0) {
    alert("No hay ausencias registradas.");
    return;
  }

  const fechaDesde = prompt("Introduce la fecha DESDE (dd/mm/aaaa) para generar el cuadrante:");
  const fechaHasta = prompt("Introduce la fecha HASTA (dd/mm/aaaa) para generar el cuadrante:");
  if (!fechaDesde || !fechaHasta) {
    alert("Debes introducir un rango de fechas.");
    return;
  }

  cuadranteGenerado = [];
  computoGlobal = {};

  const desdeISO = convertirAISO(fechaDesde);
  const hastaISO = convertirAISO(fechaHasta);
  const ausenciasFiltradas = ausencias.filter(a => convertirAISO(a.desde) <= hastaISO && convertirAISO(a.hasta) >= desdeISO);

  ausenciasFiltradas.forEach(a => {
    const rango = obtenerRangoFechas(convertirAISO(a.desde), convertirAISO(a.hasta));
    rango.forEach(fechaISO => {
      const diaTexto = obtenerDiaSemana(fechaISO);
      (a.horas.includes("dia_completo")
        ? ["9:00-10:00", "10:00-11:00", "11:30-12:30", "14:30-15:15", "15:15-16:00"]
        : a.horas
      ).forEach(hora => {
        const claseAfectada = buscarClaseDelProfesor(a.profesor, diaTexto, hora);
        const sustituto = buscarSustituto(a.profesor, diaTexto, hora, claseAfectada);
        registrarCuadrante(fechaISO, a.profesor, claseAfectada, hora, sustituto);
      });
    });
  });

  mostrarCuadrante();
}

// Registrar sustitución
function registrarCuadrante(fecha, profesorAusente, clase, hora, sustituto) {
  let bloque = cuadranteGenerado.find(d => d.fecha === fecha && d.profesorAusente === profesorAusente);
  if (!bloque) {
    bloque = { fecha, profesorAusente, horas: [] };
    cuadranteGenerado.push(bloque);
  }
  bloque.horas.push({ hora, clase, sustituye: sustituto });
  if (sustituto && sustituto !== "—") computoGlobal[sustituto] = (computoGlobal[sustituto] || 0) + 1;
}

// Buscar sustituto según jerarquía
function buscarSustituto(profAusente, dia, hora, clase) {
  const guardia = datosGuardias.find(f => f["Día"]?.trim().toLowerCase() === dia.trim().toLowerCase() && f["Hora"] === hora);
  const horarioDia = datosHorario.filter(f => f["Día"]?.trim().toLowerCase() === dia.trim().toLowerCase() && f["Hora"] === hora);

  // 1️⃣ Guardia 1 ↔ Guardia 2 (alternando)
  if (guardia) {
    const turno = Math.random() < 0.5 ? "Profesor de guardia" : "Profesor de guardia 2";
    if (guardia[turno]) return guardia[turno];
  }

  // 2️⃣ Apoyo interno
  const filaClase = horarioDia.find(f => f["Clase"] === clase);
  if (filaClase) {
    const apoyos = [filaClase["Profesor de apoyo 1"], filaClase["Profesor de apoyo 2"]].filter(Boolean);
    if (apoyos.length > 0) {
      const pamela = apoyos.find(p => p && p.toLowerCase().includes("pamela"));
      if (pamela) return pamela;
      return elegirPorRotacion(apoyos);
    }
  }

  // 3️⃣ Apoyo externo
  const apoyosExternos = horarioDia.flatMap(f => [f["Profesor de apoyo 1"], f["Profesor de apoyo 2"]]).filter(Boolean);
  if (apoyosExternos.length > 0) {
    const pamela = apoyosExternos.find(p => p && p.toLowerCase().includes("pamela"));
    if (pamela) return pamela;
    return elegirPorRotacion(apoyosExternos);
  }

  // 4️⃣ Pastoral / convivencia disponibles
  const pastoral = datosPastoral.find(f => f["Día"]?.trim().toLowerCase() === dia.trim().toLowerCase() && f["Hora"] === hora && f["Disponible"] === "Sí");
  if (pastoral) return pastoral["Profesor"];

  return "—";
}

// Buscar clase del profesor ausente
function buscarClaseDelProfesor(nombre, dia, hora) {
  const fila = datosHorario.find(f =>
    f["Día"]?.trim().toLowerCase() === dia.trim().toLowerCase() &&
    f["Hora"]?.trim() === hora.trim() &&
    (f["Profesor de referencia"] === nombre ||
     f["Profesor de apoyo 1"] === nombre ||
     f["Profesor de apoyo 2"] === nombre)
  );
  return fila ? fila["Clase"] : "-";
}

// Rotación entre apoyos
function elegirPorRotacion(lista) {
  if (lista.length === 0) return "—";
  for (let i = 0; i < lista.length; i++) {
    const candidato = rotacionApoyos[indiceRotacion % rotacionApoyos.length];
    if (lista.includes(candidato)) {
      indiceRotacion++;
      return candidato;
    }
  }
  indiceRotacion++;
  return lista[0];
}

// =======================================
// UTILIDADES DE FECHAS
// =======================================
function convertirAISO(fechaDDMM) {
  const [d, m, y] = fechaDDMM.split("/");
  return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
}

function obtenerDiaSemana(fechaISO) {
  const diasES = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
  const d = new Date(fechaISO);
  const nombre = diasES[d.getDay()];
  return nombre.trim().charAt(0).toUpperCase() + nombre.trim().slice(1).toLowerCase();
}

function obtenerRangoFechas(desde, hasta) {
  const arr = [];
  let cur = new Date(desde);
  const fin = new Date(hasta);
  while (cur <= fin) {
    arr.push(cur.toISOString().slice(0, 10));
    cur.setDate(cur.getDate() + 1);
  }
  return arr;
}
// ============================================================
// BLOQUE 4 - MOSTRAR CUADRANTE Y EXPORTAR A EXCEL
// ============================================================

// Mostrar cuadrante en pantalla
function mostrarCuadrante() {
  const seccion = document.getElementById("cuadrante");
  seccion.innerHTML = "";

  if (cuadranteGenerado.length === 0) {
    seccion.innerHTML = "<p>No hay sustituciones para el rango seleccionado.</p>";
    return;
  }

  const ordenado = cuadranteGenerado.sort((a, b) => a.fecha.localeCompare(b.fecha));

  ordenado.forEach(bloque => {
    const fecha = new Date(bloque.fecha);
    const fechaStr = fecha.toLocaleDateString("es-ES").replace(/\//g, "/");
    const div = document.createElement("div");
    div.classList.add("dia-cuadrante");

    div.innerHTML = `
      <h3>🗓️ Sustituciones del ${fechaStr}</h3>
      <h4>Profesor ausente: <strong>${bloque.profesorAusente}</strong></h4>
      <table class="tabla-cuadrante">
        <thead>
          <tr>
            <th>Hora</th>
            <th>Clase</th>
            <th>Profesor ausente</th>
            <th>Profesor que cubre</th>
          </tr>
        </thead>
        <tbody>
          ${bloque.horas.map(h =>
            `<tr>
              <td>${h.hora}</td>
              <td>${h.clase || "-"}</td>
              <td>${bloque.profesorAusente}</td>
              <td>${h.sustituye || "—"}</td>
            </tr>`
          ).join("")}
        </tbody>
      </table>
    `;
    seccion.appendChild(div);
  });
}

// Volver al panel de gestión
function mostrarPanelAusencias() {
  document.getElementById("panel-ausencias").style.display = "block";
  document.getElementById("panel-cuadrante").style.display = "none";
}

// Descargar Excel con las dos hojas
function descargarExcel() {
  if (cuadranteGenerado.length === 0) {
    alert("No hay cuadrante generado para exportar.");
    return;
  }

  const wb = XLSX.utils.book_new();

  // CUADRANTE_GENERADO
  const hoja1 = [];
  cuadranteGenerado.forEach(b => {
    b.horas.forEach(h => {
      hoja1.push({
        Fecha: new Date(b.fecha).toLocaleDateString("es-ES"),
        Clase: h.clase,
        "Profesor ausente": b.profesorAusente,
        "Profesor que cubre": h.sustituye
      });
    });
  });
  const ws1 = XLSX.utils.json_to_sheet(hoja1);
  XLSX.utils.book_append_sheet(wb, ws1, "CUADRANTE_GENERADO");

  // CÓMPUTO_GLOBAL
  const hoja2 = Object.entries(computoGlobal).map(([prof, horas]) => ({
    Profesor: prof,
    "Horas cubiertas": horas
  }));
  const ws2 = XLSX.utils.json_to_sheet(hoja2);
  XLSX.utils.book_append_sheet(wb, ws2, "CÓMPUTO_GLOBAL");

  XLSX.writeFile(wb, "Guardias_Primaria.xlsx");
}
