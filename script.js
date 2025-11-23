// ============================================================
// BLOQUE 1 - CONFIGURACIÓN Y CARGA DE DATOS DESDE EXCEL
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

    const hojaHorario = XLSX.utils.sheet_to_json(workbook.Sheets["HORARIO"]);
    const hojaGuardias = XLSX.utils.sheet_to_json(workbook.Sheets["GUARDIAS"]);
    const hojaPastoral = XLSX.utils.sheet_to_json(workbook.Sheets["PASTORAL_CONVIVENCIA"]);

    datosHorario = hojaHorario;
    datosGuardias = hojaGuardias;
    datosPastoral = hojaPastoral;

    // Lista única de profesores (de horario y guardias)
    const nombres = new Set();
    hojaHorario.forEach(f => {
      if (f["Profesor de referencia"]) nombres.add(f["Profesor de referencia"]);
      if (f["Profesor de apoyo 1"]) nombres.add(f["Profesor de apoyo 1"]);
      if (f["Profesor de apoyo 2"]) nombres.add(f["Profesor de apoyo 2"]);
    });
    hojaGuardias.forEach(f => {
      if (f["Profesor de guardia"]) nombres.add(f["Profesor de guardia"]);
      if (f["Profesor de guardia 2"]) nombres.add(f["Profesor de guardia 2"]);
    });
    profesores = Array.from(nombres).sort();

    const select = document.getElementById("profesor");
    select.innerHTML = '<option value="">-- Selecciona profesor --</option>';
    profesores.forEach(p => {
      const opt = document.createElement("option");
      opt.value = p;
      opt.textContent = p;
      select.appendChild(opt);
    });

    console.log("Datos cargados:", { datosHorario, datosGuardias, datosPastoral });
  } catch (err) {
    alert("Error al cargar el Excel remoto.");
    console.error(err);
  }
}
// ============================================================
// BLOQUE 2 - GESTIÓN DE AUSENCIAS (REGISTRO Y LOCALSTORAGE)
// ============================================================

let ausencias = [];

// Registrar una nueva ausencia
function registrarAusencia() {
  const profesor = document.getElementById("profesor").value;
  const desde = document.getElementById("fecha-desde").value;
  const hasta = document.getElementById("fecha-hasta").value || desde;
  const horasSel = Array.from(document.getElementById("horas").selectedOptions).map(o => o.value);

  if (!profesor || !desde) {
    alert("Debes seleccionar al menos un profesor y una fecha.");
    return;
  }

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
// BLOQUE 3 (CORREGIDO) - GENERACIÓN DE CUADRANTE Y SUSTITUCIONES
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
  const guardia = datosGuardias.find(f => f["Día"] === dia && f["Hora"] === hora);
  const horarioDia = datosHorario.filter(f => f["Día"] === dia && f["Hora"] === hora);

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

  // 3️⃣ Apoyo externo (de otras clases)
  const apoyosExternos = horarioDia.flatMap(f => [f["Profesor de apoyo 1"], f["Profesor de apoyo 2"]]).filter(Boolean);
  if (apoyosExternos.length > 0) {
    const pamela = apoyosExternos.find(p => p && p.toLowerCase().includes("pamela"));
    if (pamela) return pamela;
    return elegirPorRotacion(apoyosExternos);
  }

  // 4️⃣ Pastoral / convivencia disponibles
  const pastoral = datosPastoral.find(f => f["Día"] === dia && f["Hora"] === hora && f["Disponible"] === "Sí");
  if (pastoral) return pastoral["Profesor"];

  return "—";
}

// Buscar la clase del profesor ausente
function buscarClaseDelProfesor(nombre, dia, hora) {
  const fila = datosHorario.find(f =>
    f["Día"] === dia &&
    f["Hora"] === hora &&
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
  const dias = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
  const d = new Date(fechaISO);
  return dias[d.getDay()];
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

function mostrarCuadrante() {
  document.getElementById("panel-ausencias").classList.add("oculto");
  document.getElementById("vista-cuadrante").classList.remove("oculto");

  const contenedor = document.getElementById("contenedor-cuadrante");
  contenedor.innerHTML = "";

  const diasUnicos = [...new Set(cuadranteGenerado.map(d => d.fecha))];

  diasUnicos.forEach(dia => {
    const bloqueDia = document.createElement("div");
    bloqueDia.className = "bloque-dia";
    bloqueDia.innerHTML = `<h3>Sustituciones del ${formatearFecha(dia)}</h3>`;

    const ausentesDia = cuadranteGenerado.filter(d => d.fecha === dia);

    ausentesDia.forEach(a => {
      const bloqueProf = document.createElement("div");
      bloqueProf.className = "bloque-profesor";
      bloqueProf.innerHTML = `<h4>Profesor ausente: ${a.profesorAusente}</h4>`;

      const tabla = document.createElement("table");
      tabla.className = "tabla-cuadrante";
      tabla.innerHTML = `
        <thead>
          <tr>
            <th>Hora</th>
            <th>Clase</th>
            <th>Profesor ausente</th>
            <th>Profesor que cubre</th>
          </tr>
        </thead>
        <tbody>
          ${a.horas
            .map(
              h => `
              <tr>
                <td>${h.hora}</td>
                <td>${h.clase}</td>
                <td>${a.profesorAusente}</td>
                <td>${h.sustituye}</td>
              </tr>`
            )
            .join("")}
        </tbody>
      `;
      bloqueProf.appendChild(tabla);
      bloqueDia.appendChild(bloqueProf);
    });

    contenedor.appendChild(bloqueDia);
  });
}

function mostrarPanelAusencias() {
  document.getElementById("vista-cuadrante").classList.add("oculto");
  document.getElementById("panel-ausencias").classList.remove("oculto");
}

function formatearFecha(fechaStr) {
  const [y, m, d] = fechaStr.split("-");
  return `${d}/${m}/${y.slice(2)}`;
}

// ============================================================
// DESCARGAR EXCEL FINAL CON DOS HOJAS
// ============================================================

function descargarExcel() {
  if (cuadranteGenerado.length === 0) {
    alert("Primero genera el cuadrante.");
    return;
  }

  const hoja1 = [];
  cuadranteGenerado.forEach(d =>
    d.horas.forEach(h =>
      hoja1.push({
        Fecha: formatearFecha(d.fecha),
        Hora: h.hora,
        Clase: h.clase,
        "Profesor ausente": d.profesorAusente,
        "Profesor que cubre": h.sustituye
      })
    )
  );

  const hoja2 = Object.keys(computoGlobal).map(p => ({
    Profesor: p,
    "Horas cubiertas": computoGlobal[p]
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(hoja1), "CUADRANTE_GENERADO");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(hoja2), "CÓMPUTO_GLOBAL");

  XLSX.writeFile(wb, "Guardias_Primaria.xlsx");
}
