// ============================================================
// CONFIGURACIÓN INICIAL
// ============================================================
const URL_EXCEL =
  "https://raw.githubusercontent.com/MARIABIOMATE/Guardias_primaria/refs/heads/main/HORARIOS_GUARDIAS.xlsx";

let profesores = [];
let ausencias = [];
let cuadranteGenerado = [];
let computoGlobal = {};
let repartoRotativo = { Belén: 0, "María Teresa": 0, Cris: 0 };

// ============================================================
// INICIALIZACIÓN
// ============================================================
window.addEventListener("DOMContentLoaded", () => {
  cargarExcel();
  cargarAusencias();
  document
    .getElementById("add-ausencia")
    .addEventListener("click", registrarAusencia);
  document
    .getElementById("generar-cuadrante")
    .addEventListener("click", generarCuadrante);
  document
    .getElementById("descargar-excel")
    .addEventListener("click", descargarExcel);
  document
    .getElementById("volver")
    .addEventListener("click", mostrarPanelAusencias);
});

// ============================================================
// CARGA DE EXCEL REMOTO
// ============================================================
async function cargarExcel() {
  try {
    const response = await fetch(URL_EXCEL);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    const hojaHorario = XLSX.utils.sheet_to_json(workbook.Sheets["HORARIO"]);
    profesores = [
      ...new Set(hojaHorario.map((fila) => fila["Profesor"] || fila["Profesor de referencia"])),
    ].sort();

    const select = document.getElementById("profesor");
    profesores.forEach((p) => {
      const opt = document.createElement("option");
      opt.value = p;
      opt.textContent = p;
      select.appendChild(opt);
    });
  } catch (err) {
    alert("Error al cargar el Excel remoto. Revisa la URL.");
    console.error(err);
  }
}

// ============================================================
// GESTIÓN DE AUSENCIAS
// ============================================================
function registrarAusencia() {
  const profesor = document.getElementById("profesor").value;
  const desde = document.getElementById("fecha-desde").value;
  const hasta = document.getElementById("fecha-hasta").value || desde;
  const horasSel = Array.from(document.getElementById("horas").selectedOptions).map(
    (opt) => opt.value
  );

  if (!profesor || !desde) {
    alert("Debe seleccionar profesor y fecha.");
    return;
  }

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
    const btn = document.createElement("button");
    btn.textContent = "🗑";
    btn.onclick = () => {
      ausencias.splice(i, 1);
      guardarAusencias();
      renderAusencias();
    };
    li.appendChild(btn);
    lista.appendChild(li);
  });
}

// ============================================================
// GENERACIÓN DE CUADRANTE
// ============================================================
function generarCuadrante() {
  cuadranteGenerado = [];
  computoGlobal = {};

  ausencias.forEach((a) => {
    const fechas = obtenerRangoFechas(a.desde, a.hasta);
    fechas.forEach((fecha) => {
      const diaObj = {
        fecha,
        profesorAusente: a.profesor,
        horas: [],
      };

      (a.horas.includes("dia_completo")
        ? ["9:00-10:00", "10:00-11:00", "11:30-12:30", "14:30-15:15", "15:15-16:00"]
        : a.horas
      ).forEach((hora) => {
        const sustituto = asignarSustituto(a.profesor);
        diaObj.horas.push({ hora, clase: "-", sustituye: sustituto });
        computoGlobal[sustituto] = (computoGlobal[sustituto] || 0) + 1;
      });

      cuadranteGenerado.push(diaObj);
    });
  });

  mostrarCuadrante();
}

// ============================================================
// LÓGICA DE ASIGNACIÓN (ROTATIVA)
// ============================================================
function asignarSustituto(profAusente) {
  // Simulación simple: elige el profesor con menos guardias acumuladas
  const candidatos = ["Belén", "María Teresa", "Cris"];
  const siguiente = candidatos.reduce((a, b) =>
    repartoRotativo[a] <= repartoRotativo[b] ? a : b
  );
  repartoRotativo[siguiente]++;
  return siguiente;
}

// ============================================================
// UTILIDADES
// ============================================================
function obtenerRangoFechas(desde, hasta) {
  const fechas = [];
  let current = new Date(desde);
  const end = new Date(hasta);
  while (current <= end) {
    const d = current.toISOString().slice(0, 10);
    fechas.push(d);
    current.setDate(current.getDate() + 1);
  }
  return fechas;
}

// ============================================================
// MOSTRAR CUADRANTE EN PANTALLA
// ============================================================
function mostrarCuadrante() {
  document.getElementById("panel-ausencias").classList.add("oculto");
  document.getElementById("vista-cuadrante").classList.remove("oculto");

  const contenedor = document.getElementById("contenedor-cuadrante");
  contenedor.innerHTML = "";

  const diasUnicos = [...new Set(cuadranteGenerado.map((d) => d.fecha))];

  diasUnicos.forEach((dia) => {
    const bloqueDia = document.createElement("div");
    bloqueDia.className = "bloque-dia";
    bloqueDia.innerHTML = `<h3>Sustituciones del ${formatearFecha(dia)}</h3>`;

    const ausentes = cuadranteGenerado.filter((d) => d.fecha === dia);
    ausentes.forEach((a) => {
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
              (h) => `
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
// DESCARGAR EXCEL FINAL
// ============================================================
function descargarExcel() {
  if (!cuadranteGenerado.length) {
    alert("Primero genera el cuadrante.");
    return;
  }

  const hoja1 = [];
  cuadranteGenerado.forEach((d) =>
    d.horas.forEach((h) =>
      hoja1.push({
        Fecha: formatearFecha(d.fecha),
        Hora: h.hora,
        Clase: h.clase,
        "Profesor ausente": d.profesorAusente,
        "Profesor que cubre": h.sustituye,
      })
    )
  );

  const hoja2 = Object.keys(computoGlobal).map((p) => ({
    Profesor: p,
    "Horas cubiertas": computoGlobal[p],
  }));

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(hoja1), "CUADRANTE_GENERADO");
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(hoja2), "CÓMPUTO_GLOBAL");

  XLSX.writeFile(wb, "Guardias_Primaria.xlsx");
}
