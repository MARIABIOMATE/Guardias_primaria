// === CONFIGURACIÓN ===
// Enlace a tu hoja de Google Sheets exportada como Excel (.xlsx)
const EXCEL_URL = "https://docs.google.com/spreadsheets/d/1htrGqexP4ZDPmf6Q9JxP0CaGcNXERF7-/export?format=xlsx";

let workbookData = {};
let ausencias = [];
let guardiaAlterna = {};

// === CARGA AUTOMÁTICA DEL EXCEL DESDE GOOGLE SHEETS ===
async function cargarExcel() {
  try {
    const response = await fetch(EXCEL_URL);
    if (!response.ok) throw new Error("No se pudo acceder al archivo.");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    workbook.SheetNames.forEach(name => {
      workbookData[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name]);
    });

    alert("✅ Archivo Excel cargado correctamente desde Google Sheets.");
  } catch (error) {
    alert("⚠️ Error al cargar el archivo Excel. Verifica la URL o permisos de la hoja.");
    console.error(error);
  }
}

// === GESTIÓN DE PANEL LATERAL ===
document.getElementById("gestionBtn").addEventListener("click", () => {
  document.getElementById("panelAusencias").classList.add("abierto");
});

document.getElementById("cerrarPanel").addEventListener("click", () => {
  document.getElementById("panelAusencias").classList.remove("abierto");
});

// === AÑADIR AUSENCIAS ===
document.getElementById("agregarAusenciaBtn").addEventListener("click", () => {
  const cont = document.getElementById("ausenciasContainer");
  const div = document.createElement("div");
  div.classList.add("bloque-ausencia");
  div.innerHTML = `
    <h3>Ausencia</h3>
    <label>Profesor ausente:</label>
    <input type="text" class="profesorInput" placeholder="Nombre del profesor">

    <label>Desde:</label>
    <input type="date" class="desdeFecha">

    <label>Hasta:</label>
    <input type="date" class="hastaFecha">

    <label>Horas afectadas:</label>
    <select class="horasSelect" multiple>
      <option value="Día completo">Día completo</option>
      <option value="9:00–10:00">9:00–10:00</option>
      <option value="10:00–11:00">10:00–11:00</option>
      <option value="11:30–12:30">11:30–12:30</option>
      <option value="14:30–15:15">14:30–15:15</option>
      <option value="15:15–16:00">15:15–16:00</option>
    </select>
    <hr>
  `;
  cont.appendChild(div);
});

// === RECOGER AUSENCIAS ===
function recogerAusencias() {
  ausencias = [];
  document.querySelectorAll(".bloque-ausencia").forEach(div => {
    const profesor = div.querySelector(".profesorInput").value.trim();
    const desde = div.querySelector(".desdeFecha").value;
    const hasta = div.querySelector(".hastaFecha").value || desde;
    const horasSel = Array.from(div.querySelector(".horasSelect").selectedOptions).map(o => o.value);
    if (profesor && desde) {
      ausencias.push({ profesor, desde, hasta, horasSel });
    }
  });
  console.log("Ausencias registradas:", ausencias);
}

// === FUNCIONES DE UTILIDAD ===
function obtenerDiaSemana(fecha) {
  const dias = ["Domingo","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado"];
  return dias[new Date(fecha).getDay()];
}

function rangoFechas(desde, hasta) {
  const fechas = [];
  let actual = new Date(desde);
  const fin = new Date(hasta);
  while (actual <= fin) {
    fechas.push(new Date(actual));
    actual.setDate(actual.getDate() + 1);
  }
  return fechas;
}

// === GENERAR CUADRANTE ===
function generarCoberturas(listaClases) {
  const tbody = document.querySelector("#resultado tbody");
  tbody.innerHTML = "";
  const resumen = {};
  const ocupaciones = {};

  const fechasUnicas = [...new Set(listaClases.map(i => i.fecha))];
  fechasUnicas.forEach(fecha => {
    const trDia = document.createElement("tr");
    trDia.innerHTML = `
      <td colspan="4" style="background:#004080;color:white;font-weight:bold;text-align:left">
      📅 Sustituciones del ${fecha}
      </td>`;
    tbody.appendChild(trDia);

    const profesores = [...new Set(listaClases.filter(i => i.fecha === fecha).map(x => x.profesorAusente))];
    profesores.forEach(prof => {
      const trProf = document.createElement("tr");
      trProf.innerHTML = `<td colspan="4" style="background:#e8eefc;font-weight:bold;text-align:left">
        👩‍🏫 Profesor ausente: ${prof}
      </td>`;
      tbody.appendChild(trProf);

      const clases = listaClases.filter(c => c.fecha === fecha && c.profesorAusente === prof);
      clases.forEach(item => {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${item.hora}</td>
          <td>${item.clase}</td>
          <td>${item.profesorAusente}</td>
          <td>${item.profesorCubre || "No disponible"}</td>
        `;
        tbody.appendChild(tr);
      });
    });
  });
}

// === PROCESAR AUSENCIAS Y GENERAR LISTADO ===
document.addEventListener("DOMContentLoaded", async () => {
  await cargarExcel();
  document.getElementById("agregarAusenciaBtn").click(); // crea el primer bloque
});

// Podrás ampliar más adelante con botón "Generar cuadrante" si lo deseas.
