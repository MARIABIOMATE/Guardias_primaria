/* === CONFIGURACIÓN === */
const excelURL = "https://marbiomato.github.io/Guardias_primaria/HORARIOS_GUARDIAS.xlsx";

let workbookData = {};
let guardiaAlterna = {};
let ausencias = [];

/* === CARGA DEL ARCHIVO EXCEL DESDE GITHUB === */
async function cargarExcel() {
  try {
    const response = await fetch(excelURL);
    if (!response.ok) throw new Error("No se pudo cargar el archivo Excel desde GitHub.");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    workbook.SheetNames.forEach(name => {
      workbookData[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name]);
    });

    console.log("Archivo Excel cargado correctamente.");
  } catch (error) {
    alert("Error al cargar el archivo Excel. Verifica la URL en GitHub.");
    console.error(error);
  }
}

/* === PANEL LATERAL: ABRIR / CERRAR === */
const panel = document.getElementById("ausenciasPanel");
const toggleBtn = document.getElementById("togglePanelBtn");
const cerrarBtn = document.getElementById("cerrarPanelBtn");

toggleBtn.addEventListener("click", () => panel.classList.add("activo"));
cerrarBtn.addEventListener("click", () => panel.classList.remove("activo"));

/* === LOCALSTORAGE: GESTIÓN DE AUSENCIAS === */
const form = document.getElementById("formAusencia");
const listaAusencias = document.getElementById("listaAusencias");

function guardarEnLocalStorage() {
  localStorage.setItem("ausencias", JSON.stringify(ausencias));
}

function cargarDesdeLocalStorage() {
  const guardadas = JSON.parse(localStorage.getItem("ausencias")) || [];
  const hoy = new Date();

  // Filtrar ausencias pasadas automáticamente
  ausencias = guardadas.filter(a => new Date(a.fecha) >= hoy);
  guardarEnLocalStorage();
  renderizarAusencias();
}

form.addEventListener("submit", e => {
  e.preventDefault();

  const profesor = document.getElementById("profesorInput").value.trim();
  const fecha = document.getElementById("fechaInput").value;
  const horas = Array.from(document.getElementById("horasInput").selectedOptions).map(o => o.value);

  if (!profesor || !fecha || horas.length === 0) {
    alert("Por favor completa todos los campos.");
    return;
  }

  ausencias.push({ profesor, fecha, horas });
  guardarEnLocalStorage();
  renderizarAusencias();
  form.reset();
});

function eliminarAusencia(index) {
  ausencias.splice(index, 1);
  guardarEnLocalStorage();
  renderizarAusencias();
}

function renderizarAusencias() {
  listaAusencias.innerHTML = "";
  ausencias.forEach((a, i) => {
    const li = document.createElement("li");
    const fechaObj = new Date(a.fecha);
    const fechaTexto = `${fechaObj.getDate()}/${fechaObj.getMonth() + 1}/${String(fechaObj.getFullYear()).slice(-2)}`;
    const checkbox = `<input type="checkbox" class="chkAusencia" data-index="${i}" checked />`;
    li.innerHTML = `
      ${checkbox}
      <span><strong>${a.profesor}</strong> — ${fechaTexto} (${a.horas.join(", ")})</span>
      <button onclick="eliminarAusencia(${i})">🗑</button>
    `;
    listaAusencias.appendChild(li);
  });
}

window.addEventListener("load", () => {
  cargarExcel();
  cargarDesdeLocalStorage();
});

/* === GENERAR CUADRANTE === */
document.getElementById("generarBtn").addEventListener("click", async () => {
  const seleccionadas = Array.from(document.querySelectorAll(".chkAusencia:checked"))
    .map(chk => ausencias[parseInt(chk.dataset.index)]);

  if (seleccionadas.length === 0) {
    alert("Selecciona al menos una ausencia para generar el cuadrante.");
    return;
  }

  const resultado = [];
  seleccionadas.forEach(aus => {
    const fechaObj = new Date(aus.fecha);
    const diaSemana = obtenerDiaSemana(fechaObj);
    const horasAfectadas = aus.horas.includes("Día completo")
      ? ["9:00–10:00", "10:00–11:00", "11:30–12:30", "14:30–15:15", "15:15–16:00"]
      : aus.horas;

    horasAfectadas.forEach(hora => {
      const clases = workbookData["HORARIO"].filter(row =>
        row["Día"] === diaSemana &&
        row["Hora"] === hora &&
        (row["Profesor de referencia"] === aus.profesor ||
         row["Profesor de apoyo 1"] === aus.profesor ||
         row["Profesor de apoyo 2"] === aus.profesor)
      );

      clases.forEach(cl => {
        resultado.push({
          fecha: aus.fecha,
          dia: diaSemana,
          hora,
          clase: cl["Clase"],
          profesorAusente: aus.profesor
        });
      });
    });
  });

  generarCoberturas(resultado);
  alert("✅ Cuadrante generado correctamente.");
});

/* === FUNCIONES AUXILIARES === */
function obtenerDiaSemana(fecha) {
  const dias = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
  return dias[fecha.getDay()];
}

/* === GENERAR COBERTURAS === */
function generarCoberturas(listaClases) {
  const tbody = document.querySelector("#resultado tbody");
  tbody.innerHTML = "";
  const resumen = {};
  const ocupaciones = {};

  listaClases.sort((a, b) => new Date(a.fecha) - new Date(b.fecha));
  const fechasUnicas = [...new Set(listaClases.map(i => i.fecha))];

  fechasUnicas.forEach(fecha => {
    const fechaObj = new Date(fecha);
    const fechaTexto = `${fechaObj.getDate()}/${fechaObj.getMonth() + 1}/${String(fechaObj.getFullYear()).slice(-2)}`;
    const trDia = document.createElement("tr");
    trDia.classList.add("dia-header");
    trDia.innerHTML = `<td colspan="4">🟦 Sustituciones del ${fechaTexto}</td>`;
    tbody.appendChild(trDia);

    const clasesDelDia = listaClases.filter(i => i.fecha === fecha);
    const profesoresAusentes = [...new Set(clasesDelDia.map(i => i.profesorAusente))];

    profesoresAusentes.forEach(prof => {
      const trProf = document.createElement("tr");
      trProf.classList.add("profesor-header");
      trProf.innerHTML = `<td colspan="4">🔹 Profesor ausente: <strong>${prof}</strong></td>`;
      tbody.appendChild(trProf);

      const clasesProf = clasesDelDia.filter(c => c.profesorAusente === prof);

      clasesProf.forEach(item => {
        const dia = item.dia;
        const hora = item.hora;
        const clase = item.clase;
        const profAusente = item.profesorAusente;
        let profesorCubre = "";

        const claveHora = `${dia}_${hora}`;
        if (!ocupaciones[claveHora]) ocupaciones[claveHora] = new Set();
        const profesorLibre = nombre => nombre && !ocupaciones[claveHora].has(nombre);

        // 1️⃣ Guardia (alternando)
        const guardias = workbookData["GUARDIAS"].filter(g => g["Día"] === dia && g["Hora"] === hora);
        if (guardias.length > 0) {
          const alterna = guardiaAlterna[hora] || false;
          const posible = !alterna
            ? guardias[0]["Profesor de guardia"]
            : guardias[0]["Profesor de guardia 2"] || guardias[0]["Profesor de guardia"];
          if (profesorLibre(posible)) {
            profesorCubre = posible;
            guardiaAlterna[hora] = !alterna;
          }
        }

        // 2️⃣ Apoyo interno
        if (!profesorCubre) {
          const filaHorario = workbookData["HORARIO"].find(h =>
            h["Día"] === dia && h["Hora"] === hora && h["Clase"] === clase
          );
          if (filaHorario) {
            const posibles = [filaHorario["Profesor de apoyo 1"], filaHorario["Profesor de apoyo 2"]];
            profesorCubre = posibles.find(p => p !== profAusente && profesorLibre(p));
          }
        }

        // 3️⃣ Apoyo externo
        if (!profesorCubre) {
          const apoyos = workbookData["HORARIO"].filter(h =>
            h["Día"] === dia && h["Hora"] === hora &&
            h["Clase"] !== clase &&
            profesorLibre(h["Profesor de apoyo 1"]) && h["Profesor de apoyo 1"] !== profAusente
          );
          if (apoyos.length > 0) profesorCubre = apoyos[0]["Profesor de apoyo 1"];
        }

        // 4️⃣ Pastoral / Convivencia
        if (!profesorCubre) {
          const pastoral = workbookData["PASTORAL_CONVIVENCIA"].find(p =>
            p["Día"] === dia && p["Hora"] === hora && p["Disponible"] === "Sí"
          );
          if (pastoral && profesorLibre(pastoral["Profesor"])) profesorCubre = pastoral["Profesor"];
        }

        if (!profesorCubre) profesorCubre = "No disponible";

        ocupaciones[claveHora].add(profesorCubre);
        if (!resumen[profesorCubre]) resumen[profesorCubre] = { guardias: 0, horas: 0 };
        resumen[profesorCubre].guardias += 1;
        resumen[profesorCubre].horas += 1;

        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td>${hora}</td>
          <td>${clase}</td>
          <td>${profAusente}</td>
          <td>${profesorCubre}</td>
        `;
        tbody.appendChild(tr);
      });
    });
  });

  document.getElementById("descargarBtn").style.display = "block";
  document.getElementById("descargarBtn").onclick = () => exportarCuadrante(resumen);
}

/* === EXPORTAR EXCEL === */
function exportarCuadrante(resumen) {
  const wsCuadrante = XLSX.utils.table_to_sheet(document.getElementById("resultado"));

  const datosResumen = [["Profesor", "Horas de guardia", "Horas sustituidas", "Guardias acumuladas"]];
  Object.keys(resumen).forEach(p => {
    datosResumen.push([p, resumen[p].horas, 0, resumen[p].guardias]); // “0” placeholder para horas sustituidas
  });

  const wsGlobal = XLSX.utils.aoa_to_sheet(datosResumen);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, wsCuadrante, "CUADRANTE_GENERADO");
  XLSX.utils.book_append_sheet(wb, wsGlobal, "CÓMPUTO_GLOBAL");
  XLSX.writeFile(wb, "Cuadrante_Guardias_Actualizado.xlsx");
}
