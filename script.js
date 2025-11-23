// === CONFIGURACIÓN ===
const EXCEL_URL = "https://raw.githubusercontent.com/MARIABIOMATE/Guardias_primaria/main/HORARIOS_GUARDIAS.xlsx";

let workbookData = {};
let ausencias = [];
let guardiaAlterna = {};

// === CARGA AUTOMÁTICA DEL EXCEL DESDE GITHUB ===
async function cargarExcel() {
  try {
    const response = await fetch(EXCEL_URL);
    if (!response.ok) throw new Error("No se pudo acceder al archivo.");
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    workbook.SheetNames.forEach(name => {
      workbookData[name] = XLSX.utils.sheet_to_json(workbook.Sheets[name]);
    });

    console.log("✅ Excel cargado correctamente desde GitHub.");
    document.getElementById('opciones').style.display = 'block';
    inicializarAusencia();
  } catch (err) {
    alert("⚠️ Error al cargar el archivo Excel. Verifica la URL en GitHub.");
    console.error(err);
  }
}

// === FUNCIÓN: Crear bloque de ausencia ===
function inicializarAusencia() {
  const container = document.getElementById('ausenciasContainer');
  const div = document.createElement('div');
  div.classList.add('ausencia-bloque');
  div.innerHTML = `
    <h3>Ausencia</h3>
    <label>Profesor ausente:</label>
    <select class="profesorSelect"></select>

    <label>Desde:</label>
    <input type="date" class="desdeFecha">

    <label>Hasta (opcional):</label>
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
  container.appendChild(div);

  const profesSet = new Set();
  workbookData["HORARIO"].forEach(row => {
    ["Profesor de referencia", "Profesor de apoyo 1", "Profesor de apoyo 2"].forEach(k => {
      if (row[k]) profesSet.add(row[k]);
    });
  });

  const select = div.querySelector(".profesorSelect");
  profesSet.forEach(p => {
    const opt = document.createElement("option");
    opt.textContent = p;
    select.appendChild(opt);
  });
}

// === BLOQUE 2: Generación del cuadrante ===
function obtenerDiaSemana(fecha) {
  const dias = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
  return dias[fecha.getDay()];
}

function rangoFechas(desde, hasta) {
  const fechas = [];
  let actual = new Date(desde);
  while (actual <= hasta) {
    fechas.push(new Date(actual));
    actual.setDate(actual.getDate() + 1);
  }
  return fechas;
}

document.getElementById('generarBtn').addEventListener('click', () => {
  recogerAusencias();
  const resultado = [];

  ausencias.forEach(aus => {
    const fechas = rangoFechas(aus.desde, aus.hasta);
    fechas.forEach(f => {
      const diaSemana = obtenerDiaSemana(f);
      if (diaSemana === "Sábado" || diaSemana === "Domingo") return;

      const horasAfectadas = aus.horasSel.includes("Día completo")
        ? ["9:00–10:00", "10:00–11:00", "11:30–12:30", "14:30–15:15", "15:15–16:00"]
        : aus.horasSel;

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
            fecha: f.toLocaleDateString("es-ES"),
            diaSemana,
            hora,
            clase: cl["Clase"],
            profesorAusente: aus.profesor
          });
        });
      });
    });
  });

  generarCoberturas(resultado);
});

// === BLOQUE 3: Asignación de coberturas ===
function generarCoberturas(listaClases) {
  const tbody = document.querySelector('#resultado tbody');
  tbody.innerHTML = "";
  const resumen = {};
  const ocupaciones = {};

  listaClases.sort((a, b) => new Date(a.fecha) - new Date(b.fecha));
  const fechasUnicas = [...new Set(listaClases.map(i => i.fecha))];

  fechasUnicas.forEach(fecha => {
    const trHead = document.createElement("tr");
    trHead.innerHTML = `<td colspan="4" style="background:#004080;color:#fff;font-weight:bold;text-align:left">
      Sustituciones del ${fecha}
    </td>`;
    tbody.appendChild(trHead);

    const clasesDelDia = listaClases.filter(i => i.fecha === fecha);
    const profesoresAusentes = [...new Set(clasesDelDia.map(c => c.profesorAusente))];

    profesoresAusentes.forEach(profAusente => {
      const trSubhead = document.createElement("tr");
      trSubhead.innerHTML = `<td colspan="4" style="font-weight:bold;color:#003366;text-align:left">Profesor ausente: ${profAusente}</td>`;
      tbody.appendChild(trSubhead);

      const trCols = document.createElement("tr");
      trCols.innerHTML = `<th>Hora</th><th>Clase</th><th>Profesor Ausente</th><th>Profesor que cubre</th>`;
      trCols.style.background = "#004080";
      trCols.style.color = "#fff";
      tbody.appendChild(trCols);

      const clasesProf = clasesDelDia.filter(c => c.profesorAusente === profAusente);

      clasesProf.forEach(item => {
        const dia = item.diaSemana;
        const hora = item.hora;
        const clase = item.clase;
        let profesorCubre = "";

        const claveHora = `${dia}_${hora}`;
        if (!ocupaciones[claveHora]) ocupaciones[claveHora] = new Set();
        const profesorLibre = nombre => nombre && !ocupaciones[claveHora].has(nombre);

        const guardias = workbookData["GUARDIAS"].filter(g => g["Día"] === dia && g["Hora"] === hora);
        if (guardias.length > 0) {
          const alterna = guardiaAlterna[hora] || false;
          const posible = !alterna ? guardias[0]["Profesor de guardia"]
            : guardias[0]["Profesor de guardia 2"] || guardias[0]["Profesor de guardia"];
          if (profesorLibre(posible)) {
            profesorCubre = posible;
            guardiaAlterna[hora] = !alterna;
          }
        }

        if (!profesorCubre) {
          const filaHorario = workbookData["HORARIO"].find(h => h["Día"] === dia && h["Hora"] === hora && h["Clase"] === clase);
          if (filaHorario) {
            const posibles = [filaHorario["Profesor de apoyo 1"], filaHorario["Profesor de apoyo 2"]];
            profesorCubre = posibles.find(p => p !== profAusente && profesorLibre(p));
          }
        }

        if (!profesorCubre) {
          const apoyos = workbookData["HORARIO"].filter(h =>
            h["Día"] === dia &&
            h["Hora"] === hora &&
            h["Clase"] !== clase &&
            profesorLibre(h["Profesor de apoyo 1"]) &&
            !["convivencia", "pastoral"].some(t => (h["Profesor de apoyo 1"] || "").toLowerCase().includes(t))
          );
          if (apoyos.length > 0) profesorCubre = apoyos[0]["Profesor de apoyo 1"];
        }

        if (!profesorCubre) {
          const pastoral = workbookData["PASTORAL_CONVIVENCIA"].find(p =>
            p["Día"] === dia && p["Hora"] === hora && p["Disponible"] === "Sí");
          if (pastoral && profesorLibre(pastoral["Profesor"]))
            profesorCubre = pastoral["Profesor"];
        }

        if (!profesorCubre) profesorCubre = "No disponible";

        ocupaciones[claveHora].add(profesorCubre);
        if (!resumen[profesorCubre]) resumen[profesorCubre] = { guardias: 0, horas: 0 };
        resumen[profesorCubre].guardias++;
        resumen[profesorCubre].horas++;

        const fechaObj = new Date(item.fecha);
        const anioCorto = fechaObj.getFullYear().toString().slice(-2);
        const fechaTexto = `${fechaObj.getDate()}/${fechaObj.getMonth() + 1}/${anioCorto}`;

        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${hora}</td><td>${clase}</td><td>${profAusente}</td><td>${profesorCubre}</td>`;
        tbody.appendChild(tr);
      });
    });
  });
}

// === INICIO ===
window.addEventListener('DOMContentLoaded', cargarExcel);
