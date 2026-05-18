// static/app.js

let sessionId = null;
let resultId = null;
let profesores = [];
let clases = [];
let restricciones = [];
let resultRows = [];
let resultColumns = [];

const $ = (id) => document.getElementById(id);

function show(id) {
    const el = $(id);
    if (el) el.classList.remove("hidden");
}

function hide(id) {
    const el = $(id);
    if (el) el.classList.add("hidden");
}

function setStatus(message, type = "info") {
    const el = $("upload-status");
    el.className = "";
    el.classList.add("status-box", `status-${type}`);
    el.textContent = message;
    show("upload-status");
}

function escapeHtml(value) {
    if (value === null || value === undefined) return "";
    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function fillSelect(selectId, items, placeholder = "— Seleccionar —") {
    const select = $(selectId);
    if (!select) return;

    select.innerHTML = "";

    const first = document.createElement("option");
    first.value = "";
    first.textContent = placeholder;
    select.appendChild(first);

    items.forEach((item) => {
        const opt = document.createElement("option");

        if (typeof item === "string") {
            opt.value = item;
            opt.textContent = item;
        } else {
            opt.value = item.id;
            opt.textContent = item.label;
        }

        select.appendChild(opt);
    });
}

function renderTable(theadId, tbodyId, rows, cols) {
    const thead = $(theadId);
    const tbody = $(tbodyId);

    if (!thead || !tbody) return;

    thead.innerHTML = "";
    tbody.innerHTML = "";

    const trHead = document.createElement("tr");
    cols.forEach((col) => {
        const th = document.createElement("th");
        th.textContent = col;
        trHead.appendChild(th);
    });
    thead.appendChild(trHead);

    rows.forEach((row) => {
        const tr = document.createElement("tr");

        cols.forEach((col) => {
            const td = document.createElement("td");
            td.textContent = row[col] ?? "";

            if (col === "Estado") {
                if (row[col] === "Asignado") td.classList.add("cell-success");
                if (row[col] === "No asignado") td.classList.add("cell-error");
            }

            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    });
}

async function uploadFile(file) {
    if (!file) return;

    if (!file.name.toLowerCase().endsWith(".xlsx")) {
        setStatus("El archivo debe ser .xlsx", "error");
        return;
    }

    setStatus("Subiendo y leyendo archivo...", "info");

    const formData = new FormData();
    formData.append("file", file);

    try {
        const response = await fetch("/upload", {
            method: "POST",
            body: formData,
        });

        const data = await response.json();

        if (!response.ok) {
            setStatus(data.detail || "Error al cargar el archivo.", "error");
            return;
        }

        sessionId = data.session_id;
        profesores = data.profesores || [];
        clases = data.clases || [];
        restricciones = [];
        resultRows = [];
        resultId = null;

        setStatus(
            `Archivo cargado correctamente: ${data.total_profs} profesores y ${data.total_hor} horarios.`,
            "success"
        );

        $("preview-info").textContent =
            `Profesores detectados: ${data.total_profs} · Horarios por asignar: ${data.total_hor}`;

        renderTable(
            "preview-thead",
            "preview-tbody",
            data.preview || [],
            data.preview_cols || []
        );

        fillSelect("rest-profesor", profesores, "— Seleccionar profesor —");
        fillSelect("rest-clase", clases, "— Seleccionar clase —");
        fillSelect("rest-clase-hor", clases, "— Seleccionar clase —");
        fillSelect("rest-profesor-nivel", profesores, "— Seleccionar profesor —");

        const niveles = [...new Set(
            clases
                .map((c) => c.curso || "")
                .filter((c) => c.trim() !== "")
        )].sort();

        fillSelectSimple("rest-nivel", niveles, "— Seleccionar nivel o materia —");

        renderRestricciones();

        show("sec-preview");
        show("sec-restrictions");
        show("sec-optimize");
        hide("sec-results");
    } catch (err) {
        console.error(err);
        setStatus("Error inesperado al subir el archivo. Revisa la terminal.", "error");
    }
}

function updateRestriccionForm() {
  const tipo = $("rest-tipo").value;

  hide("rest-row-prof-clase");
  hide("rest-row-horario");
  hide("rest-row-nivel");

  const explicaciones = {
    obligatoria: `
      <strong>Asignar una clase obligatoriamente a un profesor.</strong><br>
      Use esta opción cuando una clase específica debe quedar asignada a un profesor determinado.
      El sistema intentará respetar esta regla siempre que el profesor tenga disponibilidad y no tenga cruce de horario.
      <br><em>Ejemplo: “Diana debe dictar Inglés 6 - Grupo 2”.</em>
    `,

    exclusivo: `
      <strong>Hacer que una clase solo pueda ser dictada por un profesor.</strong><br>
      Use esta opción cuando una clase no puede ser asignada a ningún otro profesor.
      Si ese profesor no está disponible, la clase quedará sin asignar.
      <br><em>Ejemplo: “Este grupo de Inglés 4 solo puede dictarlo Diana”.</em>
    `,

    permitida: `
      <strong>Permitir que un profesor pueda dictar una clase específica.</strong><br>
      Use esta opción cuando quiere habilitar o dar preferencia a un profesor para una clase.
      No obliga necesariamente la asignación, pero ayuda al modelo a considerarlo como opción válida.
      <br><em>Ejemplo: “Sarah puede dictar este grupo si el modelo lo necesita”.</em>
    `,

    solo_una: `
      <strong>Limitar a un profesor para que solo dicte una clase específica.</strong><br>
      Use esta opción cuando un profesor solo debe participar en una clase puntual y no debe ser asignado a otras clases.
      <br><em>Ejemplo: “Este profesor solo puede dictar Inglés 5 - Grupo 1”.</em>
    `,

    debe_dictar: `
      <strong>Forzar que un profesor dicte una clase específica.</strong><br>
      Es similar a la asignación obligatoria. Use esta opción cuando, por decisión académica o administrativa,
      un profesor debe quedar asignado a una clase concreta.
      <br><em>Ejemplo: “Liliana debe dictar este grupo de Minor”.</em>
    `,

    horario_especifico: `
      <strong>Cambiar el día y horario de una clase.</strong><br>
      Use esta opción cuando una clase debe moverse a otro día u otra franja horaria antes de optimizar.
      El sistema validará la disponibilidad de los profesores usando el nuevo horario.
      <br><em>Ejemplo: “Mover Inglés 2 del lunes 8:00 a martes 10:00”.</em>
    `,

    profesor_solo_nivel: `
      <strong>Limitar a un profesor a un solo nivel o materia.</strong><br>
      Use esta opción cuando un profesor no debe quedar con muchos niveles diferentes.
      Si selecciona un nivel, el modelo solo podrá asignarle clases de ese nivel o materia.
      <br><em>Ejemplo: “Diana solo debe dictar Inglés 6”.</em>
    `,

    nivel_solo_profesor: `
      <strong>Hacer que un nivel o materia solo lo dicte un profesor.</strong><br>
      Use esta opción cuando todos los grupos de un mismo nivel deben concentrarse en un solo profesor.
      El modelo intentará asignar ese nivel únicamente al profesor seleccionado.
      <br><em>Ejemplo: “Todos los grupos de Inglés 1 deben ser dictados por Sarah”.</em>
    `,
  };

  const explicacion = $("rest-explicacion");
  if (explicacion) {
    explicacion.innerHTML = explicaciones[tipo] || "Seleccione un tipo de restricción.";
  }

  if (tipo === "horario_especifico") {
    show("rest-row-horario");
  } else if (tipo === "profesor_solo_nivel" || tipo === "nivel_solo_profesor") {
    show("rest-row-nivel");
  } else {
    show("rest-row-prof-clase");
  }
}

function getClaseLabel(id) {
    const found = clases.find((c) => String(c.id) === String(id));
    return found ? found.label : id;
}

function addRestriccion() {
    const tipo = $("rest-tipo").value;

    if (tipo === "horario_especifico") {
        const idHorario = $("rest-clase-hor").value;
        const dia = $("rest-dia-esp").value;
        const horaIni = $("rest-h-ini").value;
        const horaFin = $("rest-h-fin").value;

        if (!idHorario || !dia || !horaIni || !horaFin) {
            alert("Completa clase, día, hora inicio y hora fin.");
            return;
        }

        if (horaFin <= horaIni) {
            alert("La hora fin debe ser mayor que la hora inicio.");
            return;
        }

        restricciones.push({
            tipo: "horario_especifico",
            id_horario: idHorario,
            dia_especifico: dia,
            hora_inicio_especifica: horaIni,
            hora_fin_especifica: horaFin,
            label: `Horario específico: ${getClaseLabel(idHorario)} → ${dia} ${horaIni}-${horaFin}`,
        });

    } else if (tipo === "profesor_solo_nivel" || tipo === "nivel_solo_profesor") {
        const profesor = $("rest-profesor-nivel").value;
        const curso = $("rest-nivel").value;

        if (!profesor || !curso) {
            alert("Selecciona profesor y nivel/materia.");
            return;
        }

        const texto =
            tipo === "profesor_solo_nivel"
                ? `Profesor solo puede dictar nivel/materia: ${profesor} → ${curso}`
                : `Nivel/materia solo puede ser dictado por profesor: ${curso} → ${profesor}`;

        restricciones.push({
            tipo,
            profesor,
            curso,
            label: texto,
        });

    } else {
        const profesor = $("rest-profesor").value;
        const idHorario = $("rest-clase").value;

        if (!profesor || !idHorario) {
            alert("Selecciona profesor y clase.");
            return;
        }

        const tipoTexto = {
            obligatoria: "Profesor obligatorio",
            exclusivo: "Clase exclusiva para profesor",
            permitida: "Profesor permitido para clase",
            solo_una: "Profesor solo puede dictar esta clase",
            debe_dictar: "Profesor debe dictar esta clase",
        };

        restricciones.push({
            tipo,
            profesor,
            id_horario: idHorario,
            label: `${tipoTexto[tipo] || tipo}: ${profesor} → ${getClaseLabel(idHorario)}`,
        });
    }

    renderRestricciones();
}

function removeRestriccion(index) {
    restricciones.splice(index, 1);
    renderRestricciones();
}

function renderRestricciones() {
    const list = $("rest-items");
    if (!list) return;

    list.innerHTML = "";

    if (restricciones.length === 0) {
        hide("restrictions-list");
        return;
    }

    restricciones.forEach((rest, index) => {
        const li = document.createElement("li");
        li.className = "rest-item";

        const span = document.createElement("span");
        span.textContent = rest.label || JSON.stringify(rest);

        const btn = document.createElement("button");
        btn.className = "btn btn-danger btn-sm";
        btn.textContent = "Eliminar";
        btn.onclick = () => removeRestriccion(index);

        li.appendChild(span);
        li.appendChild(btn);
        list.appendChild(li);
    });

    show("restrictions-list");
}

async function runOptimize() {
    if (!sessionId) {
        alert("Primero carga un archivo Excel.");
        return;
    }

    $("btn-optimize").disabled = true;
    show("optimize-spinner");

    const formData = new FormData();
    formData.append("session_id", sessionId);
    formData.append("restricciones", JSON.stringify(restricciones));

    try {
        const response = await fetch("/optimize", {
            method: "POST",
            body: formData,
        });

        const data = await response.json();

        if (!response.ok) {
            alert(data.detail || "Error al optimizar.");
            return;
        }

        resultId = data.result_id;
        resultRows = data.tabla || [];
        resultColumns = resultRows.length ? Object.keys(resultRows[0]) : [];

        $("result-metodo").textContent = `Método usado: ${data.metodo}`;
        $("m-total").textContent = data.total;
        $("m-asig").textContent = data.asignados;
        $("m-noasig").textContent = data.no_asig;
        $("m-cob").textContent = `${data.cobertura}%`;

        populateResultFilters(resultRows);
        applyFilters();

        show("sec-results");
        window.scrollTo({ top: $("sec-results").offsetTop - 20, behavior: "smooth" });
    } catch (err) {
        console.error(err);
        alert("Error inesperado al optimizar. Revisa la terminal.");
    } finally {
        $("btn-optimize").disabled = false;
        hide("optimize-spinner");
    }
}

function populateResultFilters(rows) {
    const dias = [...new Set(rows.map((r) => r["Día"]).filter(Boolean))].sort();
    const cursos = [...new Set(rows.map((r) => r["Curso"]).filter(Boolean))].sort();

    fillSelectSimple("fil-dia", dias, "Todos");
    fillSelectSimple("fil-curso", cursos, "Todos");
}

function fillSelectSimple(selectId, values, firstLabel = "Todos") {
    const select = $(selectId);
    if (!select) return;

    select.innerHTML = "";

    const first = document.createElement("option");
    first.value = "";
    first.textContent = firstLabel;
    select.appendChild(first);

    values.forEach((value) => {
        const opt = document.createElement("option");
        opt.value = value;
        opt.textContent = value;
        select.appendChild(opt);
    });
}

function applyFilters() {
    const estado = $("fil-estado").value;
    const dia = $("fil-dia").value;
    const curso = $("fil-curso").value;
    const buscar = $("fil-buscar").value.toLowerCase().trim();

    let filtered = [...resultRows];

    if (estado) {
        filtered = filtered.filter((r) => r["Estado"] === estado);
    }

    if (dia) {
        filtered = filtered.filter((r) => r["Día"] === dia);
    }

    if (curso) {
        filtered = filtered.filter((r) => r["Curso"] === curso);
    }

    if (buscar) {
        filtered = filtered.filter((r) =>
            Object.values(r).some((v) => String(v).toLowerCase().includes(buscar))
        );
    }

    renderTable("result-thead", "result-tbody", filtered, resultColumns);

    $("result-count").textContent =
        `Mostrando ${filtered.length} de ${resultRows.length} registros.`;
}

function downloadResult() {
    if (!resultId) {
        alert("Primero ejecuta la optimización.");
        return;
    }

    window.location.href = `/download/${resultId}`;
}

function setupUploadEvents() {
    const input = $("file-input");
    const zone = $("upload-zone");

    if (!input || !zone) return;

    input.addEventListener("change", () => {
        const file = input.files[0];
        uploadFile(file);
    });

    zone.addEventListener("dragover", (event) => {
        event.preventDefault();
        zone.classList.add("drag-over");
    });

    zone.addEventListener("dragleave", () => {
        zone.classList.remove("drag-over");
    });

    zone.addEventListener("drop", (event) => {
        event.preventDefault();
        zone.classList.remove("drag-over");

        const file = event.dataTransfer.files[0];
        if (file) uploadFile(file);
    });
}

document.addEventListener("DOMContentLoaded", () => {
    setupUploadEvents();
    updateRestriccionForm();
});
