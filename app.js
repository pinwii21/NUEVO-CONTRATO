// === CONFIGURACIÓN GENERAL ===

// Usuarios permitidos
const usuarios = {
  admin: "1234",
  kevin: "admin2025",
};

// Campos del formulario y tabla (coinciden con tu dataset)
const campos = [
  "CODIGO TRABAJADOR",
  "NOMBRE",
  "RUTA",
  "CEDULA",
  "TELEFONO",
  "DISCAPACIDAD",
  "HORARIO DE TRABAJO",
  "LUGAR DE TRABAJO",
  "ROL",
  "AREA",
  "MODALIDAD DE CONTRATO",
  "TIPO HORARIO",
  "DIRECCION",
  "SUBDISIO DE TRANSPORTE",
  "LONGITUD",
  "LATITUD",
  "VERIFICACION",
];

// Variables globales
let geojsonData = null;
let usuarioLogueado = false;
let geojsonLayer = null;
const capasOverlay = {};

// === INICIALIZACIÓN DEL MAPA LEAFLET ===

// Capas base
const osmBase = L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
  attribution: "© OpenStreetMap contributors",
  maxZoom: 19,
});

const googlemaps = L.tileLayer("https://mt1.google.com/vt/lyrs=r&x={x}&y={y}&z={z}", {
  attribution:
    "Tiles © Esri — Source: Esri, Maxar, Earthstar Geographics, and the GIS User Community",
  maxZoom: 25,
});

const googleBase = L.tileLayer("https://mt1.google.com/vt/lyrs=s&x={x}&y={y}&z={z}", {
  attribution:
    "Tiles © Esri — Source: Esri, Maxar, Earthstar Geographics, and the GIS User Community",
  maxZoom: 25,
});

// Crear mapa centrado en Quito
const map = L.map("map", {
  center: [-0.180653, -78.467838],
  zoom: 13,
  layers: [osmBase],
});

const mapasBase = {
  OpenStreetMap: osmBase,
  "Google Maps": googlemaps,
  "Google Satelital": googleBase,
};

// === CARGAR DATOS DESDE EXCEL ===

const EXCEL_URL =
  "https://raw.githubusercontent.com/TU-USUARIO/TU-REPO/main/data/BASE%20DE%20DATOS.xlsx";

async function cargarExcel() {
  try {
    const response = await fetch(EXCEL_URL);

    const arrayBuffer = await response.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, {
      type: "array",
    });

    const hoja = workbook.Sheets[workbook.SheetNames[0]];

    const datos = XLSX.utils.sheet_to_json(hoja);

    const features = datos.map((fila, i) => {
      const lat = parseFloat(fila["LATITUD"]);
      const lng = parseFloat(fila["LONGITUD"]);

      return {
        type: "Feature",
        geometry: {
          type: "Point",
          coordinates: [lng, lat],
        },
        properties: fila,
        _id: i,
      };
    });

    geojsonData = {
      type: "FeatureCollection",
      features,
    };

    crearCamposFormulario();
    mostrarTabla(geojsonData);
    mostrarMapa(geojsonData);
    centrarMapa(geojsonData);
    actualizarListaPersonas(features);

  } catch (error) {
    console.error("Error cargando Excel:", error);
  }
}

cargarExcel();

// === CARGAR RUTAS DE ENTRADA Y SALIDA ===
const carpetas = [
  { dir: "Rutas_de_ENTRADA", name: "Rutas de ENTRADA", color: "#28a745" },
  { dir: "Rutas_de_SALIDA", name: "Rutas de SALIDA", color: "#dc3545" },
];

async function cargarIndexYCapas() {
  for (const { dir, name, color } of carpetas) {
    try {
      const indexUrl = `https://raw.githubusercontent.com/pinwii21/IDE-TRANSPORTE/main/${dir}/index.json`;
      const idxRes = await fetch(indexUrl);
      if (!idxRes.ok) throw new Error(`No se pudo cargar: ${indexUrl}`);
      const lista = await idxRes.json();
      const grupo = L.layerGroup();

      // Cargar cada archivo de la lista (continue on error)
      await Promise.allSettled(
        lista.map(async (fichero) => {
          try {
            const geojsonUrl = `https://raw.githubusercontent.com/pinwii21/IDE-TRANSPORTE/main/${dir}/${fichero}`;
            const res = await fetch(geojsonUrl);
            if (!res.ok) throw new Error(`Error al cargar: ${geojsonUrl}`);
            const data = await res.json();

            const capa = L.geoJSON(data, {
              style: { color, weight: 3 },
              onEachFeature: (feature, layer) => {
                let popup = `<b>${fichero}</b><br>`;
                for (const k in feature.properties) {
                  popup += `<b>${k}:</b> ${feature.properties[k]}<br>`;
                }
                layer.bindPopup(popup);
              },
            });

            capa._nombreArchivo = fichero;
            capa.addTo(grupo);
          } catch (error) {
            console.warn(`Error en ${fichero}:`, error.message);
          }
        })
      );

      capasOverlay[name] = grupo;
      grupo.addTo(map);
    } catch (error) {
      console.error(`Error en carpeta ${dir}:`, error.message);
    }
  }

  L.control.layers(mapasBase, capasOverlay, { collapsed: false }).addTo(map);
  inicializarFiltrosRutas();
}
cargarIndexYCapas();

// === FORMULARIO Y TABLA ===
function crearCamposFormulario() {
  const cont = document.getElementById("camposForm");
  if (!cont) return;
  cont.innerHTML = "";
  campos.forEach((campo) => {
    const label = document.createElement("label");
    label.textContent = campo + ":";
    const input = document.createElement("input");
    input.id = campo;
    input.type = "text";
    input.required = ["CODIGO TRABAJADOR", "NOMBRE", "LATITUD", "LONGITUD"].includes(
      campo
    );
    label.appendChild(input);
    cont.appendChild(label);
  });
}

function mostrarTabla(data) {
  const cont = document.getElementById("tabla");
  if (!cont || !data?.features) return;

  let html = `<table>
    <thead>
      <tr>
        <th>#</th>`;

  // Encabezados
  campos.forEach((c) => {
    html += `<th>${c}</th>`;
  });

  html += `</tr></thead><tbody>`;

  // Filas
  data.features.forEach((f, i) => {
    html += `<tr data-id="${f._id}">
      <td>${i + 1}</td>`;

    campos.forEach((campo) => {
      let val = f.properties?.[campo] || "";
      if (campo === "LATITUD") val = f.geometry?.coordinates?.[1] ?? "";
      if (campo === "LONGITUD") val = f.geometry?.coordinates?.[0] ?? "";

      html += `<td ${
        usuarioLogueado ? 'contenteditable="true"' : ""
      } data-feature-id="${f._id}" data-attr="${campo}">
        ${val}
      </td>`;
    });

    html += `</tr>`;
  });

  html += `</tbody></table>`;

  // Insertar tabla en el DOM
  cont.innerHTML = html;

  // ===============================
  // CLICK EN FILA → ZOOM EN EL MAPA
  // ===============================
  document.querySelectorAll("#tabla tbody tr").forEach((tr) => {
    tr.addEventListener("click", () => {
      const id = parseInt(tr.dataset.id);
      const feature = geojsonData.features.find((f) => f._id === id);
      if (!feature || !feature.geometry) return;

      const lng = feature.geometry.coordinates[0];
      const lat = feature.geometry.coordinates[1];

      // Zoom al punto
      map.setView([lat, lng], 17, { animate: true });

      // Resaltar fila seleccionada
      document.querySelectorAll("#tabla tr").forEach((r) =>
        r.classList.remove("resaltado")
      );
      tr.classList.add("resaltado");
    });
  });

  // Columnas y edición
  crearTogglesColumnas();
  if (usuarioLogueado) asignarEventosEdicion();
}


// === EDICIÓN EN TABLA ===
function asignarEventosEdicion() {
  document.querySelectorAll('td[contenteditable="true"]').forEach((td) => {
    td.addEventListener("input", () => {
      const id = parseInt(td.dataset.featureId);
      const campo = td.dataset.attr;
      const valor = td.textContent.trim(); // conserva mayúsculas/minúsculas ingresadas
      const feature = geojsonData.features.find((f) => f._id === id);
      if (!feature) return;

      // Actualiza la propiedad correspondiente
      feature.properties[campo] = valor;

      // Si se edita latitud o longitud, actualizar la geometría y refrescar mapa
      if (campo === "LATITUD" || campo === "LONGITUD") {
        const lat = parseFloat(feature.properties["LATITUD"]);
        const lng = parseFloat(feature.properties["LONGITUD"]);
        if (!isNaN(lat) && !isNaN(lng)) {
          feature.geometry = { type: "Point", coordinates: [lng, lat] };
          mostrarMapa(geojsonData);
          centrarMapa(geojsonData);
        }
      }
    });
  });
}

// === MOSTRAR MAPA DE PUNTOS ===
function mostrarMapa(data) {
  if (geojsonLayer) map.removeLayer(geojsonLayer);
  geojsonLayer = L.geoJSON(data, {
    pointToLayer: (feature, latlng) =>
      L.circleMarker(latlng, {
        radius: 6,
        fillColor:
          getComputedStyle(document.documentElement)
            .getPropertyValue("--color-primario")
            .trim() || "#004d99",
        color: "#000",
        weight: 1,
        opacity: 1,
        fillOpacity: 0.9,
      }),
    onEachFeature: (feature, layer) => {
      let popup = "";
      for (const key in feature.properties) {
        popup += `<b>${key}:</b> ${feature.properties[key]}<br>`;
      }
      layer.bindPopup(popup);

      layer.on("click", () => {
        const row = document.querySelector(`td[data-feature-id='${feature._id}']`);
        if (!row) return;
        const tr = row.closest("tr");
        if (!tr) return;
        document.querySelectorAll("#tabla tr").forEach((r) => r.classList.remove("resaltado"));
        tr.classList.add("resaltado");
        tr.scrollIntoView({ behavior: "smooth", block: "center" });
      });
    },
  }).addTo(map);
}

// === CENTRAR MAPA EN LOS PUNTOS VISIBLES ===
function centrarMapa(data) {
  if (!data?.features?.length) return;
  const coords = data.features
    .map((f) => f.geometry?.coordinates)
    .filter((c) => Array.isArray(c) && !isNaN(c[0]) && !isNaN(c[1]));
  if (coords.length === 0) return;
  const lats = coords.map((c) => c[1]);
  const lngs = coords.map((c) => c[0]);
  const bounds = [
    [Math.min(...lats), Math.min(...lngs)],
    [Math.max(...lats), Math.max(...lngs)],
  ];
  map.fitBounds(bounds, { padding: [40, 40] });
}

// === FILTRO DOBLE DE PERSONAL ===
const searchInput1 = document.getElementById("searchInput1");
const filterField1 = document.getElementById("filterField1");
const searchInput2 = document.getElementById("searchInput2");
const filterField2 = document.getElementById("filterField2");
const suggestions1 = document.getElementById("suggestions1");
const suggestions2 = document.getElementById("suggestions2");

if (searchInput1) searchInput1.addEventListener("input", filtrarDatos);
if (filterField1)
  filterField1.addEventListener("change", () => {
    actualizarSugerencias(filterField1.value, suggestions1);
    filtrarDatos();
  });
if (searchInput2) searchInput2.addEventListener("input", filtrarDatos);
if (filterField2)
  filterField2.addEventListener("change", () => {
    actualizarSugerencias(filterField2.value, suggestions2);
    filtrarDatos();
  });

function filtrarDatos() {
  if (!geojsonData) return;

  const campo1 = filterField1?.value || "";
  const texto1 = (searchInput1?.value || "").toLowerCase();

  const campo2 = filterField2?.value || "";
  const texto2 = (searchInput2?.value || "").toLowerCase();

  const filtrados = geojsonData.features.filter((f) => {
    const valor1 = (f.properties?.[campo1] || "").toString().toLowerCase();
    const cumpleFiltro1 = !campo1 || valor1.includes(texto1);

    let cumpleFiltro2 = true;
    if (campo2 && texto2) {
      const valor2 = (f.properties?.[campo2] || "").toString().toLowerCase();
      cumpleFiltro2 = valor2.includes(texto2);
    }

    return cumpleFiltro1 && cumpleFiltro2;
  });

  const dataset = { type: "FeatureCollection", features: filtrados };
  mostrarTabla(dataset);
  mostrarMapa(dataset);
  centrarMapa(dataset);
  actualizarListaPersonas(filtrados);
}

function actualizarSugerencias(campo, datalistElement) {
  if (!campo || !geojsonData || !datalistElement) return;

  const valoresUnicos = [
    ...new Set(
      geojsonData.features
        .map((f) => f.properties?.[campo])
        .filter((v) => v !== undefined && v !== null && v !== "")
    ),
  ].sort();

  datalistElement.innerHTML = "";
  valoresUnicos.forEach((valor) => {
    const option = document.createElement("option");
    option.value = valor;
    datalistElement.appendChild(option);
  });
}

if (filterField1?.value) actualizarSugerencias(filterField1.value, suggestions1);
if (filterField2?.value) actualizarSugerencias(filterField2.value, suggestions2);

// === ACTUALIZAR SELECT DE PERSONAS (si lo usas en otra parte) ===
function actualizarListaPersonas(lista) {
  const select = document.getElementById("personSelect");
  if (!select) return;
  select.innerHTML = '<option value="">-- Selecciona persona --</option>';
  lista.forEach((f) => {
    const nombre = f.properties?.["NOMBRE"] || "SIN NOMBRE";
    const codigo = f.properties?.["CODIGO TRABAJADOR"] || "SIN CÓDIGO";
    const value = f._id;
    const option = document.createElement("option");
    option.value = value;
    option.textContent = `${nombre} (${codigo})`;
    select.appendChild(option);
  });
}

// === LOGIN / LOGOUT ===
const loginForm = document.getElementById("loginForm");
if (loginForm) {
  loginForm.addEventListener("submit", function (e) {
    e.preventDefault();
    const usuario = document.getElementById("usuario").value.trim();
    const clave = document.getElementById("clave").value.trim();

    if (usuarios[usuario] && usuarios[usuario] === clave) {
      usuarioLogueado = true;
      document.getElementById("loginContainer").style.display = "none";
      document.getElementById("addForm").style.display = "block";
      document.getElementById("logoutBtn").style.display = "inline-block";
      const descargarBtn = document.getElementById("descargarGeoJSONBtn");
      if (descargarBtn) descargarBtn.style.display = "inline-block";
      mostrarTabla(geojsonData);
    } else {
      alert("Usuario o clave incorrectos");
    }
  });
}

const logoutBtn = document.getElementById("logoutBtn");
if (logoutBtn) {
  logoutBtn.addEventListener("click", () => {
    usuarioLogueado = false;
    document.getElementById("loginContainer").style.display = "block";
    document.getElementById("addForm").style.display = "none";
    document.getElementById("logoutBtn").style.display = "none";
    const descargarBtn = document.getElementById("descargarGeoJSONBtn");
    if (descargarBtn) descargarBtn.style.display = "none";
    mostrarTabla(geojsonData);
  });
}

// === AGREGAR NUEVO PERSONAL ===
const addForm = document.getElementById("addForm");
if (addForm) {
  addForm.addEventListener("submit", async (e) => {
    e.preventDefault();

    const nuevo = {};
    let valid = true;

    // Validación de obligatorios
    for (const campo of campos) {
      const inputEl = document.getElementById(campo);
      const val = (inputEl?.value || "").trim();
      if (["CODIGO TRABAJADOR", "NOMBRE", "LATITUD", "LONGITUD"].includes(campo) && !val) {
        alert(`El campo ${campo} es obligatorio`);
        valid = false;
        break;
      }
      // Guardamos en mayúsculas para mantener formato en tu dataset
      nuevo[campo] = val.toUpperCase();
    }

    if (!valid) return;

    const lat = parseFloat(nuevo["LATITUD"]);
    const lng = parseFloat(nuevo["LONGITUD"]);
    if (isNaN(lat) || isNaN(lng)) {
      alert("LATITUD y LONGITUD deben ser números válidos");
      return;
    }

    const nuevoFeature = {
      type: "Feature",
      geometry: { type: "Point", coordinates: [lng, lat] },
      properties: nuevo,
      _id: geojsonData.features.length,
    };

    geojsonData.features.push(nuevoFeature);
    mostrarTabla(geojsonData);
    mostrarMapa(geojsonData);
    centrarMapa(geojsonData);
    actualizarListaPersonas(geojsonData.features);

    addForm.reset();
    alert("Nuevo personal agregado correctamente. Ahora puedes descargar el archivo actualizado.");
  });
}

// === DESCARGAR GEOJSON ACTUALIZADO ===
function descargarGeoJSON() {
  if (!geojsonData) return;
  const blob = new Blob([JSON.stringify(geojsonData, null, 2)], {
    type: "application/geo+json",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "BASE_DATOS_TRANSPORTE_2025_actualizado.geojson";
  document.body.appendChild(a);
  a.dispatchEvent(new MouseEvent("click"));
  setTimeout(() => {
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }, 100);
}

const descargarBtn = document.getElementById("descargarGeoJSONBtn");
if (descargarBtn) {
  descargarBtn.addEventListener("click", descargarGeoJSON);
}

// === FILTRO DOBLE AUTOMÁTICO EN RUTAS ===
function inicializarFiltrosRutas() {
  const camposFiltro = [
    "RUTA",
    "NUMERO_RUTA",
    "HORARIO",
    "DESTINO",
    "TIPO_UNIDAD",
    "FRECUENCIA",
    "TIPO_HORARIO",
    "KM",
    "KM_ACT",
  ];
  const campo1 = document.getElementById("campo1");
  const campo2 = document.getElementById("campo2");
  const valor1 = document.getElementById("valor1");
  const valor2 = document.getElementById("valor2");
  if (!campo1 || !campo2 || !valor1 || !valor2) return;

  // Llenar selects de campos
  camposFiltro.forEach((c) => {
    const opt1 = document.createElement("option");
    const opt2 = document.createElement("option");
    opt1.value = c;
    opt2.value = c;
    opt1.textContent = c;
    opt2.textContent = c;
    campo1.appendChild(opt1);
    campo2.appendChild(opt2);
  });

  // Eventos
  campo1.addEventListener("change", () => {
    actualizarValoresFiltro(campo1.value, valor1, () => {
      actualizarValoresFiltro(campo2.value, valor2, aplicarFiltroMultiple);
    });
  });

  campo2.addEventListener("change", () => {
    actualizarValoresFiltro(campo2.value, valor2, aplicarFiltroMultiple);
  });

  valor1.addEventListener("change", () => {
    actualizarValoresFiltro(campo2.value, valor2, aplicarFiltroMultiple);
  });

  valor2.addEventListener("change", aplicarFiltroMultiple);

  // Helper: actualizar valores posibles para un campo (con filtro primario aplicado)
  function actualizarValoresFiltro(campo, selectDestino, callback) {
    if (!campo) return;
    const valores = new Set();

    for (const grupo of Object.values(capasOverlay)) {
      // Algunos grupos pueden tener subcapas
      grupo.eachLayer((capa) => {
        if (!capa.feature && capa.eachLayer) {
          capa.eachLayer((layer) => {
            const props = layer.feature?.properties || {};

            // Aplicar filtro primario para refinar el segundo filtro
            const filtroPrimario = campo1.value;
            const valorPrimario = valor1.value;
            if (filtroPrimario && valorPrimario && campo !== filtroPrimario) {
              const match =
                (props[filtroPrimario] || "").toString().toLowerCase() ===
                valorPrimario.toLowerCase();
              if (!match) return;
            }

            const valor = props[campo];
            if (valor !== undefined && valor !== null && valor !== "") valores.add(valor);
          });
        }
      });
    }

    selectDestino.innerHTML = '<option value="">-- Todos --</option>';
    [...valores].sort().forEach((v) => {
      const opt = document.createElement("option");
      opt.value = v;
      opt.textContent = v;
      selectDestino.appendChild(opt);
    });

    if (callback) callback();
  }

  // Mostrar/ocultar capas según filtros
  function aplicarFiltroMultiple() {
    const c1 = campo1.value;
    const v1 = valor1.value;
    const c2 = campo2.value;
    const v2 = valor2.value;

    for (const grupo of Object.values(capasOverlay)) {
      grupo.eachLayer((capa) => {
        if (!capa.feature && capa.eachLayer) {
          capa.eachLayer((layer) => {
            const props = layer.feature?.properties || {};
            const visible1 =
              !v1 || (props[c1] + "").toString().toLowerCase() === v1.toLowerCase();
            const visible2 =
              !v2 || (props[c2] + "").toString().toLowerCase() === v2.toLowerCase();
            const visible = visible1 && visible2;

            if (visible) {
              if (!map.hasLayer(layer)) layer.addTo(map);
            } else {
              if (map.hasLayer(layer)) map.removeLayer(layer);
            }
          });
        }
      });
    }
  }
}

// === TOGGLE DE COLUMNAS EN TABLA ===
function toggleColumnMenu() {
  const menu = document.getElementById("columnMenu");
  if (!menu) return;
  // La clase .active se maneja en CSS para mostrar/ocultar
  menu.classList.toggle("active");
}

function crearTogglesColumnas() {
  const contenedor = document.getElementById("columnToggles");
  if (!contenedor) return;
  contenedor.innerHTML = "";

  campos.forEach((campo, i) => {
    const label = document.createElement("label");
    label.style.display = "flex";
    label.style.alignItems = "center";
    label.style.gap = "8px";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = true;
    checkbox.dataset.colIndex = i + 1; // +1 porque la primera columna es #

    checkbox.addEventListener("change", () => {
      const colIndex = parseInt(checkbox.dataset.colIndex);
      document.querySelectorAll(`#tabla table tr`).forEach((row) => {
        const celda = row.children[colIndex];
        if (celda) celda.style.display = checkbox.checked ? "" : "none";
      });
    });

    label.appendChild(checkbox);
    label.append(document.createTextNode(campo));
    contenedor.appendChild(label);
  });
}

// === TOGGLE TABLA FLOTANTE ===
function toggleTabla() {
  const panel = document.getElementById("panelTabla");
  if (!panel) return;

  panel.classList.toggle("oculto");

  // Forzar a Leaflet a recalcular tamaño
  setTimeout(() => {
    map.invalidateSize();
  }, 300);
}


// Botón flotante
const toggleTablaBtn = document.getElementById("toggleTablaBtn");
if (toggleTablaBtn) {
  toggleTablaBtn.addEventListener("click", toggleTabla);
}

// === DRAG TABLA ===
(function hacerTablaDraggable() {
  const panel = document.getElementById("panelTabla");
  const header = document.getElementById("tablaHeader");
  if (!panel || !header) return;

  let offsetX = 0, offsetY = 0, isDragging = false;

  header.addEventListener("mousedown", (e) => {
    isDragging = true;
    const rect = panel.getBoundingClientRect();
    offsetX = e.clientX - rect.left;
    offsetY = e.clientY - rect.top;
    panel.style.transition = "none";
  });

  document.addEventListener("mousemove", (e) => {
    if (!isDragging) return;
    panel.style.left = `${e.clientX - offsetX}px`;
    panel.style.top = `${e.clientY - offsetY}px`;
    panel.style.right = "auto";
    panel.style.bottom = "auto";
  });

  document.addEventListener("mouseup", () => {
    isDragging = false;
    panel.style.transition = "";
  });
})();



