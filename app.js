// ===== STATE =====
const state = {
  usuario: null,
  solicitudes: [],
  solicitudSeleccionada: null,
  filtroEstado: "Todos",
  filtroBuscar: "",
  filtroUnidad: "Todas",
  filtroDesde: "",
  filtroHasta: "",
  pagina: 1,
  pageSize: 10,
  adjuntosNueva: [],
  chartInstances: {}
};

// ===== INIT =====
async function initApp() {
  showLoading("Iniciando sesión...");
  try {
    const loggedIn = await initAuth();
    if (!loggedIn) {
      hideLoading();
      showLoginPage();
      return;
    }
    showLoading("Cargando perfil...");
    const msUser = await getCurrentUser();
    const perfil = await getUserByEmail(msUser.mail || msUser.userPrincipalName);
    if (!perfil) {
      hideLoading();
      showToast("error", "⚠️ Tu usuario no tiene acceso al sistema. Contacta al administrador.");
      showLoginPage(msUser);
      return;
    }
    state.usuario = { ...perfil, displayName: msUser.displayName };
    hideLoading();
    showApp();
  } catch (e) {
    hideLoading();
    console.error(e);
    showLoginPage();
  }
}

function showLoginPage(msUser) {
  document.getElementById("login-page").style.display = "flex";
  document.getElementById("app").style.display = "none";
  if (msUser) {
    document.getElementById("login-nombre").textContent = msUser.displayName || "";
    document.getElementById("login-correo").textContent = msUser.mail || msUser.userPrincipalName || "";
  }
  document.getElementById("btn-login").onclick = () => {
    showLoading("Redirigiendo...");
    login();
  };
}

function showApp() {
  document.getElementById("login-page").style.display = "none";
  const app = document.getElementById("app");
  app.style.display = "flex";

  const u = state.usuario;
  const nombre = u.NombreCompleto || u.displayName || "";
  document.getElementById("topbar-nombre").textContent = nombre;
  document.getElementById("topbar-rol").textContent = `${u.Rol}${u.Unidad ? ' — ' + u.Unidad : ''}`;
  const av = document.getElementById("topbar-avatar");
  if (av) av.textContent = nombre.charAt(0).toUpperCase();
  document.getElementById("btn-logout-app").onclick = logout;

  buildTabs();
  const rol = u.Rol;
  if (rol === CONFIG.roles.SECRETARIA) navigateTab("solicitudes");
  else if (rol === CONFIG.roles.DIRECTOR) navigateTab("gestion");
  else navigateTab("unidad");
}

// ===== TABS =====
function buildTabs() {
  const bar = document.getElementById("tabs-bar");
  bar.innerHTML = "";
  const rol = state.usuario.Rol;

  const tabs = [];
  if (rol === CONFIG.roles.SECRETARIA) {
    // Secretaria: solo ingreso de solicitudes
    tabs.push({ id: "solicitudes", icon: "📋", label: "Ingreso Solicitudes" });
  }
  if (rol === CONFIG.roles.DIRECTOR) {
    // Director: gestión completa
    tabs.push({ id: "gestion", icon: "⚙️", label: "Gestión" });
    tabs.push({ id: "graficos", icon: "📊", label: "Reportes" });
  }
  if (rol === CONFIG.roles.UNIDAD) {
    // Unidad: sus solicitudes + reportes
    tabs.push({ id: "unidad", icon: "🏢", label: "Mis Solicitudes" });
    tabs.push({ id: "graficos", icon: "📊", label: "Reportes" });
  }

  tabs.forEach(t => {
    const btn = document.createElement("button");
    btn.className = "tab-btn";
    btn.dataset.tab = t.id;
    btn.innerHTML = `<span>${t.icon}</span><span>${t.label}</span>`;
    btn.onclick = () => navigateTab(t.id);
    bar.appendChild(btn);
  });
}

function navigateTab(tabId) {
  document.querySelectorAll(".tab-btn").forEach(b => b.classList.toggle("active", b.dataset.tab === tabId));
  document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));

  const view = document.getElementById(`view-${tabId}`);
  if (view) view.classList.add("active");

  if (tabId === "solicitudes") renderSecretaria();
  if (tabId === "gestion") renderDirector();
  if (tabId === "unidad") renderUnidad();
  if (tabId === "graficos") renderGraficos();
}

// ===== SECRETARIA VIEW =====
async function renderSecretaria() {
  showLoading("Cargando solicitudes...");
  try {
    state.solicitudes = await getSolicitudes();
    state.solicitudes.sort((a, b) => new Date(b.FechaRecepcion) - new Date(a.FechaRecepcion));
    renderListaSolicitudes();
    renderFiltrosCompact();
    renderStatsCompact();
    renderFormNueva();
    updateTabBadges();
  } catch (e) {
    showToast("error", "Error cargando solicitudes: " + e.message);
  } finally {
    hideLoading();
  }
}

function toggleFiltros() {
  const panel = document.getElementById("panel-filtros-compact");
  const toggle = document.getElementById("filtros-toggle");
  if (panel.style.display === "none") {
    panel.style.display = "block";
    toggle.textContent = "▲ ocultar";
    renderFiltrosCompact();
  } else {
    panel.style.display = "none";
    toggle.textContent = "▼ ver";
  }
}

function renderFiltrosCompact() {
  const cont = document.getElementById("panel-filtros-compact");
  if (!cont) return;
  cont.innerHTML = `
    <div style="padding:10px;display:flex;flex-direction:column;gap:8px;">
      <input type="text" placeholder="🔍 Buscar..." value="${state.filtroBuscar}"
        style="padding:8px;border:1.5px solid var(--borde);border-radius:8px;font-size:13px;"
        oninput="state.filtroBuscar=this.value;state.pagina=1;renderListaSolicitudes();renderStatsCompact()">
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;">
        <select style="padding:7px;border:1.5px solid var(--borde);border-radius:8px;font-size:13px;"
          onchange="state.filtroEstado=this.value;state.pagina=1;renderListaSolicitudes();renderStatsCompact()">
          ${["Todos","Ingresada","Derivada","En Proceso","Respondida","Devuelta","Cerrada"]
            .map(e=>`<option ${state.filtroEstado===e?'selected':''}>${e}</option>`).join("")}
        </select>
        <select style="padding:7px;border:1.5px solid var(--borde);border-radius:8px;font-size:13px;"
          onchange="state.filtroUnidad=this.value;state.pagina=1;renderListaSolicitudes()">
          <option>Todas</option>
          ${CONFIG.unidades.map(u=>`<option ${state.filtroUnidad===u?'selected':''}>${u}</option>`).join("")}
        </select>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;">
        <input type="date" value="${state.filtroDesde}" style="padding:7px;border:1.5px solid var(--borde);border-radius:8px;font-size:12px;"
          onchange="state.filtroDesde=this.value;state.pagina=1;renderListaSolicitudes()">
        <input type="date" value="${state.filtroHasta}" style="padding:7px;border:1.5px solid var(--borde);border-radius:8px;font-size:12px;"
          onchange="state.filtroHasta=this.value;state.pagina=1;renderListaSolicitudes()">
      </div>
    </div>`;
}

function renderStatsCompact() {
  const cont = document.getElementById("panel-stats");
  if (!cont) return;
  const counts = {};
  getSolicitudesFiltradas().forEach(s => { counts[s.Estado] = (counts[s.Estado]||0)+1; });
  const total = getSolicitudesFiltradas().length;
  cont.innerHTML = `
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;">
      ${[
        {e:"Ingresada",color:"#dbeafe",tc:"#1d4ed8",icon:"📥"},
        {e:"Derivada",color:"#fef3c7",tc:"#b45309",icon:"📤"},
        {e:"En Proceso",color:"#cffafe",tc:"#0e7490",icon:"⚙️"},
        {e:"Respondida",color:"#dcfce7",tc:"#15803d",icon:"✅"},
        {e:"Devuelta",color:"#fee2e2",tc:"#b91c1c",icon:"↩️"},
        {e:"Cerrada",color:"#f3f4f6",tc:"#4b5563",icon:"🔒"}
      ].map(({e,color,tc,icon})=>`
        <div onclick="state.filtroEstado='${e}';state.pagina=1;renderListaSolicitudes();renderStatsCompact()"
          style="background:${color};border-radius:8px;padding:8px;text-align:center;cursor:pointer;">
          <div style="font-size:11px;">${icon}</div>
          <div style="font-size:18px;font-weight:700;color:${tc};line-height:1.2;">${counts[e]||0}</div>
          <div style="font-size:10px;color:${tc};">${e}</div>
        </div>`).join("")}
    </div>
    <div style="text-align:center;font-size:12px;color:#9ca3af;margin-top:6px;">Total: ${total}</div>`;
}

function getSolicitudesFiltradas() {
  return state.solicitudes.filter(s => {
    if (state.filtroEstado !== "Todos" && s.Estado !== state.filtroEstado) return false;
    if (state.filtroUnidad !== "Todas" && s.UnidadDerivada !== state.filtroUnidad) return false;
    if (state.filtroBuscar) {
      const q = state.filtroBuscar.toLowerCase();
      if (!s.NroSolicitud?.toLowerCase().includes(q) &&
          !s.Solicitante?.toLowerCase().includes(q) &&
          !s.Direccion?.toLowerCase().includes(q) &&
          !s.Solicitud?.toLowerCase().includes(q)) return false;
    }
    if (state.filtroDesde && new Date(s.FechaRecepcion) < new Date(state.filtroDesde)) return false;
    if (state.filtroHasta && new Date(s.FechaRecepcion) > new Date(state.filtroHasta)) return false;
    return true;
  });
}

function renderListaSolicitudes() {
  const cont = document.getElementById("lista-solicitudes");
  const filtradas = getSolicitudesFiltradas();
  const total = filtradas.length;
  const inicio = (state.pagina - 1) * state.pageSize;
  const pagina = filtradas.slice(inicio, inicio + state.pageSize);

  cont.innerHTML = pagina.length === 0 ? `<div style="text-align:center;padding:40px;color:#9ca3af;">No hay solicitudes</div>` :
    pagina.map(s => `
      <div class="sol-card ${state.solicitudSeleccionada?.id === s.id ? 'selected' : ''}" onclick="seleccionarSolicitud('${s.id}')">
        <div class="sol-card-top">
          <span class="sol-nro">${s.NroSolicitud}</span>
          <span class="estado-badge estado-${s.Estado}">${s.Estado}</span>
        </div>
        <div class="sol-card-name">${s.Solicitante}</div>
        <div class="sol-card-dir">📍 ${s.Direccion || ""}</div>
      </div>`).join("");

  const counter = document.getElementById("sec-lista-count");
  if (counter) counter.textContent = `(${total})`;

  const pag = document.getElementById("paginacion-solicitudes");
  if (pag) {
    const totalPags = Math.ceil(total / state.pageSize);
    pag.innerHTML = `
      <button onclick="cambiarPagina(${state.pagina - 1})" ${state.pagina <= 1 ? 'disabled' : ''}>‹ Anterior</button>
      <span>${state.pagina}/${Math.max(1, totalPags)}</span>
      <button onclick="cambiarPagina(${state.pagina + 1})" ${state.pagina >= totalPags ? 'disabled' : ''}>Siguiente ›</button>`;
  }
}

function cambiarPagina(p) {
  state.pagina = p;
  renderListaSolicitudes();
}

function renderFiltros() {
  const cont = document.getElementById("panel-filtros");
  const counts = {};
  state.solicitudes.forEach(s => { counts[s.Estado] = (counts[s.Estado] || 0) + 1; });

  cont.innerHTML = `
    <div class="filtros-grid">
      <div class="form-group full">
        <label>Buscar</label>
        <input type="text" placeholder="Nro, solicitante, dirección..." value="${state.filtroBuscar}"
          oninput="state.filtroBuscar=this.value;state.pagina=1;renderListaSolicitudes()">
      </div>
      <div class="form-group">
        <label>Estado</label>
        <select onchange="state.filtroEstado=this.value;state.pagina=1;renderListaSolicitudes()">
          ${["Todos","Ingresada","Derivada","En Proceso","Respondida","Devuelta","Cerrada"]
            .map(e => `<option ${state.filtroEstado===e?'selected':''}>${e}</option>`).join("")}
        </select>
      </div>
      <div class="form-group">
        <label>Unidad</label>
        <select onchange="state.filtroUnidad=this.value;state.pagina=1;renderListaSolicitudes()">
          <option>Todas</option>
          ${CONFIG.unidades.map(u => `<option ${state.filtroUnidad===u?'selected':''}>${u}</option>`).join("")}
        </select>
      </div>
      <div class="form-group">
        <label>Desde</label>
        <input type="date" value="${state.filtroDesde}" onchange="state.filtroDesde=this.value;state.pagina=1;renderListaSolicitudes()">
      </div>
      <div class="form-group">
        <label>Hasta</label>
        <input type="date" value="${state.filtroHasta}" onchange="state.filtroHasta=this.value;state.pagina=1;renderListaSolicitudes()">
      </div>
    </div>
    <div class="stats-grid">
      ${[
        {e:"Ingresada",icon:"📥",color:"#dbeafe",num:"#1d4ed8"},
        {e:"Derivada",icon:"📤",color:"#fef3c7",num:"#b45309"},
        {e:"En Proceso",icon:"⚙️",color:"#cffafe",num:"#0e7490"},
        {e:"Respondida",icon:"✅",color:"#dcfce7",num:"#15803d"},
        {e:"Devuelta",icon:"↩️",color:"#fee2e2",num:"#b91c1c"},
        {e:"Cerrada",icon:"🔒",color:"#f3f4f6",num:"#4b5563"}
      ].map(({e,icon,color,num}) => `
        <div class="stat-card" onclick="state.filtroEstado='${e}';state.pagina=1;renderListaSolicitudes()" title="Filtrar por ${e}">
          <div class="stat-icon" style="background:${color};">${icon}</div>
          <div class="stat-info">
            <div class="stat-num" style="color:${num}">${counts[e]||0}</div>
            <div class="stat-label">${e}</div>
          </div>
        </div>`).join("")}
    </div>`;
}

function renderFormNueva() {
  const cont = document.getElementById("panel-nueva");
  cont.innerHTML = `
    <div class="form-nueva">

      <!-- Sección datos -->
      <div class="form-section">
        <div class="form-section-header">📋 Datos de la Solicitud</div>
        <div class="form-section-body">
        <div class="form-row">
          <div class="form-group">
            <label>* Nro Solicitud</label>
            <input type="text" id="nueva-nro" placeholder="Ej: 16-157" style="font-weight:600;">
          </div>
          <div class="form-group">
            <label>* Fecha Recepción</label>
            <input type="date" id="nueva-fecha" value="${new Date().toISOString().split('T')[0]}">
          </div>
        </div>
        <div class="form-group">
          <label>* Nombre Solicitante</label>
          <input type="text" id="nueva-solicitante" placeholder="Nombre completo del solicitante">
        </div>
        <div class="form-group">
          <label>* Dirección</label>
          <input type="text" id="nueva-dir" placeholder="Calle, número, villa/sector">
        </div>
        <div class="form-group">
          <label>Descripción de la solicitud</label>
          <textarea id="nueva-solicitud" rows="3" placeholder="Resumen del motivo de la solicitud..."></textarea>
        </div>
        </div><!-- /form-section-body -->
      </div><!-- /form-section -->

      <!-- Sección adjunto -->
      <div class="form-section">
        <div class="form-section-header naranja">📎 Documento Adjunto</div>
        <div class="form-section-body">
        <div id="drop-area" class="upload-area"
          onclick="document.getElementById('nueva-files').click()"
          ondragover="event.preventDefault();this.style.background='#fef3c7'"
          ondragleave="this.style.background=''"
          ondrop="event.preventDefault();handleFiles(event.dataTransfer.files)">
          <div style="font-size:36px;margin-bottom:8px;">📄</div>
          <div style="font-weight:600;margin-bottom:4px;">Arrastra el PDF aquí</div>
          <div style="font-size:12px;color:#9ca3af;">o haz clic para seleccionar — PDF, JPG, PNG</div>
        </div>
        <input type="file" id="nueva-files" multiple accept=".pdf,.jpg,.jpeg,.png" onchange="handleFiles(this.files)">
        <div id="file-list" class="file-list"></div>
        <div id="pdf-preview" style="margin-top:10px;"></div>
        </div>
      </div>

      <!-- Botones centrados -->
      <div style="display:flex;justify-content:center;gap:12px;padding:4px 0 8px;">
        <button class="btn-primary" onclick="guardarSolicitud()" style="width:220px;padding:13px;">
          💾 Guardar Solicitud
        </button>
        <button onclick="limpiarFormNueva()" style="padding:10px 20px;border:1.5px solid var(--borde);border-radius:8px;background:white;cursor:pointer;font-size:13px;color:#666;">
          🗑 Limpiar
        </button>
      </div>
    </div>`;

  // Drag & drop visual en toda la zona
  const dropArea = document.getElementById("drop-area");
  if (dropArea) {
    dropArea.addEventListener("dragover", e => { e.preventDefault(); dropArea.style.background = "#fef3c7"; });
    dropArea.addEventListener("dragleave", () => { dropArea.style.background = ""; });
    dropArea.addEventListener("drop", e => { e.preventDefault(); dropArea.style.background = ""; handleFiles(e.dataTransfer.files); });
  }
}

function handleFiles(files) {
  state.adjuntosNueva = [...state.adjuntosNueva, ...Array.from(files)];
  updateFileList();
}

function updateFileList() {
  const list = document.getElementById("file-list");
  const area = document.getElementById("drop-area");
  const preview = document.getElementById("pdf-preview");
  if (!list) return;

  list.innerHTML = state.adjuntosNueva.map((f, i) => {
    const isPdf = f.name.toLowerCase().endsWith('.pdf');
    const isImg = /\.(jpg|jpeg|png)$/i.test(f.name);
    const icon = isPdf ? '📄' : isImg ? '🖼️' : '📎';
    return `
    <div class="file-item" style="border-left:3px solid ${isPdf?'#ef4444':isImg?'#3b82f6':'#6b7280'};">
      <span>${icon} <strong>${f.name}</strong> <span style="color:#9ca3af">(${(f.size/1024).toFixed(0)} KB)</span></span>
      <button onclick="removeFile(${i})" title="Quitar">✕</button>
    </div>`;
  }).join("");

  // Preview del primer PDF
  if (preview) {
    // Mostrar en visor derecho (no inline en formulario)
    const firstFile = state.adjuntosNueva[0];
    if (firstFile) {
      const url = URL.createObjectURL(firstFile);
      const isPdf = firstFile.name.toLowerCase().endsWith('.pdf');
      mostrarEnVisor(url, firstFile.name, isPdf, "Nueva");
      preview.innerHTML = ""; // No mostrar inline
    } else {
      preview.innerHTML = "";
      // Limpiar visor si no hay archivos
      const pdfPanel = document.getElementById("pdf-visor-contenido");
      if (pdfPanel) pdfPanel.innerHTML = `<div class="pdf-visor-empty"><span>📄</span><p>Selecciona una solicitud para ver el documento adjunto</p></div>`;
    }
  }

  if (area) {
    area.style.display = state.adjuntosNueva.length > 0 ? "none" : "flex";
    area.style.flexDirection = "column";
    area.style.alignItems = "center";
  }
}

function removeFile(idx) {
  state.adjuntosNueva.splice(idx, 1);
  updateFileList();
}

function limpiarFormNueva() {
  state.adjuntosNueva = [];
  renderFormNueva();
}

async function guardarSolicitud() {
  const nro = document.getElementById("nueva-nro")?.value.trim();
  const fecha = document.getElementById("nueva-fecha")?.value;
  const sol = document.getElementById("nueva-solicitante")?.value.trim();
  const dir = document.getElementById("nueva-dir")?.value.trim();
  const desc = document.getElementById("nueva-solicitud")?.value.trim();

  if (!nro || !fecha || !sol || !dir) {
    showToast("error", "Completa los campos obligatorios (*)");
    return;
  }

  // Check duplicate
  const existe = state.solicitudes.find(s => s.NroSolicitud === nro);
  if (existe) {
    showToast("error", `Ya existe la solicitud ${nro}`);
    return;
  }

  showLoading("Guardando solicitud...");
  try {
    const item = await crearSolicitud({
      NroSolicitud: nro,
      FechaRecepcion: fecha,
      Solicitante: sol,
      Direccion: dir,
      Solicitud: desc,
      Estado: CONFIG.estados.INGRESADA
    });

    // Upload attachments
    if (state.adjuntosNueva.length > 0) {
      for (const file of state.adjuntosNueva) {
        await uploadAttachment(CONFIG.lists.solicitudes, item.id, file).catch(console.error);
      }
    }

    // Historial
    await registrarHistorial({
      NroSolicitud: nro,
      Accion: "Ingreso de solicitud",
      EstadoAnterior: "",
      EstadoNuevo: CONFIG.estados.INGRESADA,
      UsuarioAccion: state.usuario.NombreCompleto,
      RolUsuario: state.usuario.Rol,
      Unidad: state.usuario.Unidad,
      FechaAccion: new Date().toISOString(),
      Observaciones: desc
    });

    // Notificar director
    await notificarDirector(item, "Nueva solicitud ingresada").catch(console.error);

    state.adjuntosNueva = [];
    showToast("success", `✅ Solicitud ${nro} guardada correctamente`);
    await renderSecretaria();
  } catch (e) {
    showToast("error", "Error al guardar: " + e.message);
  } finally {
    hideLoading();
  }
}

// ===== DIRECTOR VIEW =====
async function renderDirector() {
  showLoading("Cargando solicitudes...");
  try {
    state.solicitudes = await getSolicitudes();
    state.solicitudes.sort((a, b) => new Date(b.FechaRecepcion) - new Date(a.FechaRecepcion));
    renderSidebarEstados();
    updateTabBadges();
  } catch (e) {
    showToast("error", "Error: " + e.message);
  } finally {
    hideLoading();
  }
}

function renderSidebarEstados() {
  const counts = {};
  state.solicitudes.forEach(s => { counts[s.Estado] = (counts[s.Estado] || 0) + 1; });

  const sidebar = document.getElementById("dir-sidebar");
  sidebar.innerHTML = `
    <div class="sidebar-estados">
      ${["Todos","Ingresada","Derivada","En Proceso","Respondida","Devuelta","Cerrada"].map(e => `
        <button class="sidebar-btn ${state.filtroEstado === e ? 'active' : ''}" onclick="filtrarDirector('${e}')">
          <span>${e}</span>
          <span class="cnt">${e === "Todos" ? state.solicitudes.length : (counts[e]||0)}</span>
        </button>`).join("")}
    </div>`;

  renderListaDirector();
}

function filtrarDirector(estado) {
  state.filtroEstado = estado;
  state.pagina = 1;
  state.solicitudSeleccionada = null;
  renderSidebarEstados();
  renderDetalleDirector(null);
  renderPDFViewer(null);
}

function renderListaDirector() {
  const filtradas = state.solicitudes.filter(s =>
    state.filtroEstado === "Todos" || s.Estado === state.filtroEstado
  );
  const inicio = (state.pagina - 1) * state.pageSize;
  const pagina = filtradas.slice(inicio, inicio + state.pageSize);
  const totalPags = Math.ceil(filtradas.length / state.pageSize);

  const cont = document.getElementById("dir-lista");
  cont.innerHTML = `
    <div style="padding:12px 12px 0;">
      <input type="text" placeholder="🔍 Buscar solicitud..." style="width:100%;padding:10px;border:1.5px solid #dde3ee;border-radius:8px;font-size:14px;"
        oninput="state.filtroBuscar=this.value;state.pagina=1;renderListaDirector()">
    </div>
    <div style="flex:1;overflow-y:auto;padding:12px;">
      ${pagina.length === 0 ? '<p style="text-align:center;color:#9ca3af;padding:40px">Sin solicitudes</p>' :
        pagina.map(s => `
          <div class="sol-card ${state.solicitudSeleccionada?.id === s.id ? 'selected' : ''}" onclick="seleccionarSolicitudDirector('${s.id}')">
            <div class="sol-card-top">
              <span class="sol-nro">${s.NroSolicitud}</span>
              <span class="estado-badge estado-${s.Estado}">${s.Estado}</span>
            </div>
            <div class="sol-card-name">${s.Solicitante}</div>
            <div class="sol-card-dir">📍 ${s.Direccion || ""}</div>
            ${s.UnidadDerivada ? `<div style="font-size:12px;color:#888;margin-top:2px;">→ ${s.UnidadDerivada}</div>` : ""}
          </div>`).join("")}
    </div>
    <div class="paginacion">
      <button onclick="cambiarPaginaDir(${state.pagina-1})" ${state.pagina<=1?'disabled':''}>‹</button>
      <span>${state.pagina}/${Math.max(1,totalPags)} (${filtradas.length})</span>
      <button onclick="cambiarPaginaDir(${state.pagina+1})" ${state.pagina>=totalPags?'disabled':''}>›</button>
    </div>`;
}

function cambiarPaginaDir(p) {
  state.pagina = p;
  renderListaDirector();
}

async function seleccionarSolicitudDirector(id) {
  state.solicitudSeleccionada = state.solicitudes.find(s => s.id === id);
  renderListaDirector();
  renderPDFViewer(state.solicitudSeleccionada);
  renderDetalleDirector(state.solicitudSeleccionada);
}

async function renderPDFViewer(sol) {
  const cont = document.getElementById("dir-pdf");
  if (!sol) {
    cont.innerHTML = `<div class="pdf-placeholder"><span style="font-size:60px">📄</span><p>Selecciona una solicitud para ver el documento</p></div>`;
    return;
  }
  cont.innerHTML = `<div style="text-align:center;padding:20px;color:#666">Cargando adjuntos...</div>`;
  try {
    const attachments = await getListItemAttachments(CONFIG.lists.solicitudes, sol.id);
    if (attachments.length === 0) {
      cont.innerHTML = `<div class="pdf-placeholder"><span style="font-size:60px">📎</span><p>Sin documentos adjuntos</p></div>`;
      return;
    }
    const pdfs = attachments.filter(a => a.name?.toLowerCase().endsWith('.pdf'));
    const imgs = attachments.filter(a => /\.(jpg|jpeg|png|gif)$/i.test(a.name));
    const token = await getToken();

    cont.innerHTML = `
      ${pdfs.length > 0 ? `
        <div style="height:60%;min-height:300px;margin-bottom:12px;">
          <iframe src="${pdfs[0]['@microsoft.graph.downloadUrl']}" style="width:100%;height:100%;border:none;border-radius:8px;"></iframe>
        </div>` : ""}
      ${imgs.length > 0 ? `
        <div class="attachments-grid">
          ${imgs.map(img => `
            <div class="attach-thumb" onclick="window.open('${img['@microsoft.graph.downloadUrl']}','_blank')">
              <img src="${img['@microsoft.graph.downloadUrl']}" alt="${img.name}">
            </div>`).join("")}
        </div>` : ""}
      ${pdfs.length === 0 && imgs.length === 0 ? `<div class="pdf-placeholder"><span>📎</span><p>${attachments.map(a=>a.name).join(', ')}</p></div>` : ""}`;
  } catch (e) {
    cont.innerHTML = `<div class="pdf-placeholder"><span>⚠️</span><p>Error al cargar adjuntos</p></div>`;
  }
}

function renderDetalleDirector(sol) {
  const cont = document.getElementById("dir-detalle");
  if (!sol) { cont.innerHTML = `<div style="text-align:center;color:#9ca3af;padding:40px;">Selecciona una solicitud</div>`; return; }

  // Director puede derivar si está Ingresada o Devuelta (para re-derivar)
  const esDerivable = sol.Estado === CONFIG.estados.INGRESADA || sol.Estado === CONFIG.estados.DEVUELTA;
  // Solo Director cierra, solo cuando está Respondida
  const esCerrable = sol.Estado === CONFIG.estados.RESPONDIDA && state.usuario.Rol === CONFIG.roles.DIRECTOR;

  cont.innerHTML = `
    <div class="panel-header" style="background:#f8fafc;">📋 Detalle Solicitud</div>
    <div style="flex:1;overflow-y:auto;">
      <div class="detalle-fields">
        <div class="detalle-row">
          <div class="detalle-label">🔢 Nro</div>
          <div class="detalle-value"><strong>${sol.NroSolicitud}</strong></div>
        </div>
        <div class="detalle-row">
          <div class="detalle-label">📅 Fecha</div>
          <div class="detalle-value">${formatFecha(sol.FechaRecepcion)}</div>
        </div>
        <div class="detalle-row">
          <div class="detalle-label">👤 Solicitante</div>
          <div class="detalle-value">${sol.Solicitante}</div>
        </div>
        <div class="detalle-row">
          <div class="detalle-label">📍 Dirección</div>
          <div class="detalle-value">${sol.Direccion || "-"}</div>
        </div>
        <div class="detalle-row">
          <div class="detalle-label">📝 Solicitud</div>
          <div class="detalle-value">${sol.Solicitud || "-"}</div>
        </div>
        <div class="detalle-row">
          <div class="detalle-label">🏷 Estado</div>
          <div class="detalle-value"><span class="estado-badge estado-${sol.Estado}">${sol.Estado}</span></div>
        </div>
        ${sol.UnidadDerivada ? `<div class="detalle-row">
          <div class="detalle-label">🏢 Unidad</div>
          <div class="detalle-value">${sol.UnidadDerivada}</div>
        </div>` : ""}
        ${sol.MotivoDevolucion ? `<div class="detalle-row">
          <div class="detalle-label">↩️ Devolución</div>
          <div class="detalle-value" style="color:#b91c1c">${sol.MotivoDevolucion}</div>
        </div>` : ""}
      </div>

      ${esDerivable ? `
        <div class="acciones-panel">
          <div class="section-title">Derivar Solicitud</div>
          <select id="dir-unidad-derivar">
            <option value="">— Seleccionar unidad —</option>
            ${CONFIG.unidades.map(u => `<option>${u}</option>`).join("")}
          </select>
          <textarea id="dir-accion-obs" rows="2" placeholder="Observaciones / instrucciones para la unidad..."></textarea>
          <button class="btn-primary" onclick="derivarSolicitud('${sol.id}')">📤 Derivar</button>
        </div>` : ""}

      ${esCerrable ? `
        <div class="acciones-panel">
          <div class="section-title">Cerrar Solicitud</div>
          <textarea id="dir-cierre-obs" rows="2" placeholder="Observaciones de cierre..."></textarea>
          <button class="btn-success" onclick="cerrarSolicitud('${sol.id}')">🔒 Cerrar Solicitud</button>
        </div>` : ""}

      <div style="padding:12px;">
        <button class="btn-primary" style="background:#6b7280;width:100%;" onclick="verHistorial('${sol.NroSolicitud}')">🕐 Ver Historial</button>
      </div>
    </div>`;
}

async function derivarSolicitud(solId) {
  const unidad = document.getElementById("dir-unidad-derivar")?.value;
  const obs = document.getElementById("dir-accion-obs")?.value.trim();
  if (!unidad) { showToast("error", "Selecciona una unidad"); return; }

  showLoading("Derivando solicitud...");
  try {
    const sol = state.solicitudes.find(s => s.id === solId);
    await actualizarSolicitud(solId, { Estado: CONFIG.estados.DERIVADA, UnidadDerivada: unidad });
    await registrarHistorial({
      NroSolicitud: sol.NroSolicitud,
      Accion: "Derivada a unidad",
      EstadoAnterior: sol.Estado,
      EstadoNuevo: CONFIG.estados.DERIVADA,
      UsuarioAccion: state.usuario.NombreCompleto,
      RolUsuario: state.usuario.Rol,
      Unidad: unidad,
      FechaAccion: new Date().toISOString(),
      Observaciones: obs
    });
    await notificarUnidad({ ...sol, Estado: CONFIG.estados.DERIVADA }, unidad).catch(console.error);
    showToast("success", `✅ Derivada a ${unidad}`);
    await renderDirector();
    renderSidebarEstados();
  } catch (e) {
    showToast("error", "Error: " + e.message);
  } finally {
    hideLoading();
  }
}

async function cerrarSolicitud(solId) {
  const obs = document.getElementById("dir-cierre-obs")?.value.trim();
  showLoading("Cerrando solicitud...");
  try {
    const sol = state.solicitudes.find(s => s.id === solId);
    await actualizarSolicitud(solId, { Estado: CONFIG.estados.CERRADA });
    await registrarHistorial({
      NroSolicitud: sol.NroSolicitud,
      Accion: "Solicitud cerrada",
      EstadoAnterior: sol.Estado,
      EstadoNuevo: CONFIG.estados.CERRADA,
      UsuarioAccion: state.usuario.NombreCompleto,
      RolUsuario: state.usuario.Rol,
      Unidad: state.usuario.Unidad,
      FechaAccion: new Date().toISOString(),
      Observaciones: obs
    });
    showToast("success", "🔒 Solicitud cerrada");
    await renderDirector();
  } catch (e) {
    showToast("error", "Error: " + e.message);
  } finally {
    hideLoading();
  }
}

// ===== UNIDAD VIEW =====
// Orden de estados para mostrar más urgentes primero
const ORDEN_ESTADO = { "Devuelta": 0, "Derivada": 1, "En Proceso": 2, "Respondida": 3, "Cerrada": 4, "Ingresada": 5 };

function ordenarSolicitudes(lista) {
  return lista.sort((a, b) => {
    const oa = ORDEN_ESTADO[a.Estado] ?? 9;
    const ob = ORDEN_ESTADO[b.Estado] ?? 9;
    if (oa !== ob) return oa - ob;
    return new Date(b.FechaRecepcion) - new Date(a.FechaRecepcion);
  });
}

async function renderUnidad() {
  showLoading("Cargando...");
  try {
    const all = await getSolicitudes();
    state.solicitudes = all.filter(s => s.UnidadDerivada === state.usuario.Unidad);
    // Orden: Devuelta > Derivada > En Proceso > Respondida > Cerrada
    state.solicitudes = ordenarSolicitudes(state.solicitudes);
    renderSidebarUnidad();
    renderDetalleUnidad(null);
    updateTabBadges();
  } catch (e) {
    showToast("error", "Error: " + e.message);
  } finally {
    hideLoading();
  }
}

function renderSidebarUnidad() {
  const counts = {};
  state.solicitudes.forEach(s => { counts[s.Estado] = (counts[s.Estado] || 0) + 1; });

  const estados = ["Todos","Derivada","En Proceso","Respondida","Devuelta"];
  const filtradas = state.filtroEstado === "Todos" ? state.solicitudes :
    state.solicitudes.filter(s => s.Estado === state.filtroEstado);
  const inicio = (state.pagina - 1) * state.pageSize;
  const pagina = filtradas.slice(inicio, inicio + state.pageSize);
  const totalPags = Math.ceil(filtradas.length / state.pageSize);

  const cont = document.getElementById("uni-lista");
  cont.innerHTML = `
    <div style="padding:12px 12px 0;">
      <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px;">
        ${estados.map(e => `
          <button onclick="filtrarUnidad('${e}')" style="padding:6px 12px;border-radius:16px;border:1.5px solid ${state.filtroEstado===e?'var(--azul)':'var(--borde)'};background:${state.filtroEstado===e?'var(--azul)':'white'};color:${state.filtroEstado===e?'white':'var(--texto)'};cursor:pointer;font-size:13px;">
            ${e} ${e!=='Todos'?`(${counts[e]||0})`:`(${state.solicitudes.length})`}
          </button>`).join("")}
      </div>
      <input type="text" placeholder="🔍 Buscar..." style="width:100%;padding:9px;border:1.5px solid #dde3ee;border-radius:8px;font-size:13px;"
        oninput="state.filtroBuscar=this.value;state.pagina=1;renderSidebarUnidad()">
    </div>
    <div style="flex:1;overflow-y:auto;padding:12px;">
      ${pagina.length === 0 ? '<p style="text-align:center;color:#9ca3af;padding:40px">Sin solicitudes</p>' :
        pagina.map(s => `
          <div class="sol-card ${state.solicitudSeleccionada?.id === s.id ? 'selected' : ''}" onclick="seleccionarSolicitudUnidad('${s.id}')">
            <div class="sol-card-top">
              <span class="sol-nro">${s.NroSolicitud}</span>
              <span class="estado-badge estado-${s.Estado}">${s.Estado}</span>
            </div>
            <div class="sol-card-name">${s.Solicitante}</div>
            <div class="sol-card-dir">📍 ${s.Direccion || ""}</div>
          </div>`).join("")}
    </div>
    <div class="paginacion">
      <button onclick="cambiarPaginaUni(${state.pagina-1})" ${state.pagina<=1?'disabled':''}>‹</button>
      <span>${state.pagina}/${Math.max(1,totalPags)} (${filtradas.length})</span>
      <button onclick="cambiarPaginaUni(${state.pagina+1})" ${state.pagina>=totalPags?'disabled':''}>›</button>
    </div>`;
}

function filtrarUnidad(estado) {
  state.filtroEstado = estado;
  state.pagina = 1;
  renderSidebarUnidad();
}
function cambiarPaginaUni(p) { state.pagina = p; renderSidebarUnidad(); }

async function seleccionarSolicitudUnidad(id) {
  state.solicitudSeleccionada = state.solicitudes.find(s => s.id === id);
  renderSidebarUnidad();
  renderDetalleUnidad(state.solicitudSeleccionada);
}

async function renderDetalleUnidad(sol) {
  const cont = document.getElementById("uni-detalle");
  if (!sol) {
    cont.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#9ca3af;flex-direction:column;gap:16px;"><span style="font-size:60px">📋</span><p>Selecciona una solicitud</p></div>`;
    return;
  }

  // Load attachments
  let attachHtml = '<p style="color:#9ca3af;font-size:13px;">Cargando adjuntos...</p>';
  const evidencias = await getEvidenciasBySolicitud(sol.NroSolicitud).catch(() => []);

  try {
    const attachments = await getListItemAttachments(CONFIG.lists.solicitudes, sol.id);
    const pdfs = attachments.filter(a => a.name?.toLowerCase().endsWith('.pdf'));
    const imgs = attachments.filter(a => /\.(jpg|jpeg|png|gif)$/i.test(a.name));
    attachHtml = pdfs.length > 0 ?
      `<iframe src="${pdfs[0]['@microsoft.graph.downloadUrl']}" style="width:100%;height:300px;border:none;border-radius:8px;"></iframe>` :
      imgs.length > 0 ?
      `<div class="attachments-grid">${imgs.map(i=>`<div class="attach-thumb" onclick="window.open('${i['@microsoft.graph.downloadUrl']}','_blank')"><img src="${i['@microsoft.graph.downloadUrl']}"></div>`).join("")}</div>` :
      `<p style="color:#9ca3af;font-size:13px;text-align:center;padding:20px">Sin adjuntos</p>`;
  } catch {}

  const puedeCerrar = state.usuario.PuedeCerrar && sol.Estado === CONFIG.estados.RESPONDIDA;

  cont.innerHTML = `
    <div style="overflow-y:auto;height:100%;display:flex;flex-direction:column;gap:0;">
      <div class="panel-header" style="background:#f8fafc;border-bottom:1px solid var(--borde);">📋 ${sol.NroSolicitud} — ${sol.Solicitante}</div>
      <div style="padding:12px;background:#f8fafc;border-bottom:1px solid var(--borde);">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:13px;">
          <div><span style="color:#666">Fecha: </span>${formatFecha(sol.FechaRecepcion)}</div>
          <div><span style="color:#666">Estado: </span><span class="estado-badge estado-${sol.Estado}">${sol.Estado}</span></div>
          <div style="grid-column:1/-1;"><span style="color:#666">Dirección: </span>${sol.Direccion||"-"}</div>
          <div style="grid-column:1/-1;"><span style="color:#666">Solicitud: </span>${sol.Solicitud||"-"}</div>
        </div>
      </div>
      <div style="padding:12px;border-bottom:1px solid var(--borde);">${attachHtml}</div>

      ${evidencias.length > 0 ? `
      <div style="padding:12px;border-bottom:1px solid var(--borde);">
        <div class="section-title">Evidencias registradas</div>
        ${evidencias.map(e=>`
          <div style="background:#f8fafc;border-radius:8px;padding:10px;margin-top:6px;font-size:13px;">
            <div style="font-weight:600;color:var(--azul)">${e.Responsable} — ${formatFecha(e.FechaCarga)}</div>
            <div>${e.DescripcionEvidencia}</div>
          </div>`).join("")}
      </div>` : ""}

      <div class="acciones-panel" style="flex:1;">
        <div class="section-title">Registrar acción</div>
        <textarea id="uni-obs" rows="3" placeholder="Observaciones / descripción de la acción realizada..."></textarea>
        <div class="form-group">
          <label>Adjuntar evidencia (fotos/docs)</label>
          <div class="upload-area" onclick="document.getElementById('uni-ev-files').click()">📎 Adjuntar evidencia</div>
          <input type="file" id="uni-ev-files" multiple accept=".pdf,.jpg,.jpeg,.png">
        </div>
        <div class="btn-row">
          <button class="btn-success" onclick="responderSolicitud('${sol.id}')">✅ Responder</button>
          <button class="btn-primary" style="background:var(--en-proceso)!important;" onclick="enProcesoSolicitud('${sol.id}')">⚙️ En Proceso</button>
        </div>
        <div class="btn-row">
          <button class="btn-warning" onclick="devolverSolicitudUnidad('${sol.id}')">↩️ Devolver</button>
          ${puedeCerrar ? `<button class="btn-danger" onclick="cerrarSolicitudUnidad('${sol.id}')">🔒 Cerrar</button>` : ""}
        </div>
        <button class="btn-primary" style="background:#6b7280;" onclick="verHistorial('${sol.NroSolicitud}')">🕐 Historial</button>
      </div>
    </div>`;
}

async function responderSolicitud(solId) {
  const obs = document.getElementById("uni-obs")?.value.trim();
  const files = document.getElementById("uni-ev-files")?.files;
  if (!obs) { showToast("error", "Ingresa una observación"); return; }

  showLoading("Respondiendo...");
  try {
    const sol = state.solicitudes.find(s => s.id === solId);
    await actualizarSolicitud(solId, { Estado: CONFIG.estados.RESPONDIDA });

    const evItem = await crearEvidencia({
      NroSolicitud: sol.NroSolicitud,
      SolicitudID: parseInt(sol.id),
      Unidad: state.usuario.Unidad,
      DescripcionEvidencia: obs,
      FechaCarga: new Date().toISOString(),
      Responsable: state.usuario.NombreCompleto
    });

    if (files?.length > 0) {
      for (const f of files) {
        await uploadAttachment(CONFIG.lists.evidencias, evItem.id, f).catch(console.error);
      }
    }

    await registrarHistorial({
      NroSolicitud: sol.NroSolicitud,
      Accion: "Solicitud respondida",
      EstadoAnterior: sol.Estado,
      EstadoNuevo: CONFIG.estados.RESPONDIDA,
      UsuarioAccion: state.usuario.NombreCompleto,
      RolUsuario: state.usuario.Rol,
      Unidad: state.usuario.Unidad,
      FechaAccion: new Date().toISOString(),
      Observaciones: obs
    });

    await notificarDirector({ ...sol, Estado: CONFIG.estados.RESPONDIDA }, "Solicitud respondida por unidad").catch(console.error);
    showToast("success", "✅ Solicitud respondida");
    await renderUnidad();
  } catch (e) {
    showToast("error", "Error: " + e.message);
  } finally {
    hideLoading();
  }
}

async function enProcesoSolicitud(solId) {
  const obs = document.getElementById("uni-obs")?.value.trim();
  showLoading("Actualizando...");
  try {
    const sol = state.solicitudes.find(s => s.id === solId);
    await actualizarSolicitud(solId, { Estado: CONFIG.estados.EN_PROCESO });
    await registrarHistorial({
      NroSolicitud: sol.NroSolicitud,
      Accion: "Solicitud en proceso",
      EstadoAnterior: sol.Estado,
      EstadoNuevo: CONFIG.estados.EN_PROCESO,
      UsuarioAccion: state.usuario.NombreCompleto,
      RolUsuario: state.usuario.Rol,
      Unidad: state.usuario.Unidad,
      FechaAccion: new Date().toISOString(),
      Observaciones: obs
    });
    showToast("info", "⚙️ Solicitud marcada En Proceso");
    await renderUnidad();
  } catch (e) {
    showToast("error", "Error: " + e.message);
  } finally {
    hideLoading();
  }
}

async function devolverSolicitudUnidad(solId) {
  const obs = document.getElementById("uni-obs")?.value.trim();
  if (!obs) { showToast("error", "Ingresa el motivo de devolución"); return; }

  showLoading("Devolviendo...");
  try {
    const sol = state.solicitudes.find(s => s.id === solId);
    await actualizarSolicitud(solId, { Estado: CONFIG.estados.DEVUELTA, MotivoDevolucion: obs });
    await registrarHistorial({
      NroSolicitud: sol.NroSolicitud,
      Accion: "Solicitud devuelta",
      EstadoAnterior: sol.Estado,
      EstadoNuevo: CONFIG.estados.DEVUELTA,
      UsuarioAccion: state.usuario.NombreCompleto,
      RolUsuario: state.usuario.Rol,
      Unidad: state.usuario.Unidad,
      FechaAccion: new Date().toISOString(),
      Observaciones: obs,
      Motivo: obs
    });
    await notificarDirector({ ...sol, Estado: CONFIG.estados.DEVUELTA }, "Solicitud devuelta por unidad").catch(console.error);
    showToast("info", "↩️ Solicitud devuelta");
    await renderUnidad();
  } catch (e) {
    showToast("error", "Error: " + e.message);
  } finally {
    hideLoading();
  }
}

async function cerrarSolicitudUnidad(solId) {
  const obs = document.getElementById("uni-obs")?.value.trim();
  showLoading("Cerrando...");
  try {
    const sol = state.solicitudes.find(s => s.id === solId);
    await actualizarSolicitud(solId, { Estado: CONFIG.estados.CERRADA });
    await registrarHistorial({
      NroSolicitud: sol.NroSolicitud,
      Accion: "Solicitud cerrada",
      EstadoAnterior: sol.Estado,
      EstadoNuevo: CONFIG.estados.CERRADA,
      UsuarioAccion: state.usuario.NombreCompleto,
      RolUsuario: state.usuario.Rol,
      Unidad: state.usuario.Unidad,
      FechaAccion: new Date().toISOString(),
      Observaciones: obs
    });
    showToast("success", "🔒 Solicitud cerrada");
    await renderUnidad();
  } catch (e) {
    showToast("error", "Error: " + e.message);
  } finally {
    hideLoading();
  }
}

// ===== HISTORIAL MODAL =====
async function verHistorial(nroSolicitud) {
  showLoading("Cargando historial...");
  const historial = await getHistorialBySolicitud(nroSolicitud).catch(() => []);
  hideLoading();
  historial.sort((a, b) => new Date(b.FechaAccion) - new Date(a.FechaAccion));

  const modal = document.createElement("div");
  modal.className = "modal-overlay";
  modal.innerHTML = `
    <div class="modal">
      <div class="modal-header">
        <h3>🕐 Historial — Solicitud ${nroSolicitud}</h3>
        <button class="modal-close" onclick="this.closest('.modal-overlay').remove()">✕</button>
      </div>
      <div class="modal-body">
        ${historial.length === 0 ? '<p style="color:#9ca3af;text-align:center">Sin historial registrado</p>' :
          historial.map(h => `
            <div class="historial-item accion-${h.Accion?.toLowerCase().split(' ')[1]||''}">
              <div class="historial-fecha">${formatFechaHora(h.FechaAccion)}</div>
              <div class="historial-accion">${h.Accion}</div>
              <div class="historial-user">👤 ${h.UsuarioAccion} (${h.RolUsuario}) ${h.Unidad?`— ${h.Unidad}`:""}</div>
              ${h.EstadoAnterior ? `<div style="font-size:12px;color:#888;margin-top:2px;">${h.EstadoAnterior} → ${h.EstadoNuevo}</div>` : ""}
              ${h.Observaciones ? `<div class="historial-obs">${h.Observaciones}</div>` : ""}
            </div>`).join("")}
      </div>
    </div>`;
  document.body.appendChild(modal);
}

// ===== GRAFICOS =====
async function renderGraficos() {
  const cont = document.getElementById("view-graficos");
  cont.innerHTML = `
    <div style="padding:16px 24px;background:white;border-bottom:1px solid var(--borde);display:flex;gap:16px;align-items:center;">
      <div class="form-group" style="flex:1;">
        <label>Período Desde</label>
        <input type="date" id="graf-desde" value="${new Date(new Date().setMonth(new Date().getMonth()-3)).toISOString().split('T')[0]}" onchange="actualizarGraficos()">
      </div>
      <div class="form-group" style="flex:1;">
        <label>Hasta</label>
        <input type="date" id="graf-hasta" value="${new Date().toISOString().split('T')[0]}" onchange="actualizarGraficos()">
      </div>
      <button class="btn-primary" style="margin-top:20px;" onclick="actualizarGraficos()">📊 Actualizar</button>
      <button class="btn-primary" style="margin-top:20px;background:#22c55e;" onclick="exportarExcel()">⬇️ Exportar</button>
    </div>
    <div class="charts-view" id="charts-container">
      <div class="chart-card"><h3>📊 Solicitudes por Estado</h3><canvas id="chart-estado" height="220"></canvas></div>
      <div class="chart-card"><h3>🏢 Solicitudes por Unidad</h3><canvas id="chart-unidad" height="220"></canvas></div>
      <div class="chart-card full"><h3>📅 Evolución por Mes</h3><canvas id="chart-mes" height="120"></canvas></div>
    </div>`;

  await actualizarGraficos();
}

async function actualizarGraficos() {
  showLoading("Generando reportes...");
  try {
    const desde = document.getElementById("graf-desde")?.value;
    const hasta = document.getElementById("graf-hasta")?.value;
    const all = await getSolicitudes();
    const filtradas = all.filter(s => {
      const f = new Date(s.FechaRecepcion);
      return (!desde || f >= new Date(desde)) && (!hasta || f <= new Date(hasta));
    });

    // Chart 1: por estado
    const byEstado = {};
    filtradas.forEach(s => { byEstado[s.Estado] = (byEstado[s.Estado] || 0) + 1; });
    renderChart("chart-estado", "doughnut", Object.keys(byEstado), Object.values(byEstado),
      ["#3b82f6","#f59e0b","#06b6d4","#22c55e","#ef4444","#6b7280"]);

    // Chart 2: por unidad
    const byUnidad = {};
    filtradas.filter(s => s.UnidadDerivada).forEach(s => {
      byUnidad[s.UnidadDerivada] = (byUnidad[s.UnidadDerivada] || 0) + 1;
    });
    renderChart("chart-unidad", "bar", Object.keys(byUnidad), Object.values(byUnidad),
      ["#1a3a6b","#2563a8","#3b82f6","#60a5fa","#93c5fd"]);

    // Chart 3: por mes
    const byMes = {};
    filtradas.forEach(s => {
      const m = s.FechaRecepcion?.substring(0, 7);
      if (m) byMes[m] = (byMes[m] || 0) + 1;
    });
    const meses = Object.keys(byMes).sort();
    renderChart("chart-mes", "line", meses.map(m => {
      const [y, mo] = m.split("-");
      return `${["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"][parseInt(mo)-1]} ${y}`;
    }), meses.map(m => byMes[m]), ["#1a3a6b"]);

  } catch (e) {
    showToast("error", "Error en reportes: " + e.message);
  } finally {
    hideLoading();
  }
}

function renderChart(id, type, labels, data, colors) {
  if (state.chartInstances[id]) state.chartInstances[id].destroy();
  const ctx = document.getElementById(id)?.getContext("2d");
  if (!ctx) return;
  state.chartInstances[id] = new Chart(ctx, {
    type,
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: type === "line" ? "rgba(26,58,107,0.1)" : colors,
        borderColor: type === "line" ? "#1a3a6b" : colors,
        borderWidth: type === "line" ? 2 : 1,
        fill: type === "line",
        tension: 0.4,
        pointBackgroundColor: "#1a3a6b"
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { position: type === "doughnut" ? "right" : "top", labels: { font: { size: 12 } } }
      },
      scales: type !== "doughnut" ? {
        y: { beginAtZero: true, ticks: { stepSize: 1 } }
      } : {}
    }
  });
}

async function exportarExcel() {
  const all = await getSolicitudes();
  const desde = document.getElementById("graf-desde")?.value;
  const hasta = document.getElementById("graf-hasta")?.value;
  const filtradas = all.filter(s => {
    const f = new Date(s.FechaRecepcion);
    return (!desde || f >= new Date(desde)) && (!hasta || f <= new Date(hasta));
  });
  const csv = ["NroSolicitud,FechaRecepcion,Solicitante,Direccion,Estado,UnidadDerivada,Solicitud"]
    .concat(filtradas.map(s =>
      `"${s.NroSolicitud}","${formatFecha(s.FechaRecepcion)}","${s.Solicitante}","${s.Direccion||""}","${s.Estado}","${s.UnidadDerivada||""}","${(s.Solicitud||"").replace(/"/g,'""')}"`
    )).join("\n");
  const blob = new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `solicitudes_${desde}_${hasta}.csv`;
  a.click();
}

// ===== HELPERS =====
function seleccionarSolicitud(id) {
  const sol = state.solicitudes.find(s => s.id === id);
  if (!sol) return;

  // Si ya estaba seleccionada, deseleccionar y volver al form nuevo
  if (state.solicitudSeleccionada?.id === id) {
    state.solicitudSeleccionada = null;
    renderListaSolicitudes();
    renderFormNueva();
    return;
  }

  state.solicitudSeleccionada = sol;
  renderListaSolicitudes();
  cargarSolicitudEnFormulario(sol);
}

function cargarSolicitudEnFormulario(sol) {
  // ── Header dinámico: modo edición vs solo lectura ──
  const editable = sol.Estado === CONFIG.estados.INGRESADA;
  const header = document.getElementById("form-panel-header");
  if (header) {
    header.style.cssText = `display:flex;align-items:center;gap:10px;padding:12px 16px;color:white;
      background:${editable ? 'linear-gradient(90deg,#166534,#15803d)' : 'linear-gradient(90deg,#0f2547,#1a3a6b)'};`;
    header.innerHTML = `
      <span style="font-size:16px;">${editable ? '✏️' : '👁️'}</span>
      <span style="font-weight:600;font-size:13px;opacity:0.9;">${editable ? 'Editando' : 'Viendo'}</span>
      <span style="font-weight:800;font-size:16px;letter-spacing:0.5px;color:white;">Solicitud #${sol.NroSolicitud}</span>
      <span style="background:rgba(255,255,255,0.2);color:white;font-size:11px;font-weight:700;
            padding:3px 10px;border-radius:12px;letter-spacing:0.3px;">${sol.Estado}</span>
      <button onclick="limpiarSeleccion()" title="Cerrar"
        style="margin-left:auto;background:rgba(255,255,255,0.15);border:1.5px solid rgba(255,255,255,0.4);
               color:white;padding:5px 16px;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600;">
        ✕ Cerrar
      </button>`;
  }

  // ── Banner de modo edición ──
  const bannerHtml = editable
    ? `<div style="background:#f0fdf4;border:1.5px solid #86efac;border-radius:10px;padding:10px 14px;
                   display:flex;align-items:center;gap:10px;font-size:13px;color:#15803d;">
         <span style="font-size:20px;">✏️</span>
         <div>
           <strong>Modo edición activo</strong><br>
           <span style="font-size:12px;">Puedes modificar los datos y guardar los cambios.</span>
         </div>
       </div>`
    : `<div style="background:#f8fafc;border:1.5px solid #e2e8f0;border-radius:10px;padding:10px 14px;
                   display:flex;align-items:center;gap:10px;font-size:13px;color:#64748b;">
         <span style="font-size:20px;">👁️</span>
         <div>
           <strong>Solicitud en estado "${sol.Estado}"</strong> — solo lectura<br>
           <span style="font-size:12px;">Los datos no pueden modificarse en este estado.</span>
         </div>
       </div>`;

  const cont = document.getElementById("panel-nueva");
  cont.innerHTML = `
    <div class="form-nueva">

      ${bannerHtml}

      <!-- ── Datos ── -->
      <div class="form-section">
        <div class="form-section-header">📋 Datos de la Solicitud</div>
        <div class="form-section-body">
          <div class="form-row">
            <div class="form-group">
              <label>Nro Solicitud</label>
              <input type="text" id="nueva-nro" value="${sol.NroSolicitud||''}" ${editable?'':'readonly'}>
            </div>
            <div class="form-group">
              <label>Fecha Recepción</label>
              <input type="date" id="nueva-fecha" value="${sol.FechaRecepcion?.split('T')[0]||''}" ${editable?'':'readonly'}>
            </div>
          </div>
          <div class="form-group">
            <label>Nombre Solicitante</label>
            <input type="text" id="nueva-solicitante" value="${sol.Solicitante||''}" ${editable?'':'readonly'}>
          </div>
          <div class="form-group">
            <label>Dirección</label>
            <input type="text" id="nueva-dir" value="${sol.Direccion||''}" ${editable?'':'readonly'}>
          </div>
          <div class="form-group">
            <label>Descripción de la solicitud</label>
            <textarea id="nueva-solicitud" rows="3" ${editable?'':'readonly'}>${sol.Solicitud||''}</textarea>
          </div>
          ${sol.UnidadDerivada?`
          <div class="form-group">
            <label>📤 Unidad Derivada</label>
            <input type="text" value="${sol.UnidadDerivada}" readonly style="color:#b45309;font-weight:700;">
          </div>`:''}
          ${sol.MotivoDevolucion?`
          <div class="form-group">
            <label>↩️ Motivo de Devolución</label>
            <textarea rows="2" readonly style="color:#b91c1c;background:#fff5f5;border-color:#fca5a5;">${sol.MotivoDevolucion}</textarea>
          </div>`:''}
        </div>
      </div>

      <!-- ── Adjuntos ── -->
      <div class="form-section">
        <div class="form-section-header naranja">📎 Documentos Adjuntos</div>
        <div class="form-section-body">
          <div id="adjuntos-existentes">
            <div style="text-align:center;color:#9ca3af;padding:16px;font-size:13px;">⏳ Cargando adjuntos...</div>
          </div>
          ${editable?`
          <div style="margin-top:12px;border-top:1px solid var(--borde);padding-top:12px;">
            <div style="font-size:12px;color:#b45309;font-weight:600;margin-bottom:8px;">➕ Agregar nuevos documentos</div>
            <div id="drop-area" class="upload-area" onclick="document.getElementById('nueva-files').click()"
              ondragover="event.preventDefault()" ondrop="event.preventDefault();handleFiles(event.dataTransfer.files)">
              <div style="font-size:28px;">📄</div>
              <div style="font-size:13px;font-weight:600;">Arrastra o haz clic para subir</div>
            </div>
            <input type="file" id="nueva-files" multiple accept=".pdf,.jpg,.jpeg,.png" onchange="handleFiles(this.files)">
            <div id="file-list" class="file-list"></div>
          </div>`:''}
        </div>
      </div>

      <!-- ── Historial ── -->
      <div class="form-section">
        <div class="form-section-header verde">🕐 Historial de Movimientos</div>
        <div class="form-section-body" id="historial-inline" style="max-height:240px;overflow-y:auto;padding:8px 16px;">
          <div style="text-align:center;color:#9ca3af;padding:16px;font-size:13px;">⏳ Cargando historial...</div>
        </div>
      </div>

      <!-- ── Botones ── -->
      ${editable?`
      <div style="display:flex;justify-content:center;gap:12px;padding:8px 0 12px;">
        <button class="btn-primary" onclick="guardarEdicionSolicitud('${sol.id}')"
          style="width:220px;padding:13px;font-size:15px;">
          💾 Guardar Cambios
        </button>
        <button onclick="limpiarSeleccion()"
          style="padding:12px 22px;border:1.5px solid var(--borde);border-radius:8px;background:white;cursor:pointer;font-size:13px;color:#666;">
          ✕ Cancelar
        </button>
      </div>` : `
      <div style="display:flex;justify-content:center;padding:12px 0;">
        <button onclick="limpiarSeleccion()"
          style="padding:11px 28px;border:1.5px solid var(--azul);border-radius:8px;background:white;
                 cursor:pointer;font-size:13px;color:var(--azul);font-weight:600;">
          ➕ Ingresar nueva solicitud
        </button>
      </div>`}
    </div>`;

  // ── Cargar adjuntos → panel derecho (visor) + lista en formulario ──
  const pdfPanel = document.getElementById("pdf-visor-contenido");
  const pdfHeader = document.getElementById("pdf-panel-header");
  if (pdfPanel) pdfPanel.innerHTML = `<div style="text-align:center;color:#9ca3af;padding:30px;font-size:13px;">⏳ Cargando documento...</div>`;

  getListItemAttachments(CONFIG.lists.solicitudes, sol.id).then(atts => {
    // Lista compacta en formulario
    const c = document.getElementById("adjuntos-existentes");
    if (c) {
      if (!atts.length) {
        c.innerHTML = `<p style="color:#9ca3af;font-size:13px;text-align:center;padding:8px;">Sin documentos adjuntos</p>`;
      } else {
        c.innerHTML = atts.map(a => {
          const isPdf = a.name?.toLowerCase().endsWith('.pdf');
          const isImg = /\.(jpg|jpeg|png)$/i.test(a.name||'');
          return `<div class="file-item" style="cursor:pointer;border-left:3px solid ${isPdf?'#ef4444':isImg?'#3b82f6':'#6b7280'};"
            onclick="mostrarEnVisor('${a.downloadUrl}','${a.name}',${isPdf},'${sol.NroSolicitud}','${a.serverRelativeUrl}')">
            <span>${isPdf?'📄':isImg?'🖼️':'📎'} <strong>${a.name}</strong></span>
            <span style="font-size:11px;color:var(--azul);">Ver ↗</span>
          </div>`;
        }).join('');
      }
    }

    // Visor derecho: mostrar primer PDF o imagen
    if (!pdfPanel) return;
    if (!atts.length) {
      pdfPanel.innerHTML = `<div class="pdf-visor-empty"><span>📭</span><p>Esta solicitud no tiene documentos adjuntos</p></div>`;
      if (pdfHeader) pdfHeader.textContent = "📄 Sin adjuntos";
      return;
    }
    const first = atts.find(a => a.name?.toLowerCase().endsWith('.pdf')) || atts[0];
    mostrarEnVisor(first.downloadUrl, first.name, first.name?.toLowerCase().endsWith('.pdf'), sol.NroSolicitud, first.serverRelativeUrl);

    // Si hay varios adjuntos, mostrar miniaturas abajo
    if (atts.length > 1) {
      const thumbBar = document.createElement("div");
      thumbBar.className = "thumb-bar";
      thumbBar.style.cssText = "display:flex;gap:6px;padding:8px 0 2px;flex-shrink:0;overflow-x:auto;border-top:1px solid var(--borde);margin-top:6px;";
      atts.forEach(a => {
        const isPdf = a.name?.toLowerCase().endsWith('.pdf');
        const isImg = /\.(jpg|jpeg|png)$/i.test(a.name||'');
        const thumb = document.createElement("div");
        thumb.style.cssText = `min-width:60px;height:60px;border-radius:8px;border:2px solid var(--borde);
          cursor:pointer;display:flex;flex-direction:column;align-items:center;justify-content:center;
          font-size:20px;background:#f8fafc;flex-shrink:0;transition:border-color 0.2s;`;
        thumb.innerHTML = isPdf ? '📄' : isImg
          ? `<img src="${a.downloadUrl}" style="width:100%;height:100%;object-fit:cover;border-radius:6px;">`
          : '📎';
        thumb.title = a.name;
        thumb.onclick = async () => {
          document.querySelectorAll(".thumb-active").forEach(t => { t.classList.remove("thumb-active"); t.style.borderColor = "var(--borde)"; });
          thumb.style.borderColor = "var(--azul)";
          thumb.classList.add("thumb-active");
          await mostrarEnVisor(a.downloadUrl, a.name, isPdf, sol.NroSolicitud, a.serverRelativeUrl);
        };
        thumbBar.appendChild(thumb);
      });
      pdfPanel.appendChild(thumbBar);
    }
  }).catch(() => {
    const c = document.getElementById("adjuntos-existentes");
    if (c) c.innerHTML = `<p style="color:#9ca3af;font-size:13px;text-align:center;">No se pudieron cargar los adjuntos</p>`;
    if (pdfPanel) pdfPanel.innerHTML = `<div class="pdf-visor-empty"><span>⚠️</span><p>Error al cargar el documento</p></div>`;
  });

  // ── Cargar historial ──
  getHistorialBySolicitud(sol.NroSolicitud).then(hist => {
    const hc = document.getElementById("historial-inline");
    if (!hc) return;
    hist.sort((a,b) => new Date(b.FechaAccion) - new Date(a.FechaAccion));
    if (!hist.length) {
      hc.innerHTML = `<p style="text-align:center;color:#9ca3af;font-size:13px;padding:12px;">Sin historial registrado</p>`;
      return;
    }
    const dotColor = a => {
      a = a?.toLowerCase()||'';
      if (a.includes('deriv'))   return '#f59e0b';
      if (a.includes('respond')) return '#22c55e';
      if (a.includes('devuel'))  return '#ef4444';
      if (a.includes('cerr'))    return '#6b7280';
      return '#3b82f6';
    };
    hc.innerHTML = hist.map((h,i) => `
      <div style="display:flex;gap:12px;padding:10px 0;${i<hist.length-1?'border-bottom:1px solid #f3f4f6':''}">
        <div style="display:flex;flex-direction:column;align-items:center;gap:2px;">
          <div style="width:12px;height:12px;border-radius:50%;background:${dotColor(h.Accion)};flex-shrink:0;margin-top:2px;"></div>
          ${i<hist.length-1?`<div style="width:2px;flex:1;background:#f0f0f0;margin-top:2px;"></div>`:''}
        </div>
        <div style="flex:1;padding-bottom:4px;">
          <div style="font-size:13px;font-weight:700;color:#1a1a1a;">${h.Accion}</div>
          <div style="font-size:11px;color:#888;margin-top:1px;">
            📅 ${formatFechaHora(h.FechaAccion)}
            ${h.UsuarioAccion?`· 👤 ${h.UsuarioAccion}`:''}
            ${h.Unidad?`· 🏢 ${h.Unidad}`:''}
          </div>
          ${h.EstadoAnterior?`<div style="font-size:11px;color:#aaa;margin-top:2px;">
            <span class="estado-badge estado-${h.EstadoAnterior}" style="font-size:10px;">${h.EstadoAnterior}</span>
            → <span class="estado-badge estado-${h.EstadoNuevo}" style="font-size:10px;">${h.EstadoNuevo}</span>
          </div>`:''}
          ${h.Observaciones?`<div style="font-size:12px;color:#4b5563;margin-top:4px;padding:6px 8px;background:#f9fafb;border-left:3px solid ${dotColor(h.Accion)};border-radius:0 4px 4px 0;">"${h.Observaciones}"</div>`:''}
        </div>
      </div>`).join('');
  }).catch(() => {
    const hc = document.getElementById("historial-inline");
    if (hc) hc.innerHTML = `<p style="color:#9ca3af;font-size:13px;text-align:center;">No se pudo cargar el historial</p>`;
  });
}

async function mostrarEnVisor(downloadUrl, nombre, isPdf, nroSolicitud, serverRelativeUrl) {
  const panel = document.getElementById("pdf-visor-contenido");
  const header = document.getElementById("pdf-panel-header");
  if (!panel) return;

  // Header del visor
  if (header) {
    header.style.display = "flex";
    header.innerHTML = `
      📄 <span style="font-weight:700;">${nroSolicitud}</span>
      <span style="font-weight:400;font-size:12px;opacity:0.8;margin-left:4px;">— ${nombre}</span>
      <span id="visor-loader" style="margin-left:auto;font-size:12px;opacity:0.7;">⏳ Cargando...</span>`;
  }

  // Mantener barra de miniaturas si existe
  const thumbBar = panel.querySelector(".thumb-bar");

  // Mostrar spinner mientras carga
  const spinnerEl = document.createElement("div");
  spinnerEl.style.cssText = "flex:1;display:flex;align-items:center;justify-content:center;flex-direction:column;gap:12px;color:#9ca3af;";
  spinnerEl.innerHTML = `<div class="spinner"></div><p style="font-size:13px;">Cargando documento...</p>`;
  panel.innerHTML = "";
  panel.appendChild(spinnerEl);
  if (thumbBar) panel.appendChild(thumbBar);

  try {
    // Si es blob URL local (nueva solicitud) la usamos directo
    // Si es URL de SharePoint, descargamos vía REST API con autenticación
    let blobUrl = downloadUrl;
    if (!downloadUrl.startsWith("blob:")) {
      blobUrl = await getAttachmentBlobUrl(downloadUrl, serverRelativeUrl);
    }

    // Actualizar header con botones
    if (header) {
      header.innerHTML = `
        📄 <span style="font-weight:700;">${nroSolicitud}</span>
        <span style="font-weight:400;font-size:12px;opacity:0.8;margin-left:4px;">— ${nombre}</span>
        <div style="margin-left:auto;display:flex;gap:6px;">
          <a href="${blobUrl}" download="${nombre}"
            style="background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.3);
                   color:white;padding:4px 10px;border-radius:6px;font-size:12px;text-decoration:none;">
            ⬇ Bajar
          </a>
          <a href="${downloadUrl}" target="_blank"
            style="background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.3);
                   color:white;padding:4px 10px;border-radius:6px;font-size:12px;text-decoration:none;">
            ↗ Abrir
          </a>
        </div>`;

    panel.innerHTML = "";
    if (isPdf) {
      // Renderizar con PDF.js dentro del sistema
      const wrap = document.createElement("div");
      wrap.style.cssText = "flex:1;overflow-y:auto;background:#525659;border-radius:8px;padding:12px;display:flex;flex-direction:column;align-items:center;gap:8px;min-height:0;";
      wrap.innerHTML = `<div style="color:white;font-size:13px;opacity:0.7;">⏳ Renderizando PDF...</div>`;
      panel.appendChild(wrap);
      if (thumbBar) panel.appendChild(thumbBar);

      // Configurar PDF.js worker
      if (typeof pdfjsLib !== "undefined") {
        pdfjsLib.GlobalWorkerOptions.workerSrc =
          "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
        try {
          const loadingTask = pdfjsLib.getDocument(blobUrl);
          const pdf = await loadingTask.promise;
          wrap.innerHTML = ""; // limpiar spinner

          for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
            const page = await pdf.getPage(pageNum);
            const containerWidth = wrap.clientWidth - 24;
            const viewport = page.getViewport({ scale: 1 });
            const scale = containerWidth / viewport.width;
            const scaledViewport = page.getViewport({ scale });

            const canvas = document.createElement("canvas");
            canvas.width = scaledViewport.width;
            canvas.height = scaledViewport.height;
            canvas.style.cssText = "width:100%;border-radius:4px;box-shadow:0 2px 8px rgba(0,0,0,0.3);";

            const ctx = canvas.getContext("2d");
            await page.render({ canvasContext: ctx, viewport: scaledViewport }).promise;
            wrap.appendChild(canvas);

            // Separador entre páginas
            if (pageNum < pdf.numPages) {
              const sep = document.createElement("div");
              sep.style.cssText = "width:100%;text-align:center;color:rgba(255,255,255,0.4);font-size:11px;padding:2px 0;";
              sep.textContent = `— Página ${pageNum} de ${pdf.numPages} —`;
              wrap.appendChild(sep);
            }
          }
        } catch (pdfErr) {
          wrap.innerHTML = `<div style="color:#fca5a5;text-align:center;padding:20px;">
            <p>Error al renderizar PDF</p><small>${pdfErr.message}</small></div>`;
        }
      } else {
        wrap.innerHTML = `<div style="color:#fca5a5;padding:20px;text-align:center;">PDF.js no disponible</div>`;
      }
    } else if (/\.(jpg|jpeg|png|gif)$/i.test(nombre)) {
      const wrap = document.createElement("div");
      wrap.style.cssText = "flex:1;overflow:auto;display:flex;align-items:center;justify-content:center;background:#525659;border-radius:8px;min-height:0;";
      const img = document.createElement("img");
      img.src = blobUrl;
      img.style.cssText = "max-width:100%;max-height:100%;object-fit:contain;border-radius:4px;cursor:zoom-in;";
      img.title = "Clic para abrir";
      img.onclick = () => window.open(blobUrl, "_blank");
      wrap.appendChild(img);
      panel.appendChild(wrap);
    } else {
      panel.innerHTML = `
        <div class="pdf-visor-empty">
          <span>📎</span><p>${nombre}</p>
          <a href="${blobUrl}" download="${nombre}"
            style="color:var(--azul);font-size:13px;font-weight:600;padding:8px 20px;border:1.5px solid var(--azul);border-radius:8px;">
            ⬇ Descargar archivo
          </a>
        </div>`;
    }
    if (thumbBar) panel.appendChild(thumbBar);

  } catch (e) {
    panel.innerHTML = `
      <div class="pdf-visor-empty" style="gap:16px;">
        <span>📄</span>
        <p style="color:#374151;font-weight:600;">${nombre}</p>
        <p style="color:#9ca3af;font-size:12px;">No se puede mostrar el documento aquí,<br>pero puedes abrirlo en SharePoint:</p>
        <a href="${downloadUrl}" target="_blank"
          style="background:var(--azul);color:white;padding:10px 24px;border-radius:8px;
                 font-size:14px;font-weight:600;text-decoration:none;">
          ↗ Abrir documento en SharePoint
        </a>
        <small style="color:#d1d5db;font-size:11px;">${e.message}</small>
      </div>`;
  }
}

function limpiarSeleccion() {
  state.solicitudSeleccionada = null;
  state.adjuntosNueva = [];
  renderListaSolicitudes();
  renderFormNueva();
  const header = document.getElementById("form-panel-header");
  if (header) { header.textContent = "➕ Nueva Solicitud"; header.style.cssText = ""; }
  // Limpiar visor
  const pdfPanel = document.getElementById("pdf-visor-contenido");
  if (pdfPanel) pdfPanel.innerHTML = `<div class="pdf-visor-empty"><span>📄</span><p>Selecciona una solicitud para ver el documento adjunto</p></div>`;
  const pdfHeader = document.getElementById("pdf-panel-header");
  if (pdfHeader) { pdfHeader.textContent = "📄 Documento"; pdfHeader.style.display = ""; }
}

async function guardarEdicionSolicitud(solId) {
  const nro  = document.getElementById("nueva-nro")?.value.trim();
  const fecha = document.getElementById("nueva-fecha")?.value;
  const sol2  = document.getElementById("nueva-solicitante")?.value.trim();
  const dir   = document.getElementById("nueva-dir")?.value.trim();
  const desc  = document.getElementById("nueva-solicitud")?.value.trim();
  if (!nro||!fecha||!sol2||!dir) { showToast("error","Completa los campos obligatorios"); return; }
  showLoading("Guardando cambios...");
  try {
    await actualizarSolicitud(solId, { NroSolicitud:nro, FechaRecepcion:fecha, Solicitante:sol2, Direccion:dir, Solicitud:desc });
    if (state.adjuntosNueva.length > 0) {
      for (const file of state.adjuntosNueva) {
        await uploadAttachment(CONFIG.lists.solicitudes, solId, file).catch(console.error);
      }
      state.adjuntosNueva = [];
    }
    showToast("success", `✅ Solicitud ${nro} actualizada`);
    await renderSecretaria();
  } catch(e) {
    showToast("error","Error: "+e.message);
  } finally { hideLoading(); }
}

function formatFecha(iso) {
  if (!iso) return "-";
  return new Date(iso).toLocaleDateString("es-CL", { day: "2-digit", month: "long", year: "numeric" });
}
function formatFechaHora(iso) {
  if (!iso) return "-";
  return new Date(iso).toLocaleString("es-CL", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });
}

function updateTabBadges() {
  const ingresadas = state.solicitudes.filter(s => s.Estado === "Ingresada").length;
  document.querySelectorAll(".tab-btn").forEach(b => {
    if (b.dataset.tab === "solicitudes" && ingresadas > 0) {
      const badge = b.querySelector(".badge") || Object.assign(document.createElement("span"), {className:"badge warn"});
      badge.textContent = ingresadas;
      b.appendChild(badge);
    }
  });
}

function showLoading(msg = "Cargando...") {
  let el = document.getElementById("loading-overlay");
  if (!el) {
    el = document.createElement("div");
    el.id = "loading-overlay";
    el.className = "loading-overlay";
    document.body.appendChild(el);
  }
  el.innerHTML = `<div class="loading-spinner"><div class="spinner"></div><p style="color:#666;font-size:14px;">${msg}</p></div>`;
  el.style.display = "flex";
}

function hideLoading() {
  const el = document.getElementById("loading-overlay");
  if (el) el.style.display = "none";
}

function showToast(type, msg) {
  let cont = document.getElementById("toast-container");
  if (!cont) {
    cont = document.createElement("div");
    cont.id = "toast-container";
    cont.className = "toast-container";
    document.body.appendChild(cont);
  }
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.textContent = msg;
  cont.appendChild(toast);
  setTimeout(() => toast.remove(), 4000);
}

// ===== BOOTSTRAP =====
window.addEventListener("load", initApp);
