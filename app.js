// ===== PDF VISOR STATE (unidad) =====
let _uniPdfDoc  = null;
let _uniPdfZoom = 1.0;

// ===== STATE =====
const state = {
  usuario: null,
  solicitudes: [],
  solicitudSeleccionada: null,
  filtroEstado: "Todos",
  filtroBuscar: "",
  filtroUnidad: "Todas",
  filtroUnidadDir: "Todas",
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
      showLoginPage(msUser, `El usuario ${msUser.mail || msUser.userPrincipalName} no tiene acceso al sistema. Contacta al administrador TI.`);
      return;
    }
    state.usuario = { ...perfil, displayName: msUser.displayName };
    // Cargar unidades desde SharePoint (reemplaza la lista hardcodeada)
    const unidadesSP = await getUnidades().catch(() => []);
    if (unidadesSP.length) {
      CONFIG.unidades = unidadesSP.map(u => u.Title).filter(Boolean);
    }
    // Cargar configuración desde ConfiguracionDom
    const cfgSP = await getConfiguracionDOM().catch(() => ({}));
    CONFIG.plazoDerivacionDias = parseInt(cfgSP.PlazoDerivacionDias) || 15;
    CONFIG.plazoAlertaDias     = parseInt(cfgSP.PlazoAlertaDias)     || 1;
    CONFIG.correoSoporte       = cfgSP.CorreoSoporte                 || "enovo@mdonihue.cl";
    // Crear campo FechaDerivacion si aún no existe (idempotente, falla en silencio)
    crearCampoFechaDerivacion().catch(console.warn);
    crearCamposEvidencia().catch(console.warn);
    crearCamposHistorial().catch(console.warn);
    // Indexar columnas críticas (idempotente — no hace nada si ya están indexadas)
    crearIndicesSharePoint().catch(console.warn);
    hideLoading();
    showApp();
  } catch (e) {
    hideLoading();
    console.error(e);
    showLoginPage();
  }
}

function showLoginPage(msUser, errorMsg) {
  document.getElementById("login-page").style.display = "flex";
  document.getElementById("app").style.display = "none";

  // Mostrar error si viene
  const errEl = document.getElementById("login-error");
  const errMsg = document.getElementById("login-error-msg");
  if (errorMsg && errEl && errMsg) {
    errMsg.textContent = errorMsg;
    errEl.style.display = "flex";
  } else if (errEl) {
    errEl.style.display = "none";
  }

  // Mostrar datos del usuario si ya autenticó
  const userInfo = document.getElementById("login-user-info");
  if (msUser && userInfo) {
    document.getElementById("login-nombre").textContent = msUser.displayName || "";
    document.getElementById("login-correo").textContent = msUser.mail || msUser.userPrincipalName || "";
    userInfo.style.display = "block";
  }

  document.getElementById("btn-login").onclick = () => {
    showLoading("Redirigiendo a Microsoft...");
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
  if (rol === CONFIG.roles.ADMIN)      navigateTab("admin");
  else if (rol === CONFIG.roles.SECRETARIA) navigateTab("solicitudes");
  else if (rol === CONFIG.roles.DIRECTOR)   navigateTab("gestion");
  else navigateTab("unidad");
}

// ===== TABS =====
function buildTabs() {
  const bar = document.getElementById("tabs-bar");
  bar.innerHTML = "";
  const rol = state.usuario.Rol;

  const tabs = [];
  if (rol === CONFIG.roles.ADMIN) {
    tabs.push({ id: "admin",     icon: "🛡️", label: "Administración" });
    tabs.push({ id: "gestion",   icon: "⚙️", label: "Gestión" });
    tabs.push({ id: "graficos",  icon: "📊", label: "Reportes" });
  }
  if (rol === CONFIG.roles.SECRETARIA) {
    tabs.push({ id: "solicitudes", icon: "📋", label: "Ingreso Solicitudes" });
  }
  if (rol === CONFIG.roles.DIRECTOR) {
    tabs.push({ id: "gestion", icon: "⚙️", label: "Gestión" });
    tabs.push({ id: "graficos", icon: "📊", label: "Reportes" });
  }
  if (rol === CONFIG.roles.UNIDAD) {
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
  if (tabId === "gestion")    renderDirector();
  if (tabId === "unidad")     renderUnidad();
  if (tabId === "graficos")   renderGraficos();
  if (tabId === "admin")      renderAdmin();
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
  const filtradas = getSolicitudesFiltradas();
  const counts = {};
  filtradas.forEach(s => { counts[s.Estado] = (counts[s.Estado]||0)+1; });
  const total = filtradas.length;
  cont.innerHTML = `
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:6px;">
      ${[
        {e:"Ingresada",       lbl:"Ingresada",    color:"#dbeafe",tc:"#1d4ed8",icon:"📥"},
        {e:"Derivada",        lbl:"Derivada",      color:"#fef3c7",tc:"#b45309",icon:"📤"},
        {e:"En Proceso",      lbl:"En Proceso",    color:"#cffafe",tc:"#0e7490",icon:"⚙️"},
        {e:"Respondida",      lbl:"Respondida",    color:"#dcfce7",tc:"#15803d",icon:"✅"},
        {e:"Devuelta",        lbl:"Devuelta",      color:"#fee2e2",tc:"#b91c1c",icon:"↩️"},
        {e:"Pendiente de Cierre", lbl:"Pend. Cierre", color:"#fef9c3",tc:"#854d0e",icon:"⏳"},
        {e:"Cerrada",         lbl:"Cerrada",       color:"#f3f4f6",tc:"#4b5563",icon:"🔒"}
      ].map(({e,lbl,color,tc,icon})=>`
        <div onclick="state.filtroEstado='${e}';state.pagina=1;renderListaSolicitudes();renderStatsCompact()"
          style="background:${color};border-radius:8px;padding:8px;text-align:center;cursor:pointer;">
          <div style="font-size:11px;">${icon}</div>
          <div style="font-size:18px;font-weight:700;color:${tc};line-height:1.2;">${counts[e]||0}</div>
          <div style="font-size:10px;color:${tc};">${lbl}</div>
        </div>`).join("")}
    </div>
    <div style="text-align:center;font-size:12px;color:#9ca3af;margin-top:6px;">Total: ${total}</div>`;
}

function getSolicitudesFiltradas() {
  return state.solicitudes.filter(s => {
    if (state.filtroEstado !== "Todos" && s.Estado !== state.filtroEstado) return false;
    if (state.filtroUnidad !== "Todas" && (s.UnidadDerivada||"").trim() !== state.filtroUnidad) return false;
    if (state.filtroBuscar) {
      const q = state.filtroBuscar.toLowerCase();
      if (!s.NroSolicitud?.toLowerCase().includes(q) &&
          !s.Solicitante?.toLowerCase().includes(q) &&
          !s.Direccion?.toLowerCase().includes(q) &&
          !s.Solicitud?.toLowerCase().includes(q)) return false;
    }
    if (state.filtroDesde && new Date(s.FechaRecepcion) < new Date(state.filtroDesde)) return false;
    if (state.filtroHasta && new Date(s.FechaRecepcion) > new Date(state.filtroHasta + "T23:59:59")) return false;
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
    registrarHistorial({
      NroSolicitud: nro,
      Title:"Ingreso de solicitud",
      EstadoAnterior: "",
      EstadoNuevo: CONFIG.estados.INGRESADA,
      UsuarioAccion: state.usuario.NombreCompleto,
      RolUsuario: state.usuario.Rol,
      Unidad: state.usuario.Unidad,
      FechaAccion: new Date().toISOString(),
      Observaciones: desc
    }).catch(e => console.warn("Historial (no crítico):", e.message));

    // Notificar director (fire-and-forget)
    notificarDirector(item, "Nueva solicitud ingresada");

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
    // Orden: Devuelta (urgente) > Ingresada > Respondida > Derivada > En Proceso > Cerrada
    const prioridad = { "Devuelta":0,"Ingresada":1,"Respondida":2,"Derivada":3,"En Proceso":4,"Pendiente de Cierre":5,"Cerrada":6 };
    state.solicitudes.sort((a,b) => {
      const pa = prioridad[a.Estado]??9, pb = prioridad[b.Estado]??9;
      if (pa !== pb) return pa - pb;
      return new Date(b.FechaRecepcion) - new Date(a.FechaRecepcion);
    });
    renderDirSidebar();
    renderDirLista();
    updateTabBadges();
  } catch (e) {
    showToast("error", "Error: " + e.message);
  } finally {
    hideLoading();
  }
}

function renderDirSidebar() {
  const counts = {};
  state.solicitudes.forEach(s => { counts[s.Estado] = (counts[s.Estado]||0)+1; });

  const stats = [
    { e:"Ingresada",          icon:"📥", bg:"#dbeafe", tc:"#1d4ed8" },
    { e:"Derivada",           icon:"📤", bg:"#fef3c7", tc:"#b45309" },
    { e:"En Proceso",         icon:"⚙️", bg:"#cffafe", tc:"#0e7490" },
    { e:"Respondida",         icon:"✅", bg:"#dcfce7", tc:"#15803d" },
    { e:"Devuelta",           icon:"↩️", bg:"#fee2e2", tc:"#b91c1c" },
    { e:"Pendiente de Cierre",icon:"⏳", bg:"#fdf4ff", tc:"#7e22ce" },
    { e:"Cerrada",            icon:"🔒", bg:"#f3f4f6", tc:"#4b5563" }
  ];

  const sidebar = document.getElementById("dir-sidebar");
  sidebar.innerHTML = `
    <!-- Botón Todos -->
    <button onclick="filtrarDirector('Todos')"
      style="width:100%;margin-bottom:6px;padding:6px 10px;border-radius:7px;border:1.5px solid ${state.filtroEstado==='Todos'?'var(--azul)':'var(--borde)'};
             background:${state.filtroEstado==='Todos'?'var(--azul)':'white'};color:${state.filtroEstado==='Todos'?'white':'var(--texto)'};
             font-size:11px;font-weight:600;cursor:pointer;display:flex;justify-content:space-between;align-items:center;">
      <span>📊 Todas</span>
      <span style="background:${state.filtroEstado==='Todos'?'rgba(255,255,255,0.25)':'var(--gris-bg)'};padding:1px 7px;border-radius:10px;font-weight:700;">
        ${state.solicitudes.length}
      </span>
    </button>

    <!-- Grid 2x3 de estados compacto -->
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;">
      ${stats.map(({e,icon,bg,tc}) => {
        const cnt = counts[e]||0;
        const activo = state.filtroEstado === e;
        const urgente = e==="Devuelta" && cnt>0;
        return `
        <div onclick="filtrarDirector('${e}')"
          style="background:${activo?tc:bg};border-radius:8px;padding:6px 4px;text-align:center;cursor:pointer;
                 border:2px solid ${activo?tc:urgente?'#fca5a5':'transparent'};
                 transition:all 0.18s;box-shadow:${activo?'0 2px 8px rgba(0,0,0,0.15)':'none'};"
          onmouseenter="this.style.transform='translateY(-1px)'"
          onmouseleave="this.style.transform=''">
          <div style="font-size:13px;margin-bottom:1px;">${icon}</div>
          <div style="font-size:17px;font-weight:800;color:${activo?'white':tc};line-height:1;">${cnt}</div>
          <div style="font-size:9px;color:${activo?'rgba(255,255,255,0.85)':tc};margin-top:1px;font-weight:500;">${e}</div>
          ${urgente&&!activo?`<div style="font-size:8px;color:#b91c1c;font-weight:700;">⚠️</div>`:''}
        </div>`;
      }).join("")}
    </div>

    <!-- Filtro por unidad (valores reales de las solicitudes cargadas) -->
    <div style="margin-top:10px;padding-top:8px;border-top:1px solid var(--borde);">
      <div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:0.8px;color:#9ca3af;margin-bottom:5px;">🏢 Filtrar por Unidad</div>
      <select onchange="filtrarDirectorUnidad(this.value)"
        style="width:100%;padding:5px 8px;border:1.5px solid ${state.filtroUnidadDir!=='Todas'?'var(--azul)':'var(--borde)'};
               border-radius:7px;font-size:11px;color:var(--texto);cursor:pointer;
               background:${state.filtroUnidadDir!=='Todas'?'var(--azul-50)':'white'};">
        <option value="Todas">Todas las unidades</option>
        ${[...new Set(state.solicitudes.map(s=>(s.UnidadDerivada||"").trim()).filter(u=>u))].sort()
          .map(u=>`<option value="${u}" ${state.filtroUnidadDir===u?'selected':''}>${u}</option>`).join("")}
      </select>
      ${state.filtroUnidadDir!=='Todas'?`
      <div style="margin-top:4px;font-size:10px;color:var(--azul-claro);display:flex;align-items:center;justify-content:space-between;">
        <span>📌 ${state.filtroUnidadDir}</span>
        <button onclick="filtrarDirectorUnidad('Todas')"
          style="background:none;border:none;color:#ef4444;cursor:pointer;font-size:10px;font-weight:700;padding:0;">✕ limpiar</button>
      </div>`:''}
    </div>`;
}

function filtrarDirector(estado) {
  state.filtroEstado = estado;
  state.pagina = 1;
  state.solicitudSeleccionada = null;
  state.filtroBuscar = "";
  renderDirSidebar();
  renderDirLista();
  // Limpiar visor y detalle
  const pdf = document.getElementById("dir-pdf");
  if (pdf) pdf.innerHTML = `<div class="pdf-visor-empty"><span>📄</span><p>Selecciona una solicitud</p></div>`;
  const det = document.getElementById("dir-detalle");
  if (det) det.innerHTML = `<div class="pdf-visor-empty" style="height:100%;"><span>🏛️</span><p style="text-align:center;">Selecciona una solicitud<br>para gestionar</p></div>`;
}

function filtrarDirectorUnidad(unidad) {
  state.filtroUnidadDir = unidad;
  state.pagina = 1;
  state.solicitudSeleccionada = null;
  renderDirSidebar();
  renderDirLista();
  const pdf = document.getElementById("dir-pdf");
  if (pdf) pdf.innerHTML = `<div class="pdf-visor-empty"><span>📄</span><p>Selecciona una solicitud</p></div>`;
  const det = document.getElementById("dir-detalle");
  if (det) det.innerHTML = `<div class="pdf-visor-empty" style="height:100%;"><span>🏛️</span><p style="text-align:center;">Selecciona una solicitud<br>para gestionar</p></div>`;
}

async function cerrarDirectoDirector(id) {
  const sol = state.solicitudes.find(s => s.id === id);
  if (!sol) return;
  const obs = document.getElementById("dir-no-corresponde-obs")?.value.trim();
  if (!obs) { showToast("error", "Ingresa el motivo del cierre."); return; }

  const confirmar = confirm(`¿Cerrar la solicitud ${sol.NroSolicitud} como "No corresponde a DOM"?\n\nMotivo: ${obs}`);
  if (!confirmar) return;

  showLoading("Cerrando solicitud...");
  try {
    await actualizarSolicitud(sol.id, {
      Estado: CONFIG.estados.CERRADA,
      MotivoDevolucion: `No corresponde a DOM: ${obs}`
    });
    await registrarHistorial({
      NroSolicitud: sol.NroSolicitud,
      Accion: "Cierre directo por Director — No corresponde a DOM",
      Observaciones: obs,
      Usuario: state.usuario.NombreCompleto || state.usuario.displayName,
      Unidad: "Director",
      FechaAccion: new Date().toISOString()
    });
    showToast("ok", `Solicitud ${sol.NroSolicitud} cerrada.`);
    await renderDirector();
  } catch(e) {
    showToast("error", "Error: " + e.message);
  } finally {
    hideLoading();
  }
}

function getDirFiltradas() {
  return state.solicitudes.filter(s => {
    if (state.filtroEstado !== "Todos" && s.Estado !== state.filtroEstado) return false;
    if (state.filtroUnidadDir !== "Todas" && (s.UnidadDerivada||"").trim() !== state.filtroUnidadDir) return false;
    if (state.filtroBuscar) {
      const q = state.filtroBuscar.toLowerCase();
      return s.NroSolicitud?.toLowerCase().includes(q) ||
             s.Solicitante?.toLowerCase().includes(q) ||
             s.Direccion?.toLowerCase().includes(q);
    }
    return true;
  });
}

function renderDirLista() {
  const filtradas = getDirFiltradas();
  const total = filtradas.length;
  const inicio = (state.pagina-1)*state.pageSize;
  const pagina = filtradas.slice(inicio, inicio+state.pageSize);
  const totalPags = Math.ceil(total/state.pageSize);

  const count = document.getElementById("dir-lista-count");
  if (count) count.textContent = `(${total})`;

  // Actualizar input de búsqueda sin re-renderizar
  const buscarInput = document.getElementById("dir-buscar");
  if (buscarInput) buscarInput.oninput = e => {
    state.filtroBuscar = e.target.value;
    state.pagina = 1;
    renderDirLista();
  };

  const cont = document.getElementById("dir-lista-cards");
  if (!cont) return;

  if (pagina.length === 0) {
    cont.innerHTML = `<div style="text-align:center;color:#9ca3af;padding:40px;font-size:13px;">Sin solicitudes en este estado</div>`;
  } else {
    cont.innerHTML = pagina.map(s => {
      const esUrgente = s.Estado === CONFIG.estados.DEVUELTA;
      const esAccion  = s.Estado === CONFIG.estados.INGRESADA || s.Estado === CONFIG.estados.DEVUELTA || s.Estado === CONFIG.estados.RESPONDIDA;
      const sem       = calcularSemaforo(s);
      const semBadge  = sem
        ? `<span style="background:${sem.bg};color:${sem.color};border:1px solid ${sem.color}40;padding:1px 7px;border-radius:10px;font-size:10px;font-weight:700;white-space:nowrap;">${sem.emoji} ${sem.texto}</span>`
        : "";
      return `
        <div class="sol-card ${state.solicitudSeleccionada?.id===s.id?'selected':''}"
          style="border-left-color:${sem&&sem.dias<=1?'#dc2626':esUrgente?'#ef4444':esAccion?'var(--azul)':'var(--borde)'};"
          onclick="seleccionarSolicitudDirector('${s.id}')">
          <div class="sol-card-top">
            <span class="sol-nro">${s.NroSolicitud}</span>
            <span class="estado-badge estado-${s.Estado}">${s.Estado}</span>
          </div>
          <div class="sol-card-name">${s.Solicitante}</div>
          <div style="display:flex;align-items:center;justify-content:space-between;gap:4px;margin-top:2px;">
            <div class="sol-card-dir" style="margin-top:0;">📍 ${s.Direccion||""}</div>
            ${semBadge}
          </div>
          ${s.UnidadDerivada?`<div style="font-size:11px;color:#888;margin-top:3px;">🏢 ${s.UnidadDerivada}</div>`:""}
          ${esUrgente?`<div style="font-size:11px;color:#b91c1c;margin-top:3px;font-weight:600;">⚠️ Requiere re-derivación</div>`:""}
        </div>`;
    }).join("");
  }

  const pag = document.getElementById("dir-paginacion");
  if (pag) pag.innerHTML = `
    <button onclick="cambiarPaginaDir(${state.pagina-1})" ${state.pagina<=1?'disabled':''}>‹</button>
    <span>${state.pagina}/${Math.max(1,totalPags)} (${total})</span>
    <button onclick="cambiarPaginaDir(${state.pagina+1})" ${state.pagina>=totalPags?'disabled':''}>›</button>`;
}

function cambiarPaginaDir(p) { state.pagina=p; renderDirLista(); }

async function seleccionarSolicitudDirector(id) {
  const sol = state.solicitudes.find(s => s.id === id);
  if (!sol) return;
  state.solicitudSeleccionada = sol;
  renderDirLista();
  renderDetalleDirector(sol);
  mostrarDetalleMovil('.dir-layout');
  // Cargar PDF con PDF.js en panel central
  const pdfPanel = document.getElementById("dir-pdf");
  const pdfHeader = document.getElementById("dir-pdf-header");
  if (pdfPanel) {
    pdfPanel.innerHTML = `<div style="text-align:center;color:#9ca3af;padding:30px;"><div class="spinner" style="margin:0 auto 12px;"></div><p>Cargando documento...</p></div>`;
  }
  try {
    const atts = await getListItemAttachments(CONFIG.lists.solicitudes, sol.id);
    if (!atts.length) {
      if (pdfPanel) pdfPanel.innerHTML = `<div class="pdf-visor-empty"><span>📭</span><p>Sin documentos adjuntos</p></div>`;
      return;
    }
    const first = atts.find(a=>a.name?.toLowerCase().endsWith('.pdf'))||atts[0];
    // Reusar mostrarEnVisor pero apuntando al panel del director
    const tempPanel = document.getElementById("pdf-visor-contenido");
    // Temporalmente redirigir al panel del director
    if (pdfHeader) pdfHeader.textContent = `📄 ${sol.NroSolicitud} — ${first.name}`;
    const blobUrl = await getAttachmentBlobUrl(first.downloadUrl, first.serverRelativeUrl);
    if (pdfPanel) {
      pdfPanel.innerHTML = "";
      if (first.name?.toLowerCase().endsWith('.pdf') && typeof pdfjsLib !== "undefined") {
        pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
        const wrap = document.createElement("div");
        wrap.style.cssText = "flex:1;overflow-y:auto;background:#525659;border-radius:8px;padding:10px;display:flex;flex-direction:column;align-items:center;gap:6px;min-height:0;";
        pdfPanel.appendChild(wrap);
        const pdf = await pdfjsLib.getDocument(blobUrl).promise;
        for (let p=1; p<=pdf.numPages; p++) {
          const page = await pdf.getPage(p);
          const w = (wrap.clientWidth||350)-20;
          const vp = page.getViewport({scale:1});
          const scale = w/vp.width;
          const svp = page.getViewport({scale});
          const canvas = document.createElement("canvas");
          canvas.width=svp.width; canvas.height=svp.height;
          canvas.style.cssText="width:100%;border-radius:4px;box-shadow:0 2px 8px rgba(0,0,0,0.3);";
          await page.render({canvasContext:canvas.getContext("2d"),viewport:svp}).promise;
          wrap.appendChild(canvas);
          if (p<pdf.numPages) {
            const sep=document.createElement("div");
            sep.style.cssText="color:rgba(255,255,255,0.35);font-size:11px;";
            sep.textContent=`— ${p} / ${pdf.numPages} —`;
            wrap.appendChild(sep);
          }
        }
      } else {
        const img = document.createElement("img");
        img.src = blobUrl;
        img.style.cssText="width:100%;border-radius:8px;";
        pdfPanel.appendChild(img);
      }
    }
  } catch(e) {
    if (pdfPanel) pdfPanel.innerHTML=`<div class="pdf-visor-empty"><span>⚠️</span><p style="color:#ef4444;font-size:13px;">No se pudo cargar el documento</p></div>`;
  }
}

function renderDetalleDirector(sol) {
  const cont = document.getElementById("dir-detalle");
  const header = document.getElementById("dir-detalle-header");
  if (!sol) {
    if (cont) cont.innerHTML = `<div class="pdf-visor-empty" style="height:100%;"><span>🏛️</span><p>Selecciona una solicitud</p></div>`;
    return;
  }

  const esDerivable       = sol.Estado === CONFIG.estados.INGRESADA || sol.Estado === CONFIG.estados.DEVUELTA;
  const esCerrable        = sol.Estado === CONFIG.estados.RESPONDIDA || sol.Estado === CONFIG.estados.PENDIENTE_CIERRE;
  const esPendienteCierre = sol.Estado === CONFIG.estados.PENDIENTE_CIERRE;
  const tieneEvidencia    = sol.Estado === CONFIG.estados.RESPONDIDA || sol.Estado === CONFIG.estados.CERRADA || sol.Estado === CONFIG.estados.PENDIENTE_CIERRE;
  const esDevuelta        = sol.Estado === CONFIG.estados.DEVUELTA;

  if (header) {
    header.style.cssText = `background:${esDerivable?'linear-gradient(90deg,#1e3a5f,#1a3a6b)':esPendienteCierre?'linear-gradient(90deg,#4a1772,#7e22ce)':esCerrable?'linear-gradient(90deg,#14532d,#15803d)':'linear-gradient(90deg,#374151,#6b7280)'};color:white;display:flex;flex-direction:column;padding:0;`;
    header.innerHTML = `
      <button class="mobile-back-bar" onclick="volverAListaMovil('.dir-layout')" style="font-size:13px;padding:8px 12px;">
        ← Volver a lista
      </button>
      <div style="display:flex;align-items:center;gap:8px;flex:1;padding:10px 14px;">
        <span>${esDerivable?'📤':esPendienteCierre?'⏳':esCerrable?'✅':'📋'}</span>
        <span style="font-weight:700;">${sol.NroSolicitud}</span>
        <span class="estado-badge estado-${sol.Estado}" style="font-size:11px;">${sol.Estado}</span>
      </div>`;
  }

  const sem = calcularSemaforo(sol);
  const semBanner = sem ? `
    <div style="background:${sem.bg};border-bottom:2px solid ${sem.color};padding:8px 14px;display:flex;align-items:center;gap:10px;flex-shrink:0;">
      <span style="font-size:20px;">${sem.emoji}</span>
      <div style="flex:1;">
        <div style="font-size:12px;font-weight:700;color:${sem.color};">
          ${sem.dias > 0
            ? `Plazo: ${sem.dias === 1 ? "Vence mañana" : `${sem.dias} días restantes`}`
            : sem.dias === 0 ? "⚠️ Vence hoy" : `⚠️ Plazo vencido hace ${Math.abs(sem.dias)} día${Math.abs(sem.dias)>1?'s':''}`}
        </div>
        <div style="font-size:11px;color:${sem.color};opacity:0.8;">Límite ${CONFIG.plazoDerivacionDias||15} días · Vence el ${formatFecha(sem.vencimiento.toISOString())}</div>
      </div>
    </div>` : "";

  cont.innerHTML = `${semBanner}<div style="display:flex;flex-direction:column;gap:12px;padding:14px;">

    <!-- Datos básicos: solicitante, fecha, nro -->
    <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:10px 14px;">
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;font-size:13px;margin-bottom:6px;">
        <div><span style="color:#888;font-size:11px;display:block;">Nro Solicitud</span><strong>${sol.NroSolicitud}</strong></div>
        <div><span style="color:#888;font-size:11px;display:block;">Fecha</span>${formatFecha(sol.FechaRecepcion)}</div>
      </div>
      <div style="font-size:13px;margin-bottom:4px;"><span style="color:#888;font-size:11px;display:block;">Solicitante</span>${sol.Solicitante}</div>
      <div style="font-size:13px;"><span style="color:#888;font-size:11px;display:block;">Dirección</span>${sol.Direccion||"-"}</div>
      ${sol.UnidadDerivada?`<div style="margin-top:8px;padding:5px 10px;background:#fef3c7;border-radius:6px;font-size:12px;color:#b45309;">🏢 Derivada a: <strong>${sol.UnidadDerivada}</strong></div>`:""}
    </div>

    <!-- DESCRIPCIÓN DE LA SOLICITUD — siempre visible y prominente -->
    <div style="background:white;border:2px solid var(--azul);border-radius:10px;overflow:hidden;">
      <div style="background:linear-gradient(90deg,#1e3a5f,#1a3a6b);color:white;padding:8px 14px;font-size:12px;font-weight:700;letter-spacing:0.3px;">
        📋 Descripción de la Solicitud
      </div>
      <div style="padding:12px 14px;font-size:13px;color:#374151;line-height:1.7;min-height:48px;">
        ${sol.Solicitud ? sol.Solicitud : '<span style="color:#9ca3af;font-style:italic;">Sin descripción registrada</span>'}
      </div>
    </div>

    ${esDevuelta ? `
    <!-- Motivo devolución -->
    <div class="accion-panel">
      <div class="accion-header devuelta">↩️ Motivo de Devolución</div>
      <div class="accion-body" style="background:#fff5f5;">
        <p style="font-size:13px;color:#b91c1c;margin:0;">${sol.MotivoDevolucion||"Sin motivo registrado"}</p>
      </div>
    </div>` : ""}

    ${tieneEvidencia ? `
    <!-- SOLUCIÓN EJECUTADA POR LA UNIDAD -->
    <div style="background:white;border:2px solid #15803d;border-radius:10px;overflow:hidden;">
      <div style="background:linear-gradient(90deg,#14532d,#15803d);color:white;padding:8px 14px;font-size:12px;font-weight:700;letter-spacing:0.5px;text-transform:uppercase;">
        🏢 Solución Ejecutada por la Unidad
      </div>
      <div id="dir-evidencia-panel" style="padding:0;">
        <div style="text-align:center;color:#9ca3af;font-size:12px;padding:16px;">
          <div class="spinner" style="margin:0 auto 8px;width:20px;height:20px;border-width:2px;"></div>
          Cargando solución...
        </div>
      </div>
    </div>` : ""}

    ${esDerivable ? `
    <!-- Panel derivar -->
    <div class="accion-panel">
      <div class="accion-header derivar">📤 ${esDevuelta?"Re-derivar":"Derivar"} Solicitud</div>
      <div class="accion-body">
        <label style="font-size:11px;font-weight:700;color:#666;text-transform:uppercase;letter-spacing:0.5px;display:block;margin-bottom:6px;">Unidad de destino</label>
        <select id="dir-unidad-derivar">
          <option value="">— Seleccionar unidad —</option>
          ${CONFIG.unidades.map(u=>`<option ${sol.UnidadDerivada===u?'selected':''}>${u}</option>`).join("")}
        </select>
        <label style="font-size:11px;font-weight:700;color:#666;text-transform:uppercase;letter-spacing:0.5px;display:block;margin-bottom:6px;">Instrucciones / Acciones</label>
        <textarea id="dir-accion-obs" rows="3" placeholder="Instrucciones específicas para la unidad..."></textarea>
        <button class="btn-primary" onclick="derivarSolicitud('${sol.id}')" style="width:100%;padding:12px;">
          📤 ${esDevuelta?"Re-derivar Solicitud":"Derivar Solicitud"}
        </button>
      </div>
    </div>` : ""}

    ${esCerrable ? `
    <!-- Panel resolución -->
    <div class="accion-panel">
      <div class="accion-header cerrar" style="background:linear-gradient(90deg,#14532d,#15803d);">📋 Resolución de Solicitud</div>
      <div class="accion-body">

        ${!esPendienteCierre ? `
        <!-- Desde Respondida: marcar pendiente o cerrar directo -->
        <div style="background:#fdf4ff;border:1.5px solid #d8b4fe;border-radius:10px;padding:12px;margin-bottom:10px;">
          <div style="font-size:12px;font-weight:700;color:#7e22ce;margin-bottom:6px;">⏳ Marcar Pendiente de Cierre</div>
          <p style="font-size:11px;color:#6b21a8;margin:0 0 8px;">Queda en evaluación hasta generar el parte final (inspección, multa o solución).</p>
          <div style="display:flex;gap:8px;align-items:center;margin-bottom:8px;">
            <label style="font-size:11px;font-weight:600;color:#7e22ce;white-space:nowrap;">Plazo hasta:</label>
            <input type="date" id="dir-plazo-cierre" style="flex:1;padding:5px 8px;border:1.5px solid #d8b4fe;border-radius:6px;font-size:12px;">
          </div>
          <button onclick="pendienteCierreSolicitud('${sol.id}')"
            style="width:100%;padding:9px;background:#7e22ce;color:white;border:none;border-radius:8px;cursor:pointer;font-size:13px;font-weight:600;">
            ⏳ Marcar Pendiente de Cierre
          </button>
        </div>
        <div style="background:#f0fdf4;border:1.5px solid #86efac;border-radius:10px;padding:12px;">
          <div style="font-size:12px;font-weight:700;color:#15803d;margin-bottom:6px;">🔒 Cierre Directo</div>
          <p style="font-size:11px;color:#166534;margin:0 0 8px;">Cierra formalmente sin pasar por pendiente.</p>
          <textarea id="dir-cierre-obs" rows="2" placeholder="Observaciones (opcional)..." style="width:100%;padding:7px 9px;border:1.5px solid #86efac;border-radius:7px;font-size:12px;resize:vertical;box-sizing:border-box;margin-bottom:8px;"></textarea>
          <button class="btn-success" onclick="cerrarSolicitud('${sol.id}')" style="width:100%;padding:9px;">
            🔒 Cerrar Solicitud Formalmente
          </button>
        </div>` : `

        <!-- Desde Pendiente de Cierre: Solución o Multa -->
        <p style="font-size:12px;color:#555;margin:0 0 12px;">⏳ En plazo de evaluación. Selecciona cómo se resolvió:</p>

        <!-- OPCIÓN A: Solución ejecutada -->
        <div style="background:#f0fdf4;border:1.5px solid #86efac;border-radius:10px;padding:12px;margin-bottom:10px;">
          <div style="font-size:12px;font-weight:700;color:#15803d;margin-bottom:8px;">✅ Se Solucionó</div>
          <label style="font-size:10px;font-weight:700;color:#166534;text-transform:uppercase;letter-spacing:0.5px;display:block;margin-bottom:4px;">N° Parte / Informe <span style="color:#ef4444;">*</span></label>
          <input type="text" id="dir-cierre-parte"
            placeholder="Ej: Parte N° 123-2026"
            style="width:100%;padding:7px 9px;border:1.5px solid #86efac;border-radius:7px;font-size:12px;box-sizing:border-box;margin-bottom:8px;">
          <label style="font-size:10px;font-weight:700;color:#166534;text-transform:uppercase;letter-spacing:0.5px;display:block;margin-bottom:4px;">Descripción de la solución <span style="color:#ef4444;">*</span></label>
          <textarea id="dir-cierre-solucion-obs" rows="2"
            placeholder="Detalla cómo se resolvió (obra ejecutada, inspección realizada, etc.)"
            style="width:100%;padding:7px 9px;border:1.5px solid #86efac;border-radius:7px;font-size:12px;resize:vertical;box-sizing:border-box;margin-bottom:8px;"></textarea>
          <button onclick="cerrarPorSolucion('${sol.id}')"
            style="width:100%;padding:10px;background:#15803d;color:white;border:none;border-radius:8px;cursor:pointer;font-size:13px;font-weight:700;">
            ✅ Cerrar — Solución Ejecutada
          </button>
        </div>

        <!-- OPCIÓN B: Multa aplicada -->
        <div style="background:#fefce8;border:1.5px solid #fde047;border-radius:10px;padding:12px;">
          <div style="font-size:12px;font-weight:700;color:#854d0e;margin-bottom:8px;">📋 Se Aplicó Multa</div>
          <label style="font-size:10px;font-weight:700;color:#854d0e;text-transform:uppercase;letter-spacing:0.5px;display:block;margin-bottom:4px;">N° Informe de Multa <span style="color:#ef4444;">*</span></label>
          <input type="text" id="dir-multa-nro"
            placeholder="Ej: Informe N° 045-2026"
            style="width:100%;padding:7px 9px;border:1.5px solid #fde047;border-radius:7px;font-size:12px;box-sizing:border-box;margin-bottom:8px;">
          <label style="font-size:10px;font-weight:700;color:#854d0e;text-transform:uppercase;letter-spacing:0.5px;display:block;margin-bottom:4px;">Monto de la multa (UTM)</label>
          <input type="text" id="dir-multa-monto"
            placeholder="Ej: 5 UTM (opcional)"
            style="width:100%;padding:7px 9px;border:1.5px solid #fde047;border-radius:7px;font-size:12px;box-sizing:border-box;margin-bottom:8px;">
          <label style="font-size:10px;font-weight:700;color:#854d0e;text-transform:uppercase;letter-spacing:0.5px;display:block;margin-bottom:4px;">Observaciones <span style="color:#ef4444;">*</span></label>
          <textarea id="dir-multa-obs" rows="2"
            placeholder="Motivo y detalle de la multa aplicada"
            style="width:100%;padding:7px 9px;border:1.5px solid #fde047;border-radius:7px;font-size:12px;resize:vertical;box-sizing:border-box;margin-bottom:8px;"></textarea>
          <button onclick="cerrarPorMulta('${sol.id}')"
            style="width:100%;padding:10px;background:#b45309;color:white;border:none;border-radius:8px;cursor:pointer;font-size:13px;font-weight:700;">
            📋 Cerrar — Multa Aplicada
          </button>
        </div>`}

      </div>
    </div>` : ""}

    ${esPendienteCierre ? `
    <div style="background:#fdf4ff;border:1.5px solid #d8b4fe;border-radius:10px;padding:12px 14px;display:flex;align-items:center;gap:10px;margin-bottom:8px;">
      <span style="font-size:20px;">⏳</span>
      <div style="flex:1;">
        <div style="font-size:13px;font-weight:700;color:#7e22ce;">Pendiente de Cierre</div>
        <div style="font-size:11px;color:#6b21a8;">Aguardando parte final${sol.FechaCierre ? ' — Plazo: ' + formatFecha(sol.FechaCierre) : ''}</div>
      </div>
    </div>` : ""}

    ${sol.Estado === CONFIG.estados.CERRADA ? `
    <div style="background:#f0fdf4;border:1.5px solid #86efac;border-radius:10px;padding:12px 14px;display:flex;align-items:center;gap:10px;">
      <span style="font-size:20px;">🔒</span>
      <div>
        <div style="font-size:13px;font-weight:700;color:#15803d;">Solicitud Cerrada Formalmente</div>
        <div style="font-size:11px;color:#16a34a;">Caso completado y archivado en el sistema</div>
      </div>
    </div>` : ""}

    ${sol.Estado === CONFIG.estados.INGRESADA ? `
    <!-- Cierre directo por Director — solo disponible antes de derivar -->
    <div class="accion-panel">
      <div class="accion-header" style="background:linear-gradient(90deg,#7f1d1d,#b91c1c);color:white;padding:8px 14px;font-size:12px;font-weight:700;">
        🚫 Cierre Directo — No Corresponde a DOM
      </div>
      <div class="accion-body" style="background:#fff5f5;">
        <p style="font-size:12px;color:#7f1d1d;margin:0 0 10px;">
          Cierra la solicitud inmediatamente indicando que no corresponde a la Dirección de Obras. Queda registro en el historial.
        </p>
        <textarea id="dir-no-corresponde-obs" rows="2"
          placeholder="Motivo (ej: Corresponde a DIDECO, solicitud duplicada, fuera de competencia DOM...)"
          style="width:100%;padding:8px 10px;border:1.5px solid #fca5a5;border-radius:8px;font-size:12px;resize:vertical;box-sizing:border-box;margin-bottom:10px;"></textarea>
        <button onclick="cerrarDirectoDirector('${sol.id}')"
          style="width:100%;padding:11px;background:#b91c1c;color:white;border:none;border-radius:8px;cursor:pointer;font-size:13px;font-weight:700;letter-spacing:0.2px;">
          🚫 Cerrar — No Corresponde a DOM
        </button>
      </div>
    </div>` : ""}

    <!-- Historial -->
    <div class="form-section">
      <div class="form-section-header verde" style="font-size:11px;cursor:pointer;" onclick="cargarHistorialDir('${sol.NroSolicitud}')">
        🕐 Historial de Movimientos <span style="float:right;font-weight:400;">ver ▼</span>
      </div>
      <div class="form-section-body" id="dir-historial" style="max-height:200px;overflow-y:auto;padding:8px;">
        <div style="text-align:center;color:#9ca3af;font-size:12px;padding:10px;">Clic en "ver" para cargar</div>
      </div>
    </div>

  </div>`;

  // Setear fecha plazo siempre en tiempo real (evita caché del template)
  const inputPlazo = document.getElementById("dir-plazo-cierre");
  if (inputPlazo) inputPlazo.value = new Date().toISOString().split('T')[0];

  // Cargar historial automáticamente
  cargarHistorialDir(sol.NroSolicitud);

  // Cargar evidencia si está Respondida o Cerrada
  if (tieneEvidencia) cargarEvidenciaDir(sol);
}

async function cargarEvidenciaDir(sol) {
  window._solDir = sol;
  const panel = document.getElementById("dir-evidencia-panel");
  if (!panel) return;
  try {
    const evidencias = await getEvidenciasBySolicitud(sol.NroSolicitud, sol.id);
    if (!evidencias.length) {
      panel.innerHTML = `
        <div style="text-align:center;padding:20px;color:#9ca3af;">
          <div style="font-size:32px;margin-bottom:8px;">📭</div>
          <p style="font-size:13px;">Sin evidencia registrada por la unidad</p>
        </div>`;
      return;
    }
    panel.innerHTML = "";
    for (const ev of evidencias) {
      const evWrap = document.createElement("div");
      evWrap.style.cssText = "border-bottom:2px solid #e2e8f0;padding-bottom:14px;margin-bottom:14px;";
      evWrap.innerHTML = `
        <div style="display:flex;align-items:center;gap:8px;padding:10px 14px;background:#f8fafc;border-bottom:1px solid #e2e8f0;">
          <div style="width:32px;height:32px;border-radius:50%;background:var(--azul);color:white;display:flex;align-items:center;justify-content:center;font-size:14px;flex-shrink:0;">🏢</div>
          <div>
            <div style="font-size:13px;font-weight:700;color:#1a3a6b;">${ev.Unidad||"Unidad"}</div>
            <div style="font-size:11px;color:#888;">👤 ${ev.Responsable||""} &nbsp;·&nbsp; 📅 ${formatFecha(ev.FechaCarga)}</div>
          </div>
        </div>
        <div style="padding:12px 14px 8px;">
          <div style="font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;letter-spacing:0.4px;margin-bottom:6px;">📝 Descripción de la solución</div>
          <div style="font-size:13px;color:#374151;line-height:1.7;background:#f9fafb;border-left:3px solid var(--azul);padding:10px 12px;border-radius:0 8px 8px 0;">${ev.DescripcionEvidencia||"Sin descripción."}</div>
        </div>
        <div id="ev-media-${ev.id}" style="padding:0 14px 4px;"></div>`;
      panel.appendChild(evWrap);

      // Cargar adjuntos (imágenes y PDFs de la evidencia)
      const mediaCont = evWrap.querySelector(`#ev-media-${ev.id}`);
      getListItemAttachments(CONFIG.lists.evidencias, ev.id).then(async atts => {
        if (!atts.length) {
          mediaCont.innerHTML = `<p style="font-size:12px;color:#d1d5db;margin:0;padding:4px 0;">Sin archivos adjuntos</p>`;
          return;
        }
        mediaCont.innerHTML = `<div style="font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;letter-spacing:0.4px;margin-bottom:8px;">📎 Archivos adjuntos (${atts.length})</div>`;

        const pdfs = atts.filter(a => a.name?.toLowerCase().endsWith('.pdf'));
        const imgs = atts.filter(a => /\.(jpg|jpeg|png|gif)$/i.test(a.name||''));

        // PDFs con PDF.js
        for (const pdf of pdfs) {
          const pdfBlock = document.createElement("div");
          pdfBlock.style.cssText = "margin-bottom:10px;";
          pdfBlock.innerHTML = `<div style="font-size:11px;color:#888;margin-bottom:4px;display:flex;align-items:center;gap:4px;">📄 <strong>${pdf.name}</strong></div>`;
          mediaCont.appendChild(pdfBlock);
          try {
            const blobUrl = await getAttachmentBlobUrl(pdf.downloadUrl, pdf.serverRelativeUrl);
            if (typeof pdfjsLib !== "undefined") {
              pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
              const pdfDoc = await pdfjsLib.getDocument(blobUrl).promise;
              const wrap = document.createElement("div");
              wrap.style.cssText = "background:#525659;border-radius:8px;padding:8px;display:flex;flex-direction:column;gap:4px;";
              pdfBlock.appendChild(wrap);
              for (let p=1; p<=pdfDoc.numPages; p++) {
                const page = await pdfDoc.getPage(p);
                const w = (mediaCont.clientWidth||300)-16;
                const vp = page.getViewport({scale:1});
                const svp = page.getViewport({scale:w/vp.width});
                const canvas = document.createElement("canvas");
                canvas.width=svp.width; canvas.height=svp.height;
                canvas.style.cssText="width:100%;border-radius:4px;";
                await page.render({canvasContext:canvas.getContext("2d"),viewport:svp}).promise;
                wrap.appendChild(canvas);
              }
            }
          } catch {}
        }

        // Grid de imágenes (fotos de la ejecución)
        if (imgs.length) {
          const gridLabel = document.createElement("div");
          gridLabel.style.cssText = "font-size:11px;color:#888;margin-bottom:6px;display:flex;align-items:center;gap:4px;";
          gridLabel.innerHTML = `📸 <strong>${imgs.length} foto${imgs.length>1?'s':''} de la ejecución</strong>`;
          mediaCont.appendChild(gridLabel);

          const grid = document.createElement("div");
          grid.style.cssText = `display:grid;grid-template-columns:repeat(${Math.min(imgs.length,2)},1fr);gap:6px;margin-bottom:8px;`;
          mediaCont.appendChild(grid);

          for (const img of imgs) {
            const box = document.createElement("div");
            box.style.cssText = "border-radius:8px;overflow:hidden;aspect-ratio:4/3;background:#f3f4f6;border:2px solid var(--borde);cursor:zoom-in;transition:border-color 0.2s;position:relative;";
            box.title = `Ver: ${img.name}`;
            box.onmouseenter = () => box.style.borderColor = "var(--azul)";
            box.onmouseleave = () => box.style.borderColor = "var(--borde)";
            grid.appendChild(box);

            // Placeholder mientras carga
            box.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#d1d5db;"><div class="spinner" style="width:20px;height:20px;border-width:2px;"></div></div>`;

            getAttachmentBlobUrl(img.downloadUrl, img.serverRelativeUrl).then(blobUrl => {
              const imgEl = document.createElement("img");
              imgEl.src = blobUrl;
              imgEl.style.cssText = "width:100%;height:100%;object-fit:cover;";
              imgEl.onclick = () => abrirLightbox(blobUrl, img.name);
              box.innerHTML = "";
              box.appendChild(imgEl);
              // Badge zoom
              const badge = document.createElement("div");
              badge.style.cssText = "position:absolute;bottom:4px;right:4px;background:rgba(0,0,0,0.5);color:white;border-radius:4px;padding:2px 6px;font-size:10px;";
              badge.textContent = "🔍 ver";
              box.appendChild(badge);
            }).catch(() => {
              box.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#fca5a5;font-size:11px;">⚠️ Error</div>`;
            });
          }
        }
      }).catch(() => {
        mediaCont.innerHTML = `<p style="font-size:12px;color:#fca5a5;">Error al cargar adjuntos</p>`;
      });
    }
  } catch(e) {
    console.warn("Evidencia no disponible:", e?.message || e);
    // Si la lista no existe o no hay permisos, mostrar "sin evidencia" en lugar de error
    const esErrorLista = e?.message?.includes("404") || e?.message?.includes("does not exist") || e?.message?.includes("no existe");
    if (panel) {
      if (esErrorLista) {
        panel.innerHTML = `
          <div style="text-align:center;padding:20px;color:#9ca3af;">
            <div style="font-size:32px;margin-bottom:8px;">📭</div>
            <p style="font-size:13px;">Sin evidencia registrada por la unidad</p>
          </div>`;
      } else {
        panel.innerHTML = `
          <div style="padding:14px;text-align:center;">
            <div style="font-size:28px;margin-bottom:8px;">⚠️</div>
            <p style="color:#b45309;font-size:13px;margin-bottom:10px;">No se pudo cargar la respuesta de la unidad</p>
            <p style="color:#9ca3af;font-size:11px;margin-bottom:12px;word-break:break-all;">${e?.message||""}</p>
            <button onclick="cargarEvidenciaDir(window._solDir)"
              style="padding:7px 18px;border:1.5px solid var(--azul);border-radius:7px;background:white;color:var(--azul);cursor:pointer;font-size:13px;font-weight:600;">
              🔄 Reintentar
            </button>
          </div>`;
      }
    }
  }
}

async function abrirAdjuntoUnidad(downloadUrl, serverRelativeUrl, nombre, isPdf) {
  try {
    showLoading("Cargando documento...");
    const blobUrl = await getAttachmentBlobUrl(downloadUrl, serverRelativeUrl);
    hideLoading();
    if (isPdf) {
      // Mostrar en el panel PDF central
      const pdfPanel = document.getElementById("uni-pdf");
      const pdfHeader = document.getElementById("uni-pdf-header");
      if (pdfPanel) {
        if (pdfHeader) pdfHeader.textContent = `📄 ${nombre}`;
        pdfPanel.innerHTML = "";
        pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
        const wrap = document.createElement("div");
        wrap.style.cssText = "flex:1;overflow-y:auto;background:#525659;border-radius:8px;padding:10px;display:flex;flex-direction:column;align-items:center;gap:6px;min-height:0;";
        pdfPanel.appendChild(wrap);
        const pdf = await pdfjsLib.getDocument(blobUrl).promise;
        for (let p = 1; p <= pdf.numPages; p++) {
          const page = await pdf.getPage(p);
          const w = (wrap.clientWidth || 350) - 20;
          const vp = page.getViewport({ scale: 1 });
          const canvas = document.createElement("canvas");
          const svp = page.getViewport({ scale: w / vp.width });
          canvas.width = svp.width; canvas.height = svp.height;
          canvas.style.cssText = "width:100%;border-radius:4px;box-shadow:0 2px 8px rgba(0,0,0,0.3);";
          await page.render({ canvasContext: canvas.getContext("2d"), viewport: svp }).promise;
          wrap.appendChild(canvas);
        }
      } else {
        window.open(blobUrl, "_blank");
      }
    } else {
      abrirLightbox(blobUrl, nombre);
    }
  } catch(e) {
    hideLoading();
    showToast("error", "No se pudo abrir el archivo");
  }
}

function abrirLightbox(blobUrl, nombre) {
  const overlay = document.createElement("div");
  overlay.style.cssText = "position:fixed;inset:0;background:rgba(0,0,0,0.92);z-index:9999;display:flex;flex-direction:column;align-items:center;justify-content:center;cursor:zoom-out;";
  overlay.innerHTML = `
    <div style="position:absolute;top:16px;right:20px;color:white;font-size:24px;cursor:pointer;opacity:0.7;" onclick="this.closest('div').remove()">✕</div>
    <img src="${blobUrl}" style="max-width:90vw;max-height:85vh;object-fit:contain;border-radius:8px;box-shadow:0 8px 32px rgba(0,0,0,0.5);">
    <div style="margin-top:12px;color:rgba(255,255,255,0.6);font-size:13px;">${nombre}</div>`;
  overlay.onclick = e => { if (e.target === overlay) overlay.remove(); };
  document.body.appendChild(overlay);
}

async function cargarHistorialDir(nroSolicitud, containerId = "dir-historial") {
  const hc = document.getElementById(containerId);
  if (!hc) return;
  hc.innerHTML = `<div style="text-align:center;color:#9ca3af;font-size:12px;padding:10px;"><div class="spinner" style="margin:0 auto 8px;width:20px;height:20px;border-width:2px;"></div>Cargando...</div>`;
  try {
    // Cargar historial y evidencias en paralelo
    const sol = state.solicitudSeleccionada;
    const [hist, evidencias] = await Promise.all([
      getHistorialBySolicitud(nroSolicitud),
      sol ? getEvidenciasBySolicitud(nroSolicitud, sol.id).catch(()=>[]) : Promise.resolve([])
    ]);

    // Unificar en línea de tiempo: historial + respuestas de unidad
    const eventos = [
      ...hist.map(h => ({
        tipo: 'historial',
        fecha: h.FechaAccion || h.Modified || "",
        accion: h.Accion || h.Title || "—",
        usuario: h.UsuarioAccion || "",
        rol: h.RolUsuario || "",
        unidad: h.Unidad || "",
        obs: h.Observaciones || "",
        estAnt: h.EstadoAnterior || "",
        estNuevo: h.EstadoNuevo || ""
      })),
      ...evidencias.map(e => ({
        tipo: 'respuesta',
        fecha: e.FechaCarga || "",
        accion: `Respuesta de ${e.Unidad || "Unidad"}`,
        usuario: e.Responsable || "",
        rol: "",
        unidad: e.Unidad || "",
        obs: e.DescripcionEvidencia || "",
        estAnt: "", estNuevo: "",
        evId: e.id
      }))
    ];

    eventos.sort((a,b) => new Date(b.fecha) - new Date(a.fecha));

    if (!eventos.length) { hc.innerHTML=`<p style="text-align:center;color:#9ca3af;font-size:12px;">Sin historial</p>`; return; }

    const dot = (tipo, accion) => {
      if (tipo === 'respuesta') return '#15803d';
      const a = (accion||"").toLowerCase();
      return a.includes('deriv')?'#f59e0b':a.includes('respond')?'#22c55e':a.includes('devuel')?'#ef4444':a.includes('cerr')?'#6b7280':a.includes('pend')?'#7e22ce':'#3b82f6';
    };

    hc.innerHTML = eventos.map(ev => `
      <div style="display:flex;gap:8px;padding:8px 0;border-bottom:1px solid #f3f4f6;">
        <div style="width:9px;height:9px;border-radius:50%;background:${dot(ev.tipo,ev.accion)};margin-top:4px;flex-shrink:0;"></div>
        <div style="flex:1;">
          <div style="font-size:12px;font-weight:700;color:${ev.tipo==='respuesta'?'#15803d':'#1a1a1a'};">
            ${ev.tipo==='respuesta'?'🏢 ':''}${ev.accion}
          </div>
          <div style="font-size:11px;color:#888;">
            ${formatFechaHora(ev.fecha)}${ev.usuario?' · '+ev.usuario:''}${ev.unidad?' · <strong>'+ev.unidad+'</strong>':''}
          </div>
          ${ev.estAnt?`<div style="font-size:10px;color:#aaa;margin-top:1px;">${ev.estAnt} → ${ev.estNuevo}</div>`:""}
          ${ev.obs?`<div style="font-size:12px;color:#374151;background:${ev.tipo==='respuesta'?'#f0fdf4':'#f8fafc'};border-left:3px solid ${dot(ev.tipo,ev.accion)};padding:6px 8px;border-radius:0 6px 6px 0;margin-top:4px;line-height:1.5;">${ev.obs}</div>`:""}
          ${ev.tipo==='respuesta'?`<div id="dir-ev-atts-${ev.evId}" style="margin-top:4px;font-size:11px;color:#9ca3af;">Cargando adjuntos...</div>`:""}
        </div>
      </div>`).join("");

    // Cargar adjuntos de cada evidencia
    for (const ev of eventos.filter(e=>e.tipo==='respuesta')) {
      const cont = document.getElementById(`dir-ev-atts-${ev.evId}`);
      if (!cont) continue;
      getListItemAttachments(CONFIG.lists.evidencias, ev.evId).then(atts => {
        if (!atts.length) { cont.remove(); return; }
        const imgs = atts.filter(a=>/\.(jpg|jpeg|png|gif)$/i.test(a.name||''));
        const pdfs = atts.filter(a=>a.name?.toLowerCase().endsWith('.pdf'));
        cont.innerHTML = `<div style="font-weight:600;color:#6b7280;margin-bottom:4px;">📎 ${atts.length} adjunto${atts.length>1?'s':''}</div>`;
        if (imgs.length) {
          const grid = document.createElement("div");
          grid.style.cssText = `display:grid;grid-template-columns:repeat(${Math.min(imgs.length,3)},1fr);gap:4px;margin-bottom:4px;`;
          cont.appendChild(grid);
          imgs.forEach(img => {
            const box = document.createElement("div");
            box.style.cssText = "border-radius:6px;overflow:hidden;aspect-ratio:4/3;background:#f3f4f6;cursor:zoom-in;";
            box.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;height:100%;"><div class="spinner" style="width:16px;height:16px;border-width:2px;"></div></div>`;
            grid.appendChild(box);
            getAttachmentBlobUrl(img.downloadUrl, img.serverRelativeUrl).then(url => {
              const el = document.createElement("img");
              el.src = url; el.style.cssText = "width:100%;height:100%;object-fit:cover;";
              el.onclick = () => abrirLightbox(url, img.name);
              box.innerHTML = ""; box.appendChild(el);
            }).catch(()=>{ box.innerHTML=`<div style="font-size:10px;color:#fca5a5;text-align:center;padding:4px;">⚠️</div>`; });
          });
        }
        pdfs.forEach(pdf => {
          const row = document.createElement("div");
          row.style.cssText = "font-size:11px;padding:4px 6px;background:#f8fafc;border:1px solid var(--borde);border-radius:4px;margin-bottom:3px;cursor:pointer;color:var(--azul);";
          row.innerHTML = `📄 ${pdf.name}`;
          row.onclick = async () => { try { window.open(await getAttachmentBlobUrl(pdf.downloadUrl, pdf.serverRelativeUrl),"_blank"); } catch{} };
          cont.appendChild(row);
        });
      }).catch(()=>{ if(cont) cont.remove(); });
    }
  } catch(e) {
    hc.innerHTML=`<p style="color:#9ca3af;font-size:12px;text-align:center;">Error al cargar historial</p>`;
  }
}

async function derivarSolicitud(solId) {
  const unidad = document.getElementById("dir-unidad-derivar")?.value;
  const obs    = document.getElementById("dir-accion-obs")?.value.trim();
  if (!unidad) { showToast("error","Selecciona una unidad de destino"); return; }
  showLoading("Derivando solicitud...");
  try {
    const sol = state.solicitudes.find(s=>s.id===solId);
    const esRederivar = sol.Estado === CONFIG.estados.DEVUELTA;
    const updateFields = { Estado:CONFIG.estados.DERIVADA, UnidadDerivada:unidad, FechaDerivacion:new Date().toISOString() };
    if (obs) updateFields.Acciones = obs;
    await actualizarSolicitud(solId, updateFields);
    registrarHistorial({
      NroSolicitud:sol.NroSolicitud,
      Title:esRederivar ? "Re-derivada a unidad" : "Derivada a unidad",
      EstadoAnterior:sol.Estado, EstadoNuevo:CONFIG.estados.DERIVADA,
      UsuarioAccion:state.usuario.NombreCompleto, RolUsuario:state.usuario.Rol,
      Unidad:unidad, FechaAccion:new Date().toISOString(), Observaciones:obs
    }).catch(e => console.warn("Historial (no crítico):", e.message));
    notificarUnidad({...sol,Estado:CONFIG.estados.DERIVADA},unidad).catch(console.error);
    showToast("success",`✅ Derivada a ${unidad}`);
    await renderDirector();
  } catch(e) { showToast("error","Error: "+e.message); }
  finally { hideLoading(); }
}

async function pendienteCierreSolicitud(solId) {
  const obs   = document.getElementById("dir-cierre-obs")?.value.trim();
  const plazo = document.getElementById("dir-plazo-cierre")?.value;
  if (!confirm("¿Confirmas marcar esta solicitud como Pendiente de Cierre?\nQuedará en evaluación hasta que se genere el parte final.")) return;
  showLoading("Actualizando estado...");
  try {
    const sol = state.solicitudes.find(s=>s.id===solId);
    const pcFields = { Estado: CONFIG.estados.PENDIENTE_CIERRE };
    if (plazo) pcFields.FechaCierre = new Date(plazo + "T12:00:00").toISOString();
    await actualizarSolicitud(solId, pcFields);
    registrarHistorial({
      NroSolicitud:sol.NroSolicitud, Title:"Pendiente de Cierre — en plazo de evaluación",
      EstadoAnterior:sol.Estado, EstadoNuevo:CONFIG.estados.PENDIENTE_CIERRE,
      UsuarioAccion:state.usuario.NombreCompleto, RolUsuario:state.usuario.Rol,
      Unidad:state.usuario.Unidad, FechaAccion:new Date().toISOString(),
      Observaciones: (obs ? obs + (plazo ? ` | Plazo: ${plazo}` : "") : (plazo ? `Plazo: ${plazo}` : ""))
    }).catch(e => console.warn("Historial (no crítico):", e.message));
    showToast("success","⏳ Solicitud marcada como Pendiente de Cierre");
    await renderDirector();
  } catch(e) { showToast("error","Error: "+e.message); }
  finally { hideLoading(); }
}

async function cerrarSolicitud(solId) {
  const obs = document.getElementById("dir-cierre-obs")?.value.trim();
  if (!confirm("¿Confirmas el cierre formal de esta solicitud?")) return;
  showLoading("Cerrando solicitud...");
  try {
    const sol = state.solicitudes.find(s=>s.id===solId);
    await actualizarSolicitud(solId, { Estado:CONFIG.estados.CERRADA });
    registrarHistorial({
      NroSolicitud:sol.NroSolicitud, Title:"Solicitud cerrada formalmente",
      EstadoAnterior:sol.Estado, EstadoNuevo:CONFIG.estados.CERRADA,
      UsuarioAccion:state.usuario.NombreCompleto, RolUsuario:state.usuario.Rol,
      Unidad:state.usuario.Unidad, FechaAccion:new Date().toISOString(), Observaciones:obs
    }).catch(e => console.warn("Historial (no crítico):", e.message));
    showToast("success","🔒 Solicitud cerrada formalmente");
    await renderDirector();
  } catch(e) { showToast("error","Error: "+e.message); }
  finally { hideLoading(); }
}

async function cerrarPorSolucion(solId) {
  const parte = document.getElementById("dir-cierre-parte")?.value.trim();
  const obs   = document.getElementById("dir-cierre-solucion-obs")?.value.trim();
  if (!parte) { showToast("error", "Ingresa el N° de Parte o Informe."); return; }
  if (!obs)   { showToast("error", "Describe la solución ejecutada."); return; }
  if (!confirm(`¿Cerrar solicitud como "Solución Ejecutada"?\nParte: ${parte}`)) return;
  showLoading("Cerrando solicitud...");
  try {
    const sol = state.solicitudes.find(s => s.id === solId);
    await actualizarSolicitud(solId, { Estado: CONFIG.estados.CERRADA });
    await registrarHistorial({
      NroSolicitud: sol.NroSolicitud,
      Title: `Cierre por Solución — ${parte}`,
      EstadoAnterior: sol.Estado, EstadoNuevo: CONFIG.estados.CERRADA,
      UsuarioAccion: state.usuario.NombreCompleto, RolUsuario: state.usuario.Rol,
      Unidad: state.usuario.Unidad, FechaAccion: new Date().toISOString(),
      Observaciones: `Parte/Informe: ${parte} | ${obs}`
    });
    showToast("success", `✅ Solicitud ${sol.NroSolicitud} cerrada — Solución ejecutada`);
    await renderDirector();
  } catch(e) { showToast("error", "Error: " + e.message); }
  finally { hideLoading(); }
}

async function cerrarPorMulta(solId) {
  const nro   = document.getElementById("dir-multa-nro")?.value.trim();
  const monto = document.getElementById("dir-multa-monto")?.value.trim();
  const obs   = document.getElementById("dir-multa-obs")?.value.trim();
  if (!nro) { showToast("error", "Ingresa el N° de Informe de Multa."); return; }
  if (!obs) { showToast("error", "Ingresa las observaciones de la multa."); return; }
  if (!confirm(`¿Cerrar solicitud como "Multa Aplicada"?\nInforme: ${nro}${monto ? ' | ' + monto : ''}`)) return;
  showLoading("Cerrando solicitud...");
  try {
    const sol = state.solicitudes.find(s => s.id === solId);
    await actualizarSolicitud(solId, { Estado: CONFIG.estados.CERRADA });
    await registrarHistorial({
      NroSolicitud: sol.NroSolicitud,
      Title: `Cierre por Multa — ${nro}`,
      EstadoAnterior: sol.Estado, EstadoNuevo: CONFIG.estados.CERRADA,
      UsuarioAccion: state.usuario.NombreCompleto, RolUsuario: state.usuario.Rol,
      Unidad: state.usuario.Unidad, FechaAccion: new Date().toISOString(),
      Observaciones: `Informe multa: ${nro}${monto ? ' | Monto: ' + monto : ''} | ${obs}`
    });
    showToast("success", `📋 Solicitud ${sol.NroSolicitud} cerrada — Multa aplicada`);
    await renderDirector();
  } catch(e) { showToast("error", "Error: " + e.message); }
  finally { hideLoading(); }
}

async function reabrirSolicitud(solId) {
  const nuevoEstado = document.getElementById("dir-reabrir-estado")?.value;
  const obs         = document.getElementById("dir-reabrir-obs")?.value.trim();
  if (!nuevoEstado) { showToast("error", "Selecciona el estado al que deseas cambiar"); return; }
  if (!obs)         { showToast("error", "Ingresa el motivo de reapertura"); return; }
  if (!confirm(`¿Confirmas cambiar el estado a "${nuevoEstado}"?\nMotivo: ${obs}`)) return;
  showLoading("Actualizando estado...");
  try {
    const sol = state.solicitudes.find(s => s.id === solId);
    await actualizarSolicitud(solId, { Estado: nuevoEstado });
    registrarHistorial({
      NroSolicitud: sol.NroSolicitud,
      Title: `Reabierta — cambiada a ${nuevoEstado}`,
      EstadoAnterior: sol.Estado, EstadoNuevo: nuevoEstado,
      UsuarioAccion: state.usuario.NombreCompleto, RolUsuario: state.usuario.Rol,
      Unidad: state.usuario.Unidad, FechaAccion: new Date().toISOString(), Observaciones: obs
    }).catch(e => console.warn("Historial (no crítico):", e.message));
    showToast("success", `🔄 Estado cambiado a "${nuevoEstado}"`);
    await renderDirector();
  } catch(e) { showToast("error", "Error: " + e.message); }
  finally { hideLoading(); }
}

// ===== UNIDAD VIEW =====
// Orden de estados para mostrar más urgentes primero
const ORDEN_ESTADO = { "Devuelta": 0, "Derivada": 1, "En Proceso": 2, "Respondida": 3, "Pendiente de Cierre": 5, "Cerrada": 6, "Ingresada": 4 };

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
    state.solicitudes = all.filter(s => (s.UnidadDerivada||"").trim() === (state.usuario.Unidad||"").trim());
    // Orden: Devuelta > Derivada > En Proceso > Respondida > Cerrada
    state.solicitudes = ordenarSolicitudes(state.solicitudes);
    renderSidebarUnidad();
    renderDetalleUnidad(null);
    updateTabBadges();
    verificarVencimientosUnidad();
  } catch (e) {
    showToast("error", "Error: " + e.message);
  } finally {
    hideLoading();
  }
}

function renderSidebarUnidad() {
  const counts = {};
  state.solicitudes.forEach(s => { counts[s.Estado] = (counts[s.Estado] || 0) + 1; });

  const estados = ["Todos","Derivada","En Proceso","Respondida","Devuelta","Pendiente de Cierre","Cerrada"];
  const q = (state.filtroBuscar || "").toLowerCase().trim();
  const filtradas = state.solicitudes.filter(s => {
    if (state.filtroEstado !== "Todos" && s.Estado !== state.filtroEstado) return false;
    if (q) return (s.NroSolicitud||"").toLowerCase().includes(q) ||
                  (s.Solicitante||"").toLowerCase().includes(q) ||
                  (s.Direccion||"").toLowerCase().includes(q) ||
                  (s.Solicitud||"").toLowerCase().includes(q);
    return true;
  });
  const inicio = (state.pagina - 1) * state.pageSize;
  const pagina = filtradas.slice(inicio, inicio + state.pageSize);
  const totalPags = Math.ceil(filtradas.length / state.pageSize);

  const uniStats = [
    { e:"Derivada",            icon:"📤", bg:"#fef3c7", tc:"#b45309" },
    { e:"En Proceso",          icon:"⚙️", bg:"#cffafe", tc:"#0e7490" },
    { e:"Respondida",          icon:"✅", bg:"#dcfce7", tc:"#15803d" },
    { e:"Devuelta",            icon:"↩️", bg:"#fee2e2", tc:"#b91c1c" },
    { e:"Pendiente de Cierre", icon:"⏳", bg:"#fdf4ff", tc:"#7e22ce" },
    { e:"Cerrada",             icon:"🔒", bg:"#f3f4f6", tc:"#4b5563" }
  ];

  const cont = document.getElementById("uni-lista");
  cont.innerHTML = `
    <div style="padding:8px 10px 0;">
      <!-- Botón Todos -->
      <button onclick="filtrarUnidad('Todos')"
        style="width:100%;margin-bottom:6px;padding:6px 10px;border-radius:7px;border:1.5px solid ${state.filtroEstado==='Todos'?'var(--azul)':'var(--borde)'};
               background:${state.filtroEstado==='Todos'?'var(--azul)':'white'};color:${state.filtroEstado==='Todos'?'white':'var(--texto)'};
               font-size:11px;font-weight:600;cursor:pointer;display:flex;justify-content:space-between;align-items:center;">
        <span>📊 Todas</span>
        <span style="background:${state.filtroEstado==='Todos'?'rgba(255,255,255,0.25)':'var(--gris-bg)'};padding:1px 7px;border-radius:10px;font-weight:700;">
          ${state.solicitudes.length}
        </span>
      </button>
      <!-- Grid 2x3 estados -->
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;margin-bottom:8px;">
        ${uniStats.map(({e,icon,bg,tc}) => {
          const cnt = counts[e]||0;
          const activo = state.filtroEstado === e;
          return `
          <div onclick="filtrarUnidad('${e}')"
            style="background:${activo?tc:bg};border-radius:8px;padding:6px 4px;text-align:center;cursor:pointer;
                   border:2px solid ${activo?tc:'transparent'};transition:all 0.18s;
                   box-shadow:${activo?'0 2px 8px rgba(0,0,0,0.15)':'none'};"
            onmouseenter="this.style.transform='translateY(-1px)'"
            onmouseleave="this.style.transform=''">
            <div style="font-size:13px;margin-bottom:1px;">${icon}</div>
            <div style="font-size:17px;font-weight:800;color:${activo?'white':tc};line-height:1;">${cnt}</div>
            <div style="font-size:9px;color:${activo?'rgba(255,255,255,0.85)':tc};margin-top:1px;font-weight:500;">${e}</div>
          </div>`;
        }).join("")}
      </div>
      <input type="text" id="uni-buscar" placeholder="🔍 Buscar..." style="width:100%;padding:9px;border:1.5px solid #dde3ee;border-radius:8px;font-size:13px;box-sizing:border-box;"
        oninput="state.filtroBuscar=this.value;state.pagina=1;renderSidebarUnidad()">
    </div>
    <div style="flex:1;overflow-y:auto;padding:12px;">
      ${pagina.length === 0 ? '<p style="text-align:center;color:#9ca3af;padding:40px">Sin solicitudes</p>' :
        pagina.map(s => `
          ${(() => {
            const sem = calcularSemaforo(s);
            const semBadge = sem
              ? `<span style="background:${sem.bg};color:${sem.color};border:1px solid ${sem.color}40;padding:1px 7px;border-radius:10px;font-size:10px;font-weight:700;white-space:nowrap;">${sem.emoji} ${sem.texto}</span>`
              : "";
            return `
          <div class="sol-card ${state.solicitudSeleccionada?.id === s.id ? 'selected' : ''}" onclick="seleccionarSolicitudUnidad('${s.id}')">
            <div class="sol-card-top">
              <span class="sol-nro">${s.NroSolicitud}</span>
              <span class="estado-badge estado-${s.Estado}">${s.Estado}</span>
            </div>
            <div class="sol-card-name">${s.Solicitante}</div>
            <div style="display:flex;align-items:center;justify-content:space-between;gap:4px;margin-top:2px;">
              <div class="sol-card-dir" style="margin-top:0;">📍 ${s.Direccion || ""}</div>
              ${semBadge}
            </div>
          </div>`;
          })()}`).join("")}
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

function verificarVencimientosUnidad() {
  const umbral = CONFIG.plazoAlertaDias ?? 1;
  const criticas = state.solicitudes.filter(s => {
    const sem = calcularSemaforo(s);
    return sem && sem.dias <= umbral;
  });
  if (!criticas.length) return;
  const hoy    = criticas.filter(s => calcularSemaforo(s).dias === 0);
  const manana = criticas.filter(s => calcularSemaforo(s).dias === 1);
  const venc   = criticas.filter(s => calcularSemaforo(s).dias < 0);
  const partes = [];
  if (venc.length)   partes.push(`${venc.length} vencida${venc.length>1?'s':''}`);
  if (hoy.length)    partes.push(`${hoy.length} vence hoy`);
  if (manana.length) partes.push(`${manana.length} vence mañana`);
  showToast("error", `⏰ Plazo crítico: ${partes.join(" · ")}`);
}
function cambiarPaginaUni(p) { state.pagina = p; renderSidebarUnidad(); }

async function seleccionarSolicitudUnidad(id) {
  state.solicitudSeleccionada = state.solicitudes.find(s => s.id === id);
  const sol = state.solicitudSeleccionada;
  renderSidebarUnidad();
  renderDetalleUnidad(sol);
  mostrarDetalleMovil('.uni-layout');

  // Cargar PDF en panel central
  const pdfPanel = document.getElementById("uni-pdf");
  const pdfHeader = document.getElementById("uni-pdf-header");
  if (!pdfPanel || !sol) return;
  _uniPdfDoc  = null;
  _uniPdfZoom = 1.0;
  pdfPanel.innerHTML = `<div style="text-align:center;color:#9ca3af;padding:30px;"><div class="spinner" style="margin:0 auto 12px;"></div><p>Cargando documento...</p></div>`;
  try {
    const atts = await getListItemAttachments(CONFIG.lists.solicitudes, sol.id);
    if (!atts.length) {
      pdfPanel.innerHTML = `<div class="pdf-visor-empty"><span>📭</span><p>Sin documentos adjuntos</p></div>`;
      return;
    }
    const first = atts.find(a => a.name?.toLowerCase().endsWith('.pdf')) || atts[0];
    const blobUrl = await getAttachmentBlobUrl(first.downloadUrl, first.serverRelativeUrl);
    const isPdf = first.name?.toLowerCase().endsWith('.pdf') && typeof pdfjsLib !== "undefined";
    if (pdfHeader) {
      pdfHeader.style.display = "flex";
      pdfHeader.innerHTML = `
        📄 <span style="font-weight:700;margin-left:4px;">${sol.NroSolicitud}</span>
        <span style="font-weight:400;font-size:12px;opacity:0.8;margin-left:4px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;">— ${first.name}</span>
        <div style="display:flex;align-items:center;gap:4px;margin-left:8px;flex-shrink:0;">
          ${isPdf ? `
          <button onclick="zoomUniPdf(-0.25)"
            style="background:#1f2937;border:1px solid #374151;color:white;width:28px;height:28px;border-radius:6px;font-size:16px;cursor:pointer;line-height:1;display:flex;align-items:center;justify-content:center;font-weight:700;" title="Alejar">−</button>
          <span id="uni-pdf-zoom-label" style="font-size:11px;min-width:36px;text-align:center;font-weight:600;color:#1f2937;">100%</span>
          <button onclick="zoomUniPdf(0.25)"
            style="background:#1f2937;border:1px solid #374151;color:white;width:28px;height:28px;border-radius:6px;font-size:16px;cursor:pointer;line-height:1;display:flex;align-items:center;justify-content:center;font-weight:700;" title="Acercar">+</button>
          <div style="width:1px;height:20px;background:#d1d5db;margin:0 2px;"></div>` : ""}
          <a href="${blobUrl}" download="${first.name}"
            style="background:#1f2937;border:1px solid #374151;color:white;padding:4px 10px;border-radius:6px;font-size:12px;text-decoration:none;white-space:nowrap;font-weight:600;">
            ⬇ Descargar
          </a>
          <button onclick="(function(){var w=window.open('${blobUrl}','_blank');if(w)setTimeout(function(){w.print();},800);})()"
            style="background:#1f2937;border:1px solid #374151;color:white;padding:4px 10px;border-radius:6px;font-size:12px;cursor:pointer;white-space:nowrap;font-weight:600;">
            🖨 Imprimir
          </button>
        </div>`;
    }
    pdfPanel.innerHTML = "";
    if (isPdf) {
      pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
      const wrap = document.createElement("div");
      wrap.id = "uni-pdf-wrap";
      wrap.style.cssText = "flex:1;overflow-y:auto;background:#525659;border-radius:8px;padding:10px;display:flex;flex-direction:column;align-items:center;gap:6px;min-height:0;";
      pdfPanel.appendChild(wrap);
      _uniPdfDoc = await pdfjsLib.getDocument(blobUrl).promise;
      await _renderUniPdfPages();
    } else {
      const img = document.createElement("img");
      img.src = blobUrl;
      img.style.cssText = "width:100%;border-radius:8px;";
      pdfPanel.appendChild(img);
    }
  } catch(e) {
    pdfPanel.innerHTML = `<div class="pdf-visor-empty"><span>⚠️</span><p style="color:#ef4444;font-size:13px;">No se pudo cargar el documento</p></div>`;
  }
}

async function _renderUniPdfPages() {
  const wrap = document.getElementById("uni-pdf-wrap");
  if (!wrap || !_uniPdfDoc) return;
  wrap.innerHTML = `<div style="text-align:center;color:rgba(255,255,255,0.5);padding:12px;font-size:12px;">Renderizando...</div>`;
  const baseW = (wrap.clientWidth || 350) - 20;
  const frag = document.createDocumentFragment();
  for (let p = 1; p <= _uniPdfDoc.numPages; p++) {
    const page = await _uniPdfDoc.getPage(p);
    const vp = page.getViewport({ scale: 1 });
    const scale = (baseW / vp.width) * _uniPdfZoom;
    const svp = page.getViewport({ scale });
    const canvas = document.createElement("canvas");
    canvas.width = svp.width; canvas.height = svp.height;
    canvas.style.cssText = "max-width:100%;border-radius:4px;box-shadow:0 2px 8px rgba(0,0,0,0.3);display:block;";
    await page.render({ canvasContext: canvas.getContext("2d"), viewport: svp }).promise;
    frag.appendChild(canvas);
    if (p < _uniPdfDoc.numPages) {
      const sep = document.createElement("div");
      sep.style.cssText = "color:rgba(255,255,255,0.35);font-size:11px;text-align:center;padding:2px;";
      sep.textContent = `— ${p} / ${_uniPdfDoc.numPages} —`;
      frag.appendChild(sep);
    }
  }
  wrap.innerHTML = "";
  wrap.appendChild(frag);
}

function zoomUniPdf(delta) {
  _uniPdfZoom = Math.max(0.5, Math.min(3.0, _uniPdfZoom + delta));
  const label = document.getElementById("uni-pdf-zoom-label");
  if (label) label.textContent = Math.round(_uniPdfZoom * 100) + "%";
  _renderUniPdfPages();
}

function imprimirSolicitudUnidad(sol) {
  if (!sol) return;
  const sem = calcularSemaforo(sol);
  const semTexto = sem
    ? (sem.dias > 0
        ? `${sem.emoji} ${sem.dias === 1 ? "Vence mañana" : `${sem.dias} días restantes`} (límite ${CONFIG.plazoDerivacionDias||15} días, vence ${formatFecha(sem.vencimiento.toISOString())})`
        : sem.dias === 0
          ? `${sem.emoji} Vence hoy`
          : `${sem.emoji} Plazo vencido hace ${Math.abs(sem.dias)} día${Math.abs(sem.dias)>1?'s':''}`)
    : "";
  const w = window.open("", "_blank");
  if (!w) { showToast("error","El navegador bloqueó la ventana emergente. Permite ventanas emergentes para este sitio."); return; }
  w.document.write(`<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<title>Solicitud ${sol.NroSolicitud} — DOM Doñihue</title>
<style>
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:'Segoe UI',Arial,sans-serif; font-size:13px; color:#1a202c; background:#fff; padding:28px; }
  .header { display:flex; align-items:center; gap:16px; border-bottom:3px solid #1a3a6b; padding-bottom:14px; margin-bottom:18px; }
  .header h1 { font-size:17px; color:#1a3a6b; font-weight:800; }
  .header p  { font-size:11px; color:#6b7280; margin-top:2px; }
  .badge { display:inline-block; padding:3px 10px; border-radius:20px; font-size:11px; font-weight:700;
    background:#dbeafe; color:#1e40af; border:1px solid #93c5fd; }
  table { width:100%; border-collapse:collapse; margin-bottom:14px; }
  th { background:#f1f5f9; color:#475569; font-size:10px; font-weight:700; letter-spacing:0.5px;
    text-transform:uppercase; padding:7px 10px; text-align:left; border:1px solid #e2e8f0; }
  td { padding:8px 10px; border:1px solid #e2e8f0; vertical-align:top; line-height:1.5; }
  .section-title { font-size:11px; font-weight:800; color:#1a3a6b; text-transform:uppercase;
    letter-spacing:0.5px; margin:16px 0 6px; border-left:3px solid #1a3a6b; padding-left:8px; }
  .descripcion { background:#f8fafc; border:1px solid #e2e8f0; border-radius:6px;
    padding:12px; font-size:13px; line-height:1.7; margin-bottom:14px; white-space:pre-wrap; }
  .semaforo { padding:8px 14px; border-radius:6px; font-size:12px; font-weight:700;
    margin-bottom:14px; border:1px solid; }
  .footer { margin-top:24px; border-top:1px solid #e2e8f0; padding-top:10px;
    font-size:10px; color:#9ca3af; display:flex; justify-content:space-between; }
  @media print {
    body { padding:12px; }
    .no-print { display:none !important; }
  }
</style></head><body>
<div class="header">
  <div>
    <h1>Dirección de Obras — Municipalidad de Doñihue</h1>
    <p>Registro de Solicitud · Sistema DOM</p>
  </div>
</div>
<button class="no-print" onclick="window.print()"
  style="margin-bottom:16px;padding:8px 18px;background:#1a3a6b;color:white;border:none;border-radius:6px;font-size:13px;cursor:pointer;font-weight:600;">
  🖨 Imprimir / Guardar PDF
</button>
${sem ? `<div class="semaforo" style="background:${sem.bg};color:${sem.color};border-color:${sem.color}40;">${semTexto}</div>` : ""}
<div class="section-title">Datos de la Solicitud</div>
<table>
  <tr><th>Nro Solicitud</th><td><strong>${sol.NroSolicitud||"-"}</strong></td><th>Estado</th><td><span class="badge">${sol.Estado||"-"}</span></td></tr>
  <tr><th>Fecha Recepción</th><td>${formatFecha(sol.FechaRecepcion)}</td><th>Unidad</th><td>${sol.UnidadDerivada||"-"}</td></tr>
  ${sol.FechaDerivacion?`<tr><th>Fecha Derivación</th><td colspan="3">${formatFecha(sol.FechaDerivacion)}</td></tr>`:""}
</table>
<div class="section-title">Solicitante</div>
<table>
  <tr><th>Nombre</th><td>${sol.Solicitante||"-"}</td><th>RUT</th><td>${sol.Rut||"-"}</td></tr>
  <tr><th>Dirección</th><td colspan="3">${sol.Direccion||"-"}</td></tr>
  ${sol.Correo?`<tr><th>Correo</th><td colspan="3">${sol.Correo}</td></tr>`:""}
  ${sol.Telefono?`<tr><th>Teléfono</th><td colspan="3">${sol.Telefono}</td></tr>`:""}
</table>
<div class="section-title">Descripción</div>
<div class="descripcion">${sol.Solicitud||"Sin descripción registrada"}</div>
${sol.Acciones?`<div class="section-title">Instrucciones del Director</div><div class="descripcion">${sol.Acciones}</div>`:""}
<div style="margin-top:20px;border:1px solid #e2e8f0;border-radius:6px;padding:14px;">
  <div style="font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;margin-bottom:8px;">Observaciones / Respuesta de la Unidad</div>
  <div style="height:80px;border-bottom:1px dashed #d1d5db;"></div>
  <div style="height:80px;border-bottom:1px dashed #d1d5db;"></div>
  <div style="height:80px;"></div>
</div>
<div class="footer">
  <span>DOM · Municipalidad de Doñihue</span>
  <span>Impreso: ${new Date().toLocaleString("es-CL")}</span>
</div>
</body></html>`);
  w.document.close();
  w.focus();
}

async function renderDetalleUnidad(sol) {
  const cont = document.getElementById("uni-detalle");
  if (!sol) {
    cont.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#9ca3af;flex-direction:column;gap:16px;"><span style="font-size:60px">📋</span><p>Selecciona una solicitud</p></div>`;
    return;
  }

  const esCerrada = sol.Estado === CONFIG.estados.CERRADA;
  const esPendienteCierre = sol.Estado === CONFIG.estados.PENDIENTE_CIERRE;
  const esRespondida = sol.Estado === CONFIG.estados.RESPONDIDA;
  const puedeResponder = !esCerrada;

  const [solicitudAtts, evidencias] = await Promise.all([
    getListItemAttachments(CONFIG.lists.solicitudes, sol.id).catch(() => []),
    getEvidenciasBySolicitud(sol.NroSolicitud, sol.id).catch(() => [])
  ]);

  // Instrucciones del director desde campo Acciones de la solicitud (guardado al derivar)
  const instruccionDirector = sol.Acciones || "";
  // Plazo desde FechaCierre (campo de Solicitud_Dom, guardado al marcar Pendiente de Cierre)
  const placoBruto = sol.FechaCierre || null;
  const plazoCierreTexto = placoBruto ? formatFecha(placoBruto) : null;

  const sem = calcularSemaforo(sol);
  const semBanner = sem ? `
    <div style="background:${sem.bg};border-bottom:2px solid ${sem.color};padding:8px 14px;display:flex;align-items:center;gap:10px;">
      <span style="font-size:20px;">${sem.emoji}</span>
      <div style="flex:1;">
        <div style="font-size:12px;font-weight:700;color:${sem.color};">
          ${sem.dias > 0
            ? `Plazo: ${sem.dias === 1 ? "Vence mañana" : `${sem.dias} días restantes`}`
            : sem.dias === 0 ? "⚠️ Vence hoy" : `⚠️ Plazo vencido hace ${Math.abs(sem.dias)} día${Math.abs(sem.dias)>1?'s':''}`}
        </div>
        <div style="font-size:11px;color:${sem.color};opacity:0.8;">Límite ${CONFIG.plazoDerivacionDias||15} días · Vence el ${formatFecha(sem.vencimiento.toISOString())}</div>
      </div>
    </div>` : "";

  cont.innerHTML = `
    <div style="overflow-y:auto;height:100%;display:flex;flex-direction:column;gap:0;">
      <button class="mobile-back-bar" onclick="volverAListaMovil('.uni-layout')">← Volver a lista</button>
      <div class="panel-header" style="background:#f8fafc;border-bottom:1px solid var(--borde);display:flex;align-items:center;gap:8px;">
        <span style="flex:1;">📋 ${sol.NroSolicitud} — ${sol.Solicitante}</span>
        <button onclick="imprimirSolicitudUnidad(state.solicitudSeleccionada)"
          style="background:#1a3a6b;color:white;border:none;border-radius:6px;padding:4px 10px;font-size:11px;font-weight:600;cursor:pointer;white-space:nowrap;flex-shrink:0;"
          title="Imprimir ficha de esta solicitud">🖨 Imprimir</button>
      </div>
      ${semBanner}

      <!-- Datos generales -->
      <div style="padding:12px;background:#f8fafc;border-bottom:1px solid var(--borde);">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:13px;">
          <div><span style="color:#666">Fecha: </span>${formatFecha(sol.FechaRecepcion)}</div>
          <div><span style="color:#666">Estado: </span><span class="estado-badge estado-${sol.Estado}">${sol.Estado}</span></div>
          <div style="grid-column:1/-1;"><span style="color:#666">Dirección: </span>${sol.Direccion||"-"}</div>
          <div style="grid-column:1/-1;"><span style="color:#666">Solicitud: </span>${sol.Solicitud||"-"}</div>
          ${sol.ObservacionesDirector ? `<div style="grid-column:1/-1;"><span style="color:#666">Obs. Director: </span>${sol.ObservacionesDirector}</div>` : ""}
        </div>
      </div>

      <!-- Banner Pendiente de Cierre -->
      ${esPendienteCierre ? `
      <div style="background:#fdf4ff;border-left:4px solid #7e22ce;padding:12px 14px;border-bottom:1px solid #e9d5ff;">
        <div style="font-size:12px;font-weight:700;color:#7e22ce;margin-bottom:6px;">⏳ PENDIENTE DE CIERRE — Segunda instancia</div>
        ${plazoCierreTexto ? `<div style="font-size:12px;color:#6b21a8;margin-bottom:4px;">📅 Plazo fijado por Director: <strong>${plazoCierreTexto}</strong></div>` : ""}
        ${instruccionDirector ? `
        <div style="background:#ede9fe;border-left:3px solid #7c3aed;padding:8px 10px;border-radius:0 6px 6px 0;font-size:13px;color:#4c1d95;margin-top:4px;">
          🏛️ <strong>Instrucción del Director:</strong><br>${instruccionDirector}
        </div>` : ""}
        <div style="font-size:11px;color:#7e22ce;margin-top:8px;">Registra una nueva respuesta para cerrar esta solicitud.</div>
      </div>` : ""}

      <!-- Instrucciones del director (otros estados) -->
      ${!esPendienteCierre && instruccionDirector ? `
      <div style="background:#eff6ff;border-left:4px solid #1a3a6b;padding:10px 14px;border-bottom:1px solid #bfdbfe;">
        <div style="font-size:11px;font-weight:700;color:#1a3a6b;margin-bottom:4px;">🏛️ Instrucciones del Director</div>
        <div style="font-size:13px;color:#1e3a5f;">${instruccionDirector}</div>
      </div>` : ""}

      <!-- Documentos de la solicitud -->
      <div id="uni-sol-adjuntos" style="padding:12px;border-bottom:1px solid var(--borde);">
        <div style="font-size:11px;font-weight:700;color:#6b7280;margin-bottom:6px;">📎 Documentos de la solicitud</div>
        ${solicitudAtts.length === 0
          ? `<p style="color:#9ca3af;font-size:13px;text-align:center;padding:4px;">Sin documentos adjuntos</p>`
          : solicitudAtts.map(a => {
              const isPdf = a.name?.toLowerCase().endsWith('.pdf');
              const isImg = /\.(jpg|jpeg|png|gif)$/i.test(a.name||'');
              return `<div class="file-item" style="cursor:pointer;border-left:3px solid ${isPdf?'#ef4444':isImg?'#3b82f6':'#6b7280'};"
                onclick="abrirAdjuntoUnidad('${a.downloadUrl}','${a.serverRelativeUrl}','${a.name}',${isPdf})">
                <span>${isPdf?'📄':isImg?'🖼️':'📎'} <strong>${a.name}</strong></span>
                <span style="font-size:11px;color:var(--azul);">Ver ↗</span>
              </div>`;
            }).join('')
        }
      </div>

      <!-- Respuestas anteriores -->
      ${evidencias.length > 0 ? `
      <div style="border-top:2px solid #15803d;">
        <div style="background:linear-gradient(90deg,#14532d,#15803d);color:white;padding:8px 14px;font-size:12px;font-weight:700;letter-spacing:0.3px;">
          🏢 ${evidencias.length > 1 ? `Respuestas registradas (${evidencias.length})` : "Respuesta registrada"}
        </div>
        ${evidencias.map((e,idx)=>`
          <div style="padding:12px 14px;border-bottom:1px solid #e2e8f0;${idx>0?'background:#f9fafb;':''}">
            <div style="font-size:11px;font-weight:700;color:#6b7280;margin-bottom:2px;">Respuesta ${idx+1}</div>
            <div style="font-size:12px;font-weight:700;color:#15803d;margin-bottom:6px;">
              👤 ${e.Responsable||""} &nbsp;·&nbsp; 📅 ${formatFecha(e.FechaCarga)}
            </div>
            <div style="font-size:13px;color:#374151;line-height:1.7;background:#f0fdf4;border-left:3px solid #15803d;padding:10px 12px;border-radius:0 8px 8px 0;margin-bottom:8px;">
              ${e.DescripcionEvidencia||"Sin descripción."}
            </div>
            <div id="uni-ev-media-${e.id}" style="margin-top:4px;"></div>
          </div>`).join("")}
      </div>` : ""}

      <!-- Panel de acción -->
      ${puedeResponder ? `
      <div style="padding:14px;display:flex;flex-direction:column;gap:10px;">
        <div class="form-section">
          <div class="form-section-header" style="font-size:11px;">
            ${esPendienteCierre ? "📝 Nueva respuesta (2ª instancia)" : "📝 Descripción de la Solución"}
          </div>
          <div class="form-section-body">
            <textarea id="uni-obs" rows="4"
              placeholder="${esPendienteCierre ? "Describe las acciones tomadas en la segunda inspección..." : "Describe detalladamente la acción realizada, visita, notificación o solución ejecutada..."}"
              style="width:100%;padding:10px;border:1.5px solid ${esPendienteCierre?'#a855f7':'var(--borde)'};border-radius:8px;font-size:13px;font-family:inherit;resize:vertical;"></textarea>
          </div>
        </div>
        <div class="form-section">
          <div class="form-section-header naranja" style="font-size:11px;">📸 Evidencia Fotográfica / Documentos</div>
          <div class="form-section-body">
            <p style="font-size:12px;color:#888;margin-bottom:8px;">Adjunta fotos de la visita o documentos que corroboren la solución ejecutada.</p>
            <div style="border:2px dashed ${esPendienteCierre?'#a855f7':'#f59e0b'};border-radius:10px;padding:16px;text-align:center;cursor:pointer;background:${esPendienteCierre?'#fdf4ff':'#fffbeb'};"
              onclick="document.getElementById('uni-ev-files').click()"
              ondragover="event.preventDefault();this.style.opacity='0.7'"
              ondragleave="this.style.opacity='1'"
              ondrop="event.preventDefault();this.style.opacity='1';handleEvFiles(event.dataTransfer.files)">
              <div style="font-size:28px;">📷</div>
              <div style="font-weight:600;font-size:13px;color:${esPendienteCierre?'#7e22ce':'#b45309'};">Arrastra fotos aquí</div>
              <div style="font-size:11px;color:#9ca3af;margin-top:2px;">o haz clic — JPG, PNG, PDF</div>
            </div>
            <input type="file" id="uni-ev-files" multiple accept=".pdf,.jpg,.jpeg,.png" onchange="handleEvFiles(this.files)" style="display:none;">
            <div id="uni-ev-preview" style="margin-top:8px;display:grid;grid-template-columns:repeat(3,1fr);gap:6px;"></div>
          </div>
        </div>
        <button class="btn-success" onclick="responderSolicitud('${sol.id}')"
          style="width:100%;padding:13px;font-size:14px;${esPendienteCierre?'background:linear-gradient(90deg,#7e22ce,#9333ea);':''}">
          ${esPendienteCierre ? "📋 Registrar 2ª Respuesta" : "✅ Responder — Registrar Solución"}
        </button>
        ${!esPendienteCierre && !esRespondida ? `
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
          <button class="btn-primary" style="background:#0e7490;padding:10px;" onclick="enProcesoSolicitud('${sol.id}')">⚙️ En Proceso</button>
          <button class="btn-warning" style="padding:10px;" onclick="devolverSolicitudUnidad('${sol.id}')">↩️ Devolver</button>
        </div>` : ""}
      </div>` : `
      <div style="padding:14px;">
        <div style="background:#f3f4f6;border-radius:8px;padding:14px;text-align:center;color:#6b7280;font-size:13px;">
          🔒 Solicitud cerrada — solo lectura
        </div>
      </div>`}

      <!-- Hilo de interacciones -->
      <div style="border-top:2px solid #e2e8f0;">
        <div style="background:linear-gradient(90deg,#1e3a5f,#1a3a6b);color:white;padding:8px 14px;font-size:12px;font-weight:700;letter-spacing:0.3px;">
          💬 Historial de Interacciones
        </div>
        <div id="uni-hilo" style="padding:8px 12px;max-height:320px;overflow-y:auto;">
          <div style="text-align:center;color:#9ca3af;font-size:12px;padding:16px;">
            <div class="spinner" style="margin:0 auto 8px;width:20px;height:20px;border-width:2px;"></div>Cargando...
          </div>
        </div>
      </div>
    </div>`;

  // ── Hilo de interacciones (historial + evidencias) ──
  cargarHistorialDir(sol.NroSolicitud, "uni-hilo");

  // ── Cargar imágenes adjuntas de cada evidencia registrada ──
  for (const ev of evidencias) {
    const mediaCont = document.getElementById(`uni-ev-media-${ev.id}`);
    if (!mediaCont) continue;
    getListItemAttachments(CONFIG.lists.evidencias, ev.id).then(async atts => {
      const imgs = atts.filter(a => /\.(jpg|jpeg|png|gif)$/i.test(a.name||''));
      const pdfs = atts.filter(a => a.name?.toLowerCase().endsWith('.pdf'));
      if (!imgs.length && !pdfs.length) return; // sin adjuntos: no mostrar nada
      mediaCont.innerHTML = `<div style="font-size:11px;font-weight:700;color:#6b7280;margin-bottom:6px;">📎 Adjuntos (${atts.length})</div>`;

      // Grilla de imágenes
      if (imgs.length) {
        const grid = document.createElement("div");
        grid.style.cssText = `display:grid;grid-template-columns:repeat(${Math.min(imgs.length,2)},1fr);gap:6px;margin-bottom:8px;`;
        mediaCont.appendChild(grid);
        for (const img of imgs) {
          const box = document.createElement("div");
          box.style.cssText = "border-radius:8px;overflow:hidden;aspect-ratio:4/3;background:#f3f4f6;border:2px solid var(--borde);cursor:zoom-in;position:relative;";
          box.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;height:100%;"><div class="spinner" style="width:20px;height:20px;border-width:2px;"></div></div>`;
          grid.appendChild(box);
          getAttachmentBlobUrl(img.downloadUrl, img.serverRelativeUrl).then(blobUrl => {
            const imgEl = document.createElement("img");
            imgEl.src = blobUrl;
            imgEl.style.cssText = "width:100%;height:100%;object-fit:cover;";
            imgEl.onclick = () => abrirLightbox(blobUrl, img.name);
            box.innerHTML = "";
            box.appendChild(imgEl);
            const badge = document.createElement("div");
            badge.style.cssText = "position:absolute;bottom:4px;right:4px;background:rgba(0,0,0,0.5);color:white;border-radius:4px;padding:2px 6px;font-size:10px;";
            badge.textContent = "🔍 ver";
            box.appendChild(badge);
          }).catch(() => { box.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;height:100%;font-size:11px;color:#fca5a5;">⚠️ Error</div>`; });
        }
      }

      // PDFs listados
      for (const pdf of pdfs) {
        const pdfRow = document.createElement("div");
        pdfRow.style.cssText = "font-size:12px;padding:6px 8px;background:#f8fafc;border:1px solid var(--borde);border-radius:6px;margin-bottom:4px;cursor:pointer;color:var(--azul);font-weight:600;";
        pdfRow.innerHTML = `📄 ${pdf.name}`;
        pdfRow.onclick = async () => {
          try {
            const blobUrl = await getAttachmentBlobUrl(pdf.downloadUrl, pdf.serverRelativeUrl);
            window.open(blobUrl, "_blank");
          } catch {}
        };
        mediaCont.appendChild(pdfRow);
      }
    }).catch(() => {});
  }
}

// Archivos de evidencia de la unidad
const _evFiles = [];
function handleEvFiles(files) {
  Array.from(files).forEach(f => _evFiles.push(f));
  const preview = document.getElementById("uni-ev-preview");
  if (!preview) return;
  preview.innerHTML = "";
  _evFiles.forEach((f, i) => {
    const box = document.createElement("div");
    box.style.cssText = "position:relative;aspect-ratio:1;border-radius:8px;overflow:hidden;background:#f3f4f6;border:1.5px solid var(--borde);";
    if (/\.(jpg|jpeg|png|gif)$/i.test(f.name)) {
      const url = URL.createObjectURL(f);
      box.innerHTML = `<img src="${url}" style="width:100%;height:100%;object-fit:cover;">`;
    } else {
      box.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;height:100%;font-size:11px;color:#666;padding:4px;text-align:center;">📄 ${f.name}</div>`;
    }
    const del = document.createElement("button");
    del.style.cssText = "position:absolute;top:2px;right:2px;background:rgba(0,0,0,0.5);color:white;border:none;border-radius:50%;width:18px;height:18px;font-size:10px;cursor:pointer;display:flex;align-items:center;justify-content:center;";
    del.textContent = "✕";
    del.onclick = () => { _evFiles.splice(i,1); handleEvFiles([]); };
    box.appendChild(del);
    preview.appendChild(box);
  });
}

async function responderSolicitud(solId) {
  const obs = document.getElementById("uni-obs")?.value.trim();
  const files = _evFiles.length > 0 ? _evFiles : Array.from(document.getElementById("uni-ev-files")?.files||[]);
  if (!obs) { showToast("error", "Ingresa una observación"); return; }

  showLoading("Respondiendo...");
  try {
    const sol = state.solicitudes.find(s => s.id === solId);
    const esPendienteCierre = sol.Estado === CONFIG.estados.PENDIENTE_CIERRE;
    // Si está en Pendiente de Cierre, se mantiene ese estado (el Director cierra)
    // Si es cualquier otro estado activo, pasa a Respondida
    if (!esPendienteCierre) {
      await actualizarSolicitud(solId, { Estado: CONFIG.estados.RESPONDIDA });
    }

    const evItem = await crearEvidencia({
      Title: sol.NroSolicitud || `Evidencia ${sol.id}`,
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

    const nuevoEstado = esPendienteCierre ? CONFIG.estados.PENDIENTE_CIERRE : CONFIG.estados.RESPONDIDA;
    registrarHistorial({
      NroSolicitud: sol.NroSolicitud,
      Title: esPendienteCierre ? "2ª respuesta registrada (Pend. Cierre)" : "Solicitud respondida",
      EstadoAnterior: sol.Estado,
      EstadoNuevo: nuevoEstado,
      UsuarioAccion: state.usuario.NombreCompleto,
      RolUsuario: state.usuario.Rol,
      Unidad: state.usuario.Unidad,
      FechaAccion: new Date().toISOString(),
      Observaciones: obs
    }).catch(e => console.warn("Historial (no crítico):", e.message));

    const accionNotif = esPendienteCierre ? "2ª respuesta registrada — Pendiente de Cierre" : "Solicitud respondida por unidad";
    notificarDirector({ ...sol, Estado: nuevoEstado }, accionNotif);
    showToast("success", esPendienteCierre ? "📋 2ª respuesta registrada" : "✅ Solicitud respondida");
    _evFiles.splice(0);
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
    registrarHistorial({
      NroSolicitud: sol.NroSolicitud,
      Title:"Solicitud en proceso",
      EstadoAnterior: sol.Estado,
      EstadoNuevo: CONFIG.estados.EN_PROCESO,
      UsuarioAccion: state.usuario.NombreCompleto,
      RolUsuario: state.usuario.Rol,
      Unidad: state.usuario.Unidad,
      FechaAccion: new Date().toISOString(),
      Observaciones: obs
    }).catch(e => console.warn("Historial (no crítico):", e.message));
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
    registrarHistorial({
      NroSolicitud: sol.NroSolicitud,
      Title:"Solicitud devuelta",
      EstadoAnterior: sol.Estado,
      EstadoNuevo: CONFIG.estados.DEVUELTA,
      UsuarioAccion: state.usuario.NombreCompleto,
      RolUsuario: state.usuario.Rol,
      Unidad: state.usuario.Unidad,
      FechaAccion: new Date().toISOString(),
      Observaciones: obs,
      Motivo: obs
    }).catch(e => console.warn("Historial (no crítico):", e.message));
    notificarDirector({ ...sol, Estado: CONFIG.estados.DEVUELTA }, "Solicitud devuelta por unidad");
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
    registrarHistorial({
      NroSolicitud: sol.NroSolicitud,
      Title:"Solicitud cerrada",
      EstadoAnterior: sol.Estado,
      EstadoNuevo: CONFIG.estados.CERRADA,
      UsuarioAccion: state.usuario.NombreCompleto,
      RolUsuario: state.usuario.Rol,
      Unidad: state.usuario.Unidad,
      FechaAccion: new Date().toISOString(),
      Observaciones: obs
    }).catch(e => console.warn("Historial (no crítico):", e.message));
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

// ===== ADMINISTRADOR =====
const TODOS_ESTADOS = ["Ingresada","Derivada","En Proceso","Respondida","Pendiente de Cierre","Devuelta","Cerrada"];
const ESTADO_COLOR  = { "Ingresada":"#3b82f6","Derivada":"#f59e0b","En Proceso":"#06b6d4","Respondida":"#22c55e","Pendiente de Cierre":"#7e22ce","Devuelta":"#ef4444","Cerrada":"#6b7280" };

function admTabStyle(activo) {
  return activo
    ? "padding:8px 18px;background:#312e81;color:white;border:none;border-bottom:3px solid #818cf8;cursor:pointer;font-size:13px;font-weight:700;"
    : "padding:8px 18px;background:white;color:#374151;border:none;border-bottom:3px solid transparent;cursor:pointer;font-size:13px;font-weight:600;";
}

async function renderAdmin(seccion = "solicitudes") {
  const cont = document.getElementById("view-admin");
  cont.innerHTML = `
  <div style="display:flex;flex-direction:column;height:calc(100vh - 120px);overflow:hidden;">
    <!-- Header -->
    <div style="background:linear-gradient(90deg,#1e1b4b,#312e81);color:white;padding:10px 20px;display:flex;align-items:center;gap:12px;flex-shrink:0;">
      <span style="font-size:18px;">🛡️</span>
      <div style="flex:1;">
        <div style="font-size:14px;font-weight:700;">Panel de Administración</div>
        <div style="font-size:11px;opacity:0.7;">Gestión avanzada del sistema DOM</div>
      </div>
    </div>
    <!-- Pestañas internas -->
    <div style="display:flex;background:white;border-bottom:2px solid var(--borde);flex-shrink:0;">
      <button id="adm-tab-solicitudes" style="${admTabStyle(seccion==='solicitudes')}" onclick="renderAdmin('solicitudes')">📋 Solicitudes</button>
      <button id="adm-tab-unidades"    style="${admTabStyle(seccion==='unidades')}"    onclick="renderAdmin('unidades')">🏢 Unidades</button>
      <button id="adm-tab-usuarios"    style="${admTabStyle(seccion==='usuarios')}"    onclick="renderAdmin('usuarios')">👤 Usuarios</button>
    </div>
    <!-- Contenido dinámico -->
    <div id="adm-contenido" style="flex:1;overflow:hidden;display:flex;flex-direction:column;"></div>
  </div>`;

  if (seccion === "solicitudes") await renderAdmSolicitudes();
  else if (seccion === "unidades") await renderAdmUnidades();
  else if (seccion === "usuarios") await renderAdmUsuarios();
}

// ── Sección Solicitudes ──────────────────────────────────
let _adminSolicitudes = [];

async function renderAdmSolicitudes() {
  const cont = document.getElementById("adm-contenido");
  cont.innerHTML = `
    <div style="background:white;border-bottom:1px solid var(--borde);padding:8px 14px;display:flex;flex-wrap:wrap;gap:8px;align-items:center;flex-shrink:0;">
      <input type="text" id="adm-buscar" placeholder="🔍 N°, solicitante, dirección..."
        style="flex:1;min-width:180px;padding:6px 10px;border:1.5px solid var(--borde);border-radius:7px;font-size:12px;"
        oninput="filtrarAdmin()">
      <select id="adm-filtro-estado" onchange="filtrarAdmin()"
        style="padding:6px 8px;border:1.5px solid var(--borde);border-radius:7px;font-size:12px;">
        <option value="">Todos los estados</option>
        ${TODOS_ESTADOS.map(e=>`<option value="${e}">${e}</option>`).join("")}
      </select>
      <select id="adm-filtro-unidad" onchange="filtrarAdmin()"
        style="padding:6px 8px;border:1.5px solid var(--borde);border-radius:7px;font-size:12px;">
        <option value="">Todas las unidades</option>
        ${CONFIG.unidades.map(u=>`<option value="${u}">${u}</option>`).join("")}
      </select>
      <button onclick="renderAdmin('solicitudes')" class="btn-primary" style="padding:6px 12px;font-size:12px;">🔄</button>
      <span id="adm-count" style="font-size:11px;color:#6b7280;font-weight:600;"></span>
    </div>
    <div style="flex:1;overflow-y:auto;">
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead style="position:sticky;top:0;z-index:1;">
          <tr style="background:#1e1b4b;color:white;">
            <th style="padding:9px 10px;text-align:left;">N° Solicitud</th>
            <th style="padding:9px 10px;text-align:left;">Solicitante</th>
            <th style="padding:9px 10px;text-align:left;">Dirección</th>
            <th style="padding:9px 10px;text-align:left;">Fecha</th>
            <th style="padding:9px 10px;text-align:center;">Estado actual</th>
            <th style="padding:9px 10px;text-align:left;">Unidad</th>
            <th style="padding:9px 10px;text-align:center;">Cambiar estado</th>
          </tr>
        </thead>
        <tbody id="adm-tbody">
          <tr><td colspan="7" style="text-align:center;padding:40px;color:#9ca3af;">Cargando...</td></tr>
        </tbody>
      </table>
    </div>`;

  showLoading("Cargando solicitudes...");
  try {
    _adminSolicitudes = await getSolicitudes();
    _adminSolicitudes.sort((a,b) => new Date(b.FechaRecepcion)-new Date(a.FechaRecepcion));
    filtrarAdmin();
  } catch(e) { showToast("error","Error: "+e.message); }
  finally { hideLoading(); }
}

function filtrarAdmin() {
  const q      = (document.getElementById("adm-buscar")?.value||"").toLowerCase();
  const estado = document.getElementById("adm-filtro-estado")?.value||"";
  const unidad = document.getElementById("adm-filtro-unidad")?.value||"";
  const filtradas = _adminSolicitudes.filter(s => {
    if (estado && s.Estado !== estado) return false;
    if (unidad && (s.UnidadDerivada||"").trim() !== unidad) return false;
    if (q) return (s.NroSolicitud||"").toLowerCase().includes(q) ||
                  (s.Solicitante||"").toLowerCase().includes(q) ||
                  (s.Direccion||"").toLowerCase().includes(q);
    return true;
  });
  const cnt = document.getElementById("adm-count");
  if (cnt) cnt.textContent = `${filtradas.length} solicitudes`;
  const tbody = document.getElementById("adm-tbody");
  if (!tbody) return;
  if (!filtradas.length) { tbody.innerHTML=`<tr><td colspan="7" style="text-align:center;padding:28px;color:#9ca3af;">Sin resultados</td></tr>`; return; }
  tbody.innerHTML = filtradas.map(s => {
    const c = ESTADO_COLOR[s.Estado]||"#94a3b8";
    return `<tr style="border-bottom:1px solid var(--borde);" onmouseenter="this.style.background='#f8fafc'" onmouseleave="this.style.background=''">
      <td style="padding:7px 10px;font-weight:700;color:#1a3a6b;">${s.NroSolicitud||`#${s.id}`}</td>
      <td style="padding:7px 10px;">${s.Solicitante||""}</td>
      <td style="padding:7px 10px;color:#6b7280;">${s.Direccion||""}</td>
      <td style="padding:7px 10px;color:#6b7280;white-space:nowrap;">${formatFecha(s.FechaRecepcion)}</td>
      <td style="padding:7px 10px;text-align:center;"><span style="background:${c}20;color:${c};padding:2px 9px;border-radius:20px;font-size:11px;font-weight:700;border:1px solid ${c}40;">${s.Estado}</span></td>
      <td style="padding:7px 10px;font-size:11px;">${s.UnidadDerivada||"—"}</td>
      <td style="padding:7px 10px;text-align:center;">
        <div style="display:flex;gap:5px;align-items:center;justify-content:center;">
          <select id="adm-sel-${s.id}" style="padding:3px 5px;border:1.5px solid var(--borde);border-radius:5px;font-size:11px;">
            <option value="">— nuevo estado —</option>
            ${TODOS_ESTADOS.filter(e=>e!==s.Estado).map(e=>`<option value="${e}">${e}</option>`).join("")}
          </select>
          <button onclick="cambiarEstadoAdmin('${s.id}')" style="padding:4px 9px;background:#312e81;color:white;border:none;border-radius:5px;cursor:pointer;font-size:11px;font-weight:600;">✓</button>
        </div>
      </td>
    </tr>`;
  }).join("");
}

async function cambiarEstadoAdmin(solId) {
  const nuevoEstado = document.getElementById(`adm-sel-${solId}`)?.value;
  if (!nuevoEstado) { showToast("error","Selecciona un estado"); return; }
  const sol = _adminSolicitudes.find(s=>s.id===solId);
  if (!confirm(`¿Cambiar ${sol?.NroSolicitud||solId}?\n${sol?.Estado} → ${nuevoEstado}`)) return;
  showLoading("Actualizando...");
  try {
    await actualizarSolicitud(solId, { Estado: nuevoEstado });
    registrarHistorial({ NroSolicitud:sol.NroSolicitud, Title:`[Admin] Estado → ${nuevoEstado}`,
      EstadoAnterior:sol.Estado, EstadoNuevo:nuevoEstado,
      UsuarioAccion:state.usuario.NombreCompleto, RolUsuario:state.usuario.Rol,
      Unidad:state.usuario.Unidad, FechaAccion:new Date().toISOString(),
      Observaciones:"Cambio manual por Administrador"
    }).catch(e=>console.warn(e));
    showToast("success",`✅ ${sol?.NroSolicitud} → ${nuevoEstado}`);
    sol.Estado = nuevoEstado;
    filtrarAdmin();
  } catch(e) { showToast("error","Error: "+e.message); }
  finally { hideLoading(); }
}

// ── Sección Unidades ──────────────────────────────────
let _adminUnidades = [];

async function renderAdmUnidades() {
  const cont = document.getElementById("adm-contenido");
  cont.innerHTML = `
    <div style="padding:16px 20px;flex-shrink:0;background:white;border-bottom:1px solid var(--borde);display:flex;gap:10px;align-items:flex-end;">
      <div style="flex:1;">
        <label style="font-size:11px;font-weight:700;color:#374151;display:block;margin-bottom:4px;">NOMBRE DE LA UNIDAD</label>
        <input type="text" id="adm-uni-nueva" placeholder="Ej: Inspección, Operaciones..."
          style="width:100%;padding:7px 10px;border:1.5px solid var(--borde);border-radius:7px;font-size:13px;box-sizing:border-box;">
      </div>
      <button onclick="agregarUnidad()" class="btn-primary" style="padding:8px 18px;white-space:nowrap;">+ Agregar Unidad</button>
    </div>
    <div style="flex:1;overflow-y:auto;padding:16px 20px;">
      <div id="adm-uni-lista" style="display:flex;flex-direction:column;gap:8px;">
        <div style="text-align:center;padding:32px;color:#9ca3af;">Cargando unidades...</div>
      </div>
    </div>`;

  await cargarAdmUnidades();
}

async function cargarAdmUnidades() {
  showLoading("Cargando unidades...");
  try {
    _adminUnidades = await getListItems(CONFIG.lists.unidades);
    renderListaUnidades();
  } catch(e) {
    if (e.message.includes("404") || e.message.includes("no existe")) {
      // Lista no existe — mostrar botón para crearla automáticamente
      document.getElementById("adm-uni-lista").innerHTML = `
        <div style="background:#fff7ed;border:1.5px solid #fed7aa;border-radius:10px;padding:20px;text-align:center;">
          <div style="font-size:28px;margin-bottom:8px;">📋</div>
          <div style="font-size:14px;font-weight:700;color:#9a3412;margin-bottom:6px;">La lista "UnidadesDOM" no existe en SharePoint</div>
          <div style="font-size:12px;color:#c2410c;margin-bottom:16px;">Se puede crear automáticamente con las columnas necesarias.</div>
          <button onclick="crearListaUnidades()" class="btn-primary" style="padding:10px 24px;font-size:13px;">
            🚀 Crear lista automáticamente en SharePoint
          </button>
        </div>`;
    } else {
      document.getElementById("adm-uni-lista").innerHTML = `<div style="color:#ef4444;padding:16px;">Error: ${e.message}</div>`;
    }
  } finally { hideLoading(); }
}

async function crearListaUnidades() {
  showLoading("Creando lista UnidadesDOM en SharePoint...");
  try {
    const token = await getSharePointToken();
    const SP_BASE_URL = `${CONFIG.sharePointSite}/_api`;

    // 1. Crear la lista
    const resLista = await fetch(`${SP_BASE_URL}/web/lists`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=verbose"
      },
      body: JSON.stringify({
        __metadata: { type: "SP.List" },
        Title: CONFIG.lists.unidades,
        BaseTemplate: 100,
        Description: "Lista de unidades del sistema DOM"
      })
    });
    if (!resLista.ok) {
      const err = await resLista.text();
      // Si ya existe (código 409), continuamos igual
      if (!err.includes("ya existe") && !err.includes("already exists") && resLista.status !== 409) {
        throw new Error(`No se pudo crear la lista: ${err}`);
      }
    }

    // 2. Agregar columna Activo (Sí/No)
    await fetch(`${SP_BASE_URL}/web/lists/getbytitle('${encodeURIComponent(CONFIG.lists.unidades)}')/fields`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=verbose"
      },
      body: JSON.stringify({
        __metadata: { type: "SP.Field" },
        Title: "Activo",
        FieldTypeKind: 8,  // Boolean
        DefaultValue: "1"
      })
    }).catch(() => {}); // Si ya existe, ignorar

    // 3. Poblar con las unidades actuales de config
    for (const nombre of CONFIG.unidades) {
      await createListItem(CONFIG.lists.unidades, { Title: nombre, Activo: true }).catch(() => {});
    }

    showToast("success", "✅ Lista UnidadesDOM creada con las unidades actuales");
    await cargarAdmUnidades();
  } catch(e) {
    showToast("error", "Error creando lista: " + e.message);
  } finally {
    hideLoading();
  }
}

function renderListaUnidades() {
  const cont = document.getElementById("adm-uni-lista");
  if (!cont) return;
  if (!_adminUnidades.length) {
    cont.innerHTML = `<div style="text-align:center;padding:32px;color:#9ca3af;">No hay unidades. Agrega la primera.</div>`;
    return;
  }
  cont.innerHTML = _adminUnidades.map(u => {
    const activo = u.Activo !== false && u.Activo !== 0;
    return `
    <div style="background:white;border:1.5px solid ${activo?'var(--borde)':'#fecaca'};border-radius:10px;padding:12px 16px;display:flex;align-items:center;gap:12px;">
      <div id="adm-uni-view-${u.id}" style="flex:1;display:flex;align-items:center;gap:10px;">
        <span style="font-size:16px;">🏢</span>
        <span style="font-size:14px;font-weight:600;color:${activo?'#1a3a6b':'#9ca3af'};${activo?'':'text-decoration:line-through;'}">${u.Title}</span>
        <span style="font-size:11px;padding:2px 8px;border-radius:10px;background:${activo?'#dcfce7':'#fee2e2'};color:${activo?'#15803d':'#b91c1c'};font-weight:600;">${activo?'Activa':'Inactiva'}</span>
      </div>
      <div id="adm-uni-edit-${u.id}" style="flex:1;display:none;gap:8px;">
        <input type="text" id="adm-uni-nombre-${u.id}" value="${u.Title}"
          style="flex:1;padding:6px 10px;border:1.5px solid var(--azul);border-radius:6px;font-size:13px;">
      </div>
      <div style="display:flex;gap:6px;">
        <!-- Botón editar -->
        <button id="adm-uni-btn-edit-${u.id}" onclick="editarUnidadToggle('${u.id}')"
          style="padding:5px 10px;background:#f1f5f9;color:#374151;border:1px solid var(--borde);border-radius:6px;cursor:pointer;font-size:12px;">✏️ Editar</button>
        <!-- Botón guardar (oculto) -->
        <button id="adm-uni-btn-save-${u.id}" onclick="guardarUnidad('${u.id}')" style="display:none;
          padding:5px 10px;background:#15803d;color:white;border:none;border-radius:6px;cursor:pointer;font-size:12px;">💾 Guardar</button>
        <!-- Activo/Inactivo -->
        <button onclick="toggleUnidadActivo('${u.id}',${!activo})"
          style="padding:5px 10px;background:${activo?'#fef3c7':'#dcfce7'};color:${activo?'#b45309':'#15803d'};border:1px solid ${activo?'#fde68a':'#86efac'};border-radius:6px;cursor:pointer;font-size:12px;">
          ${activo?'⏸ Desactivar':'▶ Activar'}
        </button>
      </div>
    </div>`;
  }).join("");
}

function editarUnidadToggle(id) {
  const view = document.getElementById(`adm-uni-view-${id}`);
  const edit = document.getElementById(`adm-uni-edit-${id}`);
  const btnEdit = document.getElementById(`adm-uni-btn-edit-${id}`);
  const btnSave = document.getElementById(`adm-uni-btn-save-${id}`);
  const showing = edit.style.display === "flex";
  if (showing) {
    view.style.display = "flex"; edit.style.display = "none";
    btnEdit.style.display = ""; btnSave.style.display = "none";
  } else {
    view.style.display = "none"; edit.style.display = "flex";
    btnEdit.style.display = "none"; btnSave.style.display = "";
  }
}

async function guardarUnidad(id) {
  const nombre = document.getElementById(`adm-uni-nombre-${id}`)?.value.trim();
  if (!nombre) { showToast("error","El nombre no puede estar vacío"); return; }
  showLoading("Guardando...");
  try {
    await actualizarUnidad(id, { Title: nombre });
    const u = _adminUnidades.find(x=>x.id===id);
    if (u) u.Title = nombre;
    CONFIG.unidades = _adminUnidades.filter(x=>x.Activo!==false&&x.Activo!==0).map(x=>x.Title);
    showToast("success","✅ Unidad actualizada");
    renderListaUnidades();
  } catch(e) { showToast("error","Error: "+e.message); }
  finally { hideLoading(); }
}

async function toggleUnidadActivo(id, nuevoActivo) {
  showLoading("Actualizando...");
  try {
    await actualizarUnidad(id, { Activo: nuevoActivo });
    const u = _adminUnidades.find(x=>x.id===id);
    if (u) u.Activo = nuevoActivo;
    CONFIG.unidades = _adminUnidades.filter(x=>x.Activo!==false&&x.Activo!==0).map(x=>x.Title);
    showToast("success", nuevoActivo?"✅ Unidad activada":"⏸ Unidad desactivada");
    renderListaUnidades();
  } catch(e) { showToast("error","Error: "+e.message); }
  finally { hideLoading(); }
}

async function agregarUnidad() {
  const nombre = document.getElementById("adm-uni-nueva")?.value.trim();
  if (!nombre) { showToast("error","Ingresa el nombre de la unidad"); return; }
  showLoading("Creando unidad...");
  try {
    const nueva = await crearUnidad(nombre);
    _adminUnidades.push({ id: String(nueva.Id||nueva.id), Title: nombre, Activo: true });
    CONFIG.unidades = _adminUnidades.filter(u=>u.Activo!==false&&u.Activo!==0).map(u=>u.Title);
    document.getElementById("adm-uni-nueva").value = "";
    showToast("success",`✅ Unidad "${nombre}" creada`);
    renderListaUnidades();
  } catch(e) { showToast("error","Error: "+e.message); }
  finally { hideLoading(); }
}

// ── Sección Usuarios ──────────────────────────────────
let _adminUsuarios = [];

async function renderAdmUsuarios() {
  const cont = document.getElementById("adm-contenido");
  cont.innerHTML = `
    <div style="padding:12px 16px;flex-shrink:0;background:white;border-bottom:1px solid var(--borde);display:flex;gap:8px;align-items:center;">
      <input type="text" id="adm-usr-buscar" placeholder="🔍 Buscar usuario..."
        style="flex:1;padding:6px 10px;border:1.5px solid var(--borde);border-radius:7px;font-size:12px;"
        oninput="filtrarUsuarios()">
      <button onclick="abrirModalUsuario(null)" class="btn-primary" style="padding:7px 14px;white-space:nowrap;">+ Nuevo Usuario</button>
      <button onclick="renderAdmin('usuarios')" style="padding:7px 10px;background:#f1f5f9;color:#374151;border:1px solid var(--borde);border-radius:7px;cursor:pointer;font-size:12px;">🔄</button>
    </div>
    <div style="flex:1;overflow-y:auto;">
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead style="position:sticky;top:0;z-index:1;">
          <tr style="background:#1e1b4b;color:white;">
            <th style="padding:9px 12px;text-align:left;">Nombre</th>
            <th style="padding:9px 12px;text-align:left;">Correo</th>
            <th style="padding:9px 12px;text-align:center;">Rol</th>
            <th style="padding:9px 12px;text-align:left;">Unidad</th>
            <th style="padding:9px 12px;text-align:center;">Activo</th>
            <th style="padding:9px 12px;text-align:center;">Derivar</th>
            <th style="padding:9px 12px;text-align:center;">Cerrar</th>
            <th style="padding:9px 12px;text-align:center;">Admin</th>
            <th style="padding:9px 12px;text-align:center;">Acciones</th>
          </tr>
        </thead>
        <tbody id="adm-usr-tbody">
          <tr><td colspan="6" style="text-align:center;padding:40px;color:#9ca3af;">Cargando...</td></tr>
        </tbody>
      </table>
    </div>
    <!-- Modal usuario -->
    <div id="adm-usr-modal" style="display:none;position:fixed;inset:0;background:rgba(0,0,0,0.45);z-index:1000;align-items:center;justify-content:center;">
      <div style="background:white;border-radius:14px;padding:24px;width:440px;max-width:95vw;box-shadow:0 8px 32px rgba(0,0,0,0.2);">
        <h3 id="adm-usr-modal-title" style="margin:0 0 16px;font-size:15px;color:#1a3a6b;">Nuevo Usuario</h3>
        <div style="display:flex;flex-direction:column;gap:10px;">
          <div>
            <label style="font-size:11px;font-weight:700;color:#374151;display:block;margin-bottom:3px;">NOMBRE COMPLETO</label>
            <input type="text" id="adm-usr-nombre" style="width:100%;padding:7px 10px;border:1.5px solid var(--borde);border-radius:7px;font-size:13px;box-sizing:border-box;">
          </div>
          <div>
            <label style="font-size:11px;font-weight:700;color:#374151;display:block;margin-bottom:3px;">CORREO MICROSOFT 365</label>
            <input type="email" id="adm-usr-correo" style="width:100%;padding:7px 10px;border:1.5px solid var(--borde);border-radius:7px;font-size:13px;box-sizing:border-box;">
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">
            <div>
              <label style="font-size:11px;font-weight:700;color:#374151;display:block;margin-bottom:3px;">ROL</label>
              <select id="adm-usr-rol" style="width:100%;padding:7px 8px;border:1.5px solid var(--borde);border-radius:7px;font-size:13px;" onchange="toggleUnidadField()">
                <option value="Administrador">🛡️ Administrador</option>
                <option value="Director">🏛️ Director</option>
                <option value="Secretaria">📋 Secretaria</option>
                <option value="Unidad">🏢 Unidad</option>
              </select>
            </div>
            <div id="adm-usr-unidad-wrap">
              <label style="font-size:11px;font-weight:700;color:#374151;display:block;margin-bottom:3px;">UNIDAD</label>
              <select id="adm-usr-unidad" style="width:100%;padding:7px 8px;border:1.5px solid var(--borde);border-radius:7px;font-size:13px;">
                <option value="">— Sin unidad —</option>
                ${CONFIG.unidades.map(u=>`<option value="${u}">${u}</option>`).join("")}
              </select>
            </div>
          </div>
          <!-- Checkboxes de permisos -->
          <div style="background:#f8fafc;border-radius:8px;padding:12px;display:flex;flex-direction:column;gap:8px;">
            <div style="font-size:11px;font-weight:700;color:#374151;margin-bottom:2px;">PERMISOS</div>
            <div style="display:flex;flex-wrap:wrap;gap:12px;">
              <label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:13px;">
                <input type="checkbox" id="adm-usr-activo" style="width:15px;height:15px;cursor:pointer;">
                Activo
              </label>
              <label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:13px;">
                <input type="checkbox" id="adm-usr-derivar" style="width:15px;height:15px;cursor:pointer;">
                Puede Derivar
              </label>
              <label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:13px;">
                <input type="checkbox" id="adm-usr-cerrar" style="width:15px;height:15px;cursor:pointer;">
                Puede Cerrar
              </label>
              <label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:13px;">
                <input type="checkbox" id="adm-usr-esadmin" style="width:15px;height:15px;cursor:pointer;">
                Es Administrador
              </label>
            </div>
          </div>
        </div>
        <input type="hidden" id="adm-usr-id">
        <div style="display:flex;gap:10px;margin-top:20px;">
          <button onclick="guardarUsuarioAdmin()" class="btn-primary" style="flex:1;padding:10px;">💾 Guardar</button>
          <button onclick="cerrarModalUsuario()" style="flex:1;padding:10px;background:#f1f5f9;color:#374151;border:1px solid var(--borde);border-radius:8px;cursor:pointer;font-weight:600;">Cancelar</button>
        </div>
      </div>
    </div>`;

  showLoading("Cargando usuarios...");
  try {
    _adminUsuarios = await getTodosUsuarios();
    filtrarUsuarios();
  } catch(e) { showToast("error","Error: "+e.message); }
  finally { hideLoading(); }
}

function filtrarUsuarios() {
  const q = (document.getElementById("adm-usr-buscar")?.value||"").toLowerCase();
  const filtrados = _adminUsuarios.filter(u =>
    !q || (u.NombreCompleto||"").toLowerCase().includes(q) ||
          (u.Correo||"").toLowerCase().includes(q) ||
          (u.Rol||"").toLowerCase().includes(q)
  );
  const tbody = document.getElementById("adm-usr-tbody");
  if (!tbody) return;
  if (!filtrados.length) { tbody.innerHTML=`<tr><td colspan="6" style="text-align:center;padding:28px;color:#9ca3af;">Sin resultados</td></tr>`; return; }
  const ROL_COLOR = { "Administrador":"#312e81","Director":"#1a3a6b","Secretaria":"#0e7490","Unidad":"#15803d" };
  tbody.innerHTML = filtrados.map(u => {
    const c = ROL_COLOR[u.Rol]||"#6b7280";
    const activo = u.Activo !== false && u.Activo !== 0;
    return `<tr style="border-bottom:1px solid var(--borde);" onmouseenter="this.style.background='#f8fafc'" onmouseleave="this.style.background=''">
      <td style="padding:8px 12px;font-weight:600;">${u.NombreCompleto||""}</td>
      <td style="padding:8px 12px;color:#6b7280;font-size:11px;">${u.Correo||""}</td>
      <td style="padding:8px 12px;text-align:center;"><span style="background:${c}20;color:${c};padding:2px 9px;border-radius:20px;font-size:11px;font-weight:700;border:1px solid ${c}30;">${u.Rol||""}</span></td>
      <td style="padding:8px 12px;font-size:11px;">${u.Unidad||"—"}</td>
      <td style="padding:8px 12px;text-align:center;"><span style="padding:2px 9px;border-radius:10px;font-size:11px;font-weight:600;background:${activo?'#dcfce7':'#fee2e2'};color:${activo?'#15803d':'#b91c1c'};">${activo?"Sí":"No"}</span></td>
      <td style="padding:8px 12px;text-align:center;">${u.PuedeDerivar?'✅':'—'}</td>
      <td style="padding:8px 12px;text-align:center;">${u.PuedeCerrar?'✅':'—'}</td>
      <td style="padding:8px 12px;text-align:center;">${u.EsAdministrador?'✅':'—'}</td>
      <td style="padding:8px 12px;text-align:center;">
        <button onclick="abrirModalUsuario('${u.id}')" style="padding:4px 10px;background:#f1f5f9;color:#374151;border:1px solid var(--borde);border-radius:5px;cursor:pointer;font-size:11px;">✏️ Editar</button>
      </td>
    </tr>`;
  }).join("");
}

function toggleUnidadField() {
  const rol = document.getElementById("adm-usr-rol")?.value;
  const wrap = document.getElementById("adm-usr-unidad-wrap");
  if (wrap) wrap.style.opacity = (rol==="Unidad") ? "1" : "0.4";
}

function abrirModalUsuario(id) {
  const modal = document.getElementById("adm-usr-modal");
  modal.style.display = "flex";
  const u = id ? _adminUsuarios.find(x=>x.id===id) : null;
  document.getElementById("adm-usr-modal-title").textContent = u ? "Editar Usuario" : "Nuevo Usuario";
  document.getElementById("adm-usr-id").value    = u?.id||"";
  document.getElementById("adm-usr-nombre").value = u?.NombreCompleto||"";
  document.getElementById("adm-usr-correo").value = u?.Correo||"";
  document.getElementById("adm-usr-rol").value    = u?.Rol||"Unidad";
  document.getElementById("adm-usr-unidad").value = u?.Unidad||"";
  document.getElementById("adm-usr-activo").checked  = u ? (u.Activo!==false&&u.Activo!==0) : true;
  document.getElementById("adm-usr-derivar").checked = u ? (u.PuedeDerivar===true||u.PuedeDerivar===1) : false;
  document.getElementById("adm-usr-cerrar").checked  = u ? (u.PuedeCerrar===true||u.PuedeCerrar===1) : false;
  document.getElementById("adm-usr-esadmin").checked = u ? (u.EsAdministrador===true||u.EsAdministrador===1) : false;
  toggleUnidadField();
}

function cerrarModalUsuario() {
  const modal = document.getElementById("adm-usr-modal");
  if (modal) modal.style.display = "none";
}

async function guardarUsuarioAdmin() {
  const id      = document.getElementById("adm-usr-id")?.value;
  const nombre  = document.getElementById("adm-usr-nombre")?.value.trim();
  const correo  = document.getElementById("adm-usr-correo")?.value.trim();
  const rol     = document.getElementById("adm-usr-rol")?.value;
  const unidad  = document.getElementById("adm-usr-unidad")?.value;
  const activo   = document.getElementById("adm-usr-activo")?.checked;
  const derivar  = document.getElementById("adm-usr-derivar")?.checked;
  const cerrar   = document.getElementById("adm-usr-cerrar")?.checked;
  const esAdmin  = document.getElementById("adm-usr-esadmin")?.checked;
  if (!nombre||!correo||!rol) { showToast("error","Nombre, correo y rol son obligatorios"); return; }
  showLoading("Guardando usuario...");
  try {
    const fields = { NombreCompleto:nombre, Correo:correo, Rol:rol, Unidad:unidad||"",
                     Activo:activo, PuedeDerivar:derivar, PuedeCerrar:cerrar, EsAdministrador:esAdmin };
    if (id) {
      await actualizarUsuario(id, fields);
      const u = _adminUsuarios.find(x=>x.id===id);
      if (u) Object.assign(u, fields);
      showToast("success","✅ Usuario actualizado");
    } else {
      const nuevo = await crearUsuario(fields);
      _adminUsuarios.push({ id:String(nuevo.Id||nuevo.id), ...fields });
      showToast("success","✅ Usuario creado");
    }
    cerrarModalUsuario();
    filtrarUsuarios();
  } catch(e) { showToast("error","Error: "+e.message); }
  finally { hideLoading(); }
}

// ===== REPORTES =====
async function renderGraficos() {
  const cont = document.getElementById("view-graficos");
  const hoy    = new Date().toISOString().split('T')[0];
  const hace3m = new Date(new Date().setMonth(new Date().getMonth()-3)).toISOString().split('T')[0];

  cont.innerHTML = `
  <style>
    @keyframes dashFadeUp { from { opacity:0; transform:translateY(18px); } to { opacity:1; transform:translateY(0); } }
    .da { animation: dashFadeUp .45s ease forwards; opacity:0; }
    .da1{animation-delay:.05s} .da2{animation-delay:.10s} .da3{animation-delay:.15s}
    .da4{animation-delay:.20s} .da5{animation-delay:.25s} .da6{animation-delay:.30s}
    .da7{animation-delay:.35s} .da8{animation-delay:.40s}
    .dash-kpi { background:white; border-radius:14px; padding:18px 20px; border:1px solid #e8eef6;
      box-shadow:0 2px 12px rgba(15,37,71,.06); cursor:default; transition:transform .2s,box-shadow .2s; }
    .dash-kpi:hover { transform:translateY(-2px); box-shadow:0 6px 20px rgba(15,37,71,.12); }
    .dash-kpi .kv { font-size:36px; font-weight:800; line-height:1; margin:8px 0 4px; font-variant-numeric:tabular-nums; }
    .dash-kpi .kl { font-size:11px; font-weight:600; letter-spacing:.4px; text-transform:uppercase; color:#94a3b8; }
    .dash-kpi .kb { height:3px; border-radius:2px; margin-top:12px; background:#e8eef6; overflow:hidden; }
    .dash-kpi .kb-fill { height:100%; border-radius:2px; transition:width 1s ease; }
    .dash-chart { background:white; border-radius:14px; padding:20px; border:1px solid #e8eef6;
      box-shadow:0 2px 12px rgba(15,37,71,.06); }
    .dash-chart h3 { font-size:13px; font-weight:700; color:#0f2547; margin:0 0 14px;
      display:flex; align-items:center; gap:6px; }
    .dash-chart h3 span.badge-per { font-size:10px; font-weight:600; background:#eff6ff;
      color:#1d4ed8; padding:2px 8px; border-radius:10px; margin-left:auto; }
    .sem-box { border-radius:10px; padding:14px 16px; display:flex; align-items:center; gap:12px; }
    .sem-num { font-size:32px; font-weight:800; line-height:1; }
    .sem-lbl { font-size:11px; font-weight:600; letter-spacing:.3px; text-transform:uppercase; margin-top:3px; }
    .dash-per-btn { padding:5px 11px; border:1.5px solid rgba(255,255,255,.22);
      border-radius:7px; background:rgba(255,255,255,.08); color:rgba(255,255,255,.85);
      cursor:pointer; font-size:11px; font-weight:600; transition:background .15s; }
    .dash-per-btn:hover { background:rgba(255,255,255,.2); }
    #tabla-unidades-body tr { transition:background .15s; }
  </style>

  <div style="display:flex;flex-direction:column;height:calc(100vh - 62px);overflow:hidden;">

    <!-- ── Header oscuro ── -->
    <div style="background:linear-gradient(135deg,#0f2547 0%,#1a3a6b 100%);padding:14px 24px;flex-shrink:0;display:flex;flex-wrap:wrap;gap:10px;align-items:center;">
      <div style="margin-right:12px;">
        <div style="font-size:16px;font-weight:800;color:white;letter-spacing:-.2px;">Dashboard DOM</div>
        <div style="font-size:11px;color:rgba(255,255,255,.5);margin-top:1px;">Dirección de Obras · Doñihue</div>
      </div>
      <div style="display:flex;gap:4px;flex-wrap:wrap;">
        ${[["Este mes",1],["3 meses",3],["6 meses",6],["1 año",12],["Todo",0]].map(([l,m])=>
          `<button class="dash-per-btn" onclick="setPeriodo(${m})">${l}</button>`).join("")}
      </div>
      <div style="display:flex;align-items:center;gap:6px;margin-left:6px;">
        <input type="date" id="graf-desde" value="${hace3m}"
          style="padding:5px 9px;border:1.5px solid rgba(255,255,255,.22);border-radius:7px;font-size:12px;background:rgba(255,255,255,.08);color:white;cursor:pointer;"
          onchange="actualizarGraficos()">
        <span style="color:rgba(255,255,255,.4);font-size:12px;">→</span>
        <input type="date" id="graf-hasta" value="${hoy}"
          style="padding:5px 9px;border:1.5px solid rgba(255,255,255,.22);border-radius:7px;font-size:12px;background:rgba(255,255,255,.08);color:white;cursor:pointer;"
          onchange="actualizarGraficos()">
      </div>
      <div style="margin-left:auto;display:flex;gap:6px;">
        <button onclick="actualizarGraficos()" class="dash-per-btn">⟳ Actualizar</button>
        <button onclick="exportarExcel()" style="padding:5px 13px;background:#C9A84C;color:#0f2547;border:none;border-radius:7px;cursor:pointer;font-size:11px;font-weight:800;">⬇ CSV</button>
      </div>
    </div>

    <!-- ── Cuerpo scrollable ── -->
    <div style="flex:1;overflow-y:auto;padding:18px 20px;background:#f0f4f8;display:flex;flex-direction:column;gap:16px;">

      <!-- KPIs -->
      <div id="graf-kpis" style="display:grid;grid-template-columns:repeat(5,1fr);gap:12px;"></div>

      <!-- Fila 1: doughnut + línea tendencia -->
      <div style="display:grid;grid-template-columns:340px 1fr;gap:16px;">
        <div class="dash-chart da da3">
          <h3>⬤ Por estado <span class="badge-per" id="badge-per-estado"></span></h3>
          <div style="display:flex;align-items:center;gap:14px;">
            <div style="width:140px;flex-shrink:0;"><canvas id="chart-estado"></canvas></div>
            <div id="legend-estado" style="font-size:12px;display:flex;flex-direction:column;gap:7px;flex:1;"></div>
          </div>
        </div>
        <div class="dash-chart da da4">
          <h3>📅 Evolución mensual
            <span style="display:flex;gap:10px;margin-left:auto;font-size:11px;font-weight:500;">
              <span style="display:flex;align-items:center;gap:4px;"><span style="width:14px;height:2.5px;background:#1a3a6b;display:inline-block;border-radius:2px;"></span>Ingresadas</span>
              <span style="display:flex;align-items:center;gap:4px;"><span style="width:14px;height:2.5px;background:#16a34a;display:inline-block;border-radius:2px;border-style:dashed;"></span>Cerradas</span>
            </span>
          </h3>
          <canvas id="chart-mes" height="95"></canvas>
        </div>
      </div>

      <!-- Fila 2: barras unidad + semáforo plazos -->
      <div style="display:grid;grid-template-columns:1fr 320px;gap:16px;">
        <div class="dash-chart da da5">
          <h3>🏢 Solicitudes por unidad</h3>
          <canvas id="chart-unidad" height="130"></canvas>
        </div>
        <div class="dash-chart da da6">
          <h3>🚦 Plazos activos</h3>
          <div id="dash-semaforo" style="display:flex;flex-direction:column;gap:8px;"></div>
        </div>
      </div>

      <!-- Tabla resumen -->
      <div class="dash-chart da da7" style="padding:0;overflow:hidden;">
        <div style="padding:13px 18px;border-bottom:1px solid #e8eef6;display:flex;justify-content:space-between;align-items:center;">
          <span style="font-size:13px;font-weight:700;color:#0f2547;">🏢 Rendimiento por unidad</span>
          <span id="tabla-periodo" style="font-size:11px;color:#94a3b8;"></span>
        </div>
        <div style="overflow-x:auto;">
          <table style="width:100%;border-collapse:collapse;font-size:12px;">
            <thead>
              <tr style="background:#f8fafc;">
                <th style="padding:9px 16px;text-align:left;font-weight:700;color:#64748b;font-size:11px;text-transform:uppercase;letter-spacing:.4px;border-bottom:1px solid #e8eef6;">Unidad</th>
                <th style="padding:9px 10px;text-align:center;color:#3b82f6;font-size:11px;border-bottom:1px solid #e8eef6;">Derivadas</th>
                <th style="padding:9px 10px;text-align:center;color:#22c55e;font-size:11px;border-bottom:1px solid #e8eef6;">Respondidas</th>
                <th style="padding:9px 10px;text-align:center;color:#6b7280;font-size:11px;border-bottom:1px solid #e8eef6;">Cerradas</th>
                <th style="padding:9px 10px;text-align:center;color:#7e22ce;font-size:11px;border-bottom:1px solid #e8eef6;">Pend.</th>
                <th style="padding:9px 10px;text-align:center;color:#f59e0b;font-size:11px;border-bottom:1px solid #e8eef6;">En Proc.</th>
                <th style="padding:9px 16px;text-align:left;font-size:11px;border-bottom:1px solid #e8eef6;min-width:160px;">Efectividad</th>
              </tr>
            </thead>
            <tbody id="tabla-unidades-body"></tbody>
          </table>
        </div>
      </div>

    </div>
  </div>`;

  await actualizarGraficos();
}

function setPeriodo(meses) {
  const hoy = new Date();
  const hasta = hoy.toISOString().split('T')[0];
  let desde;
  if (meses === 0) {
    desde = "2000-01-01";
  } else if (meses === 1) {
    desde = new Date(hoy.getFullYear(), hoy.getMonth(), 1).toISOString().split('T')[0];
  } else {
    desde = new Date(new Date().setMonth(hoy.getMonth() - meses)).toISOString().split('T')[0];
  }
  const d = document.getElementById("graf-desde");
  const h = document.getElementById("graf-hasta");
  if (d) d.value = desde;
  if (h) h.value = hasta;
  actualizarGraficos();
}

function _countUp(id, target, suffix = "", dur = 900) {
  const el = document.getElementById(id);
  if (!el) return;
  const t0 = performance.now();
  (function tick(now) {
    const p = Math.min((now - t0) / dur, 1);
    el.textContent = Math.round((1 - Math.pow(1 - p, 3)) * target) + suffix;
    if (p < 1) requestAnimationFrame(tick);
  })(t0);
}

async function actualizarGraficos() {
  showLoading("Cargando datos...");
  try {
    const desde = document.getElementById("graf-desde")?.value;
    const hasta  = document.getElementById("graf-hasta")?.value;
    const all    = await getSolicitudes();
    await new Promise(r => requestAnimationFrame(r));

    const filtradas = all.filter(s => {
      const f = new Date(s.FechaRecepcion);
      return (!desde || f >= new Date(desde)) && (!hasta || f <= new Date(hasta + "T23:59:59"));
    });

    const CE = {
      "Ingresada":"#3b82f6","Derivada":"#f59e0b","En Proceso":"#06b6d4",
      "Respondida":"#22c55e","Devuelta":"#ef4444","Pendiente de Cierre":"#7e22ce","Cerrada":"#64748b"
    };

    // ── KPIs ──
    const total      = filtradas.length;
    const cerradas   = filtradas.filter(s => s.Estado === "Cerrada").length;
    const respondidas= filtradas.filter(s => s.Estado === "Respondida").length;
    const pendCierre = filtradas.filter(s => s.Estado === "Pendiente de Cierre").length;
    const devueltas  = filtradas.filter(s => s.Estado === "Devuelta").length;
    const tasa       = total > 0 ? Math.round(((cerradas + pendCierre) / total) * 100) : 0;

    const kpisConf = [
      { id:"kv-total", label:"Total período",   val:total,       suf:"", color:"#1a3a6b", bar:100,                    del:1 },
      { id:"kv-cerr",  label:"Cerradas",         val:cerradas,    suf:"", color:"#64748b", bar:total?cerradas/total*100:0, del:2 },
      { id:"kv-resp",  label:"Respondidas",      val:respondidas, suf:"", color:"#16a34a", bar:total?respondidas/total*100:0, del:3 },
      { id:"kv-pend",  label:"Pend. de Cierre",  val:pendCierre,  suf:"", color:"#7e22ce", bar:total?pendCierre/total*100:0, del:4 },
      { id:"kv-tasa",  label:"Tasa resolución",  val:tasa,        suf:"%",color:"#0e7490", bar:tasa,                   del:5 },
    ];
    const kpisCont = document.getElementById("graf-kpis");
    if (kpisCont) {
      kpisCont.innerHTML = kpisConf.map(k => `
        <div class="dash-kpi da da${k.del}">
          <div class="kl">${k.label}</div>
          <div class="kv" id="${k.id}" style="color:${k.color};">0${k.suf}</div>
          <div class="kb"><div class="kb-fill" id="${k.id}-bar" style="width:0%;background:${k.color};"></div></div>
        </div>`).join("");
      // Lanzar count-up + barra con pequeño delay para que el DOM pinte primero
      setTimeout(() => {
        kpisConf.forEach(k => {
          _countUp(k.id, k.val, k.suf);
          const bar = document.getElementById(k.id + "-bar");
          if (bar) setTimeout(() => bar.style.width = k.bar.toFixed(1) + "%", 80);
        });
      }, 60);
    }

    // ── Badge período ──
    const bp = document.getElementById("badge-per-estado");
    if (bp && desde && hasta) bp.textContent = `${desde.slice(5)} → ${hasta.slice(5)}`;

    // ── Chart 1: Doughnut estado ──
    const byEstado = {};
    filtradas.forEach(s => { byEstado[s.Estado] = (byEstado[s.Estado]||0)+1; });
    const eLabels = Object.keys(byEstado);
    const eData   = Object.values(byEstado);
    const eColors = eLabels.map(e => CE[e]||"#94a3b8");

    if (state.chartInstances["chart-estado"]) state.chartInstances["chart-estado"].destroy();
    const ctx1 = document.getElementById("chart-estado")?.getContext("2d");
    if (ctx1) {
      state.chartInstances["chart-estado"] = new Chart(ctx1, {
        type: "doughnut",
        data: { labels: eLabels, datasets: [{ data:eData, backgroundColor:eColors, borderWidth:3, borderColor:"white", hoverOffset:6 }] },
        options: { responsive:true, cutout:"70%", animation:{ animateRotate:true, duration:900 },
          plugins:{ legend:{ display:false }, tooltip:{ callbacks:{ label: c => ` ${c.label}: ${c.raw} (${Math.round(c.raw/total*100)}%)` } } } }
      });
    }
    const legCont = document.getElementById("legend-estado");
    if (legCont) legCont.innerHTML = eLabels.map((e,i) => `
      <div style="display:flex;align-items:center;gap:7px;font-size:12px;">
        <span style="width:10px;height:10px;border-radius:3px;background:${eColors[i]};flex-shrink:0;"></span>
        <span style="color:#475569;flex:1;">${e}</span>
        <span style="font-weight:700;color:${eColors[i]};">${eData[i]}</span>
      </div>`).join("");

    // ── Chart 2: Línea mensual con gradiente ──
    const byMesIng = {}, byMesCer = {};
    filtradas.forEach(s => { const m = s.FechaRecepcion?.substring(0,7); if(m) byMesIng[m]=(byMesIng[m]||0)+1; });
    filtradas.filter(s=>s.Estado==="Cerrada").forEach(s => { const m = s.FechaRecepcion?.substring(0,7); if(m) byMesCer[m]=(byMesCer[m]||0)+1; });
    const meses = [...new Set([...Object.keys(byMesIng),...Object.keys(byMesCer)])].sort();
    const MESES_ES = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
    const mesesLabel = meses.map(m => { const[y,mo]=m.split("-"); return `${MESES_ES[+mo-1]} ${y.slice(2)}`; });

    if (state.chartInstances["chart-mes"]) state.chartInstances["chart-mes"].destroy();
    const ctx3 = document.getElementById("chart-mes")?.getContext("2d");
    if (ctx3) {
      const g1 = ctx3.createLinearGradient(0,0,0,220);
      g1.addColorStop(0,"rgba(26,58,107,.18)"); g1.addColorStop(1,"rgba(26,58,107,0)");
      const g2 = ctx3.createLinearGradient(0,0,0,220);
      g2.addColorStop(0,"rgba(22,163,74,.14)"); g2.addColorStop(1,"rgba(22,163,74,0)");
      state.chartInstances["chart-mes"] = new Chart(ctx3, {
        type:"line",
        data:{ labels:mesesLabel, datasets:[
          { label:"Ingresadas", data:meses.map(m=>byMesIng[m]||0), borderColor:"#1a3a6b", backgroundColor:g1, fill:true, tension:.45, pointRadius:4, pointHoverRadius:6, pointBackgroundColor:"#1a3a6b", borderWidth:2.5 },
          { label:"Cerradas",   data:meses.map(m=>byMesCer[m]||0), borderColor:"#16a34a", backgroundColor:g2, fill:true, tension:.45, pointRadius:4, pointHoverRadius:6, pointBackgroundColor:"#16a34a", borderWidth:2, borderDash:[6,3] }
        ]},
        options:{ responsive:true, animation:{ duration:900 },
          plugins:{ legend:{ display:false }, tooltip:{ mode:"index", intersect:false } },
          scales:{ y:{ beginAtZero:true, ticks:{ stepSize:1, font:{ size:11 } }, grid:{ color:"#f1f5f9" } },
                   x:{ ticks:{ font:{ size:11 } }, grid:{ display:false } } } }
      });
    }

    // ── Chart 3: Barras horizontales unidades ──
    const byUnidad = {};
    filtradas.filter(s=>s.UnidadDerivada).forEach(s => { byUnidad[(s.UnidadDerivada||"").trim()]=(byUnidad[(s.UnidadDerivada||"").trim()]||0)+1; });
    const uLabels = Object.keys(byUnidad).sort((a,b)=>byUnidad[b]-byUnidad[a]);
    const uData   = uLabels.map(u=>byUnidad[u]);
    const uColors = ["#1a3a6b","#2563eb","#3b82f6","#60a5fa","#93c5fd"].slice(0,uLabels.length);

    if (state.chartInstances["chart-unidad"]) state.chartInstances["chart-unidad"].destroy();
    const ctx2 = document.getElementById("chart-unidad")?.getContext("2d");
    if (ctx2) {
      state.chartInstances["chart-unidad"] = new Chart(ctx2, {
        type:"bar",
        data:{ labels:uLabels, datasets:[{ data:uData, backgroundColor:uColors, borderRadius:6, borderSkipped:false }] },
        options:{ indexAxis:"y", responsive:true, animation:{ duration:800 },
          plugins:{ legend:{ display:false }, tooltip:{ callbacks:{ label: c=>`  ${c.raw} solicitudes` } } },
          scales:{ x:{ beginAtZero:true, ticks:{ stepSize:1, font:{size:11} }, grid:{ color:"#f1f5f9" } },
                   y:{ ticks:{ font:{size:11} }, grid:{ display:false } } } }
      });
    }

    // ── Semáforo de plazos (sobre solicitudes activas en todo el sistema) ──
    const activas = all.filter(s => s.Estado===CONFIG.estados.DERIVADA||s.Estado===CONFIG.estados.EN_PROCESO);
    const semVerde   = activas.filter(s => { const r=calcularSemaforo(s); return r&&r.dias>3; }).length;
    const semAmarillo= activas.filter(s => { const r=calcularSemaforo(s); return r&&r.dias>0&&r.dias<=3; }).length;
    const semRojo    = activas.filter(s => { const r=calcularSemaforo(s); return r&&r.dias<=0; }).length;
    const semTotal   = semVerde+semAmarillo+semRojo;
    const semCont = document.getElementById("dash-semaforo");
    if (semCont) {
      if (!semTotal) {
        semCont.innerHTML = `<div style="text-align:center;color:#94a3b8;padding:20px;font-size:13px;">Sin solicitudes activas</div>`;
      } else {
        const semItems = [
          { n:semVerde,    label:"En plazo",    sub:">3 días restantes",  bg:"#f0fdf4", bc:"#86efac", tc:"#15803d", emoji:"🟢" },
          { n:semAmarillo, label:"Por vencer",  sub:"1–3 días",           bg:"#fffbeb", bc:"#fde68a", tc:"#b45309", emoji:"🟡" },
          { n:semRojo,     label:"Crítico",     sub:"Vencidas o hoy",     bg:"#fef2f2", bc:"#fca5a5", tc:"#b91c1c", emoji:"🔴" },
        ];
        semCont.innerHTML = semItems.map(s => `
          <div class="sem-box" style="background:${s.bg};border:1.5px solid ${s.bc};">
            <span style="font-size:24px;">${s.emoji}</span>
            <div style="flex:1;">
              <div style="display:flex;align-items:baseline;gap:6px;">
                <span class="sem-num" style="color:${s.tc};">${s.n}</span>
                <span class="sem-lbl" style="color:${s.tc};">${s.label}</span>
              </div>
              <div style="font-size:11px;color:${s.tc};opacity:.7;margin-top:2px;">${s.sub}</div>
            </div>
            <div style="font-size:13px;font-weight:700;color:${s.tc};opacity:.6;">${semTotal?Math.round(s.n/semTotal*100):0}%</div>
          </div>`).join("");
      }
    }

    // ── Tabla unidades ──
    const unidades = [...new Set(filtradas.filter(s=>s.UnidadDerivada).map(s=>(s.UnidadDerivada||"").trim()))].sort();
    const tbody = document.getElementById("tabla-unidades-body");
    const per   = document.getElementById("tabla-periodo");
    if (per) per.textContent = desde && hasta ? `${desde} — ${hasta}` : "Todo el período";
    if (tbody) {
      tbody.innerHTML = unidades.map((u,ri) => {
        const sols = filtradas.filter(s=>(s.UnidadDerivada||"").trim()===u);
        const der  = sols.length;
        const resp = sols.filter(s=>s.Estado==="Respondida").length;
        const cerr = sols.filter(s=>s.Estado==="Cerrada").length;
        const penc = sols.filter(s=>s.Estado==="Pendiente de Cierre").length;
        const proc = sols.filter(s=>s.Estado==="En Proceso").length;
        const efe  = der>0 ? Math.round(((resp+cerr+penc)/der)*100) : 0;
        const ec   = efe>=80?"#15803d":efe>=50?"#b45309":"#b91c1c";
        const chip = (n,bg,tc) => `<span style="background:${bg};color:${tc};padding:2px 9px;border-radius:20px;font-size:11px;font-weight:700;">${n}</span>`;
        return `<tr style="border-bottom:1px solid #f1f5f9;${ri%2?"background:#fafbfc":""}"
          onmouseenter="this.style.background='#eff6ff'" onmouseleave="this.style.background='${ri%2?"#fafbfc":""}'">
          <td style="padding:10px 16px;font-weight:700;color:#1a3a6b;">${u}</td>
          <td style="padding:8px 10px;text-align:center;">${chip(der,"#dbeafe","#1d4ed8")}</td>
          <td style="padding:8px 10px;text-align:center;">${chip(resp,"#dcfce7","#15803d")}</td>
          <td style="padding:8px 10px;text-align:center;">${chip(cerr,"#f1f5f9","#475569")}</td>
          <td style="padding:8px 10px;text-align:center;">${chip(penc,"#fdf4ff","#7e22ce")}</td>
          <td style="padding:8px 10px;text-align:center;">${chip(proc,"#fef3c7","#b45309")}</td>
          <td style="padding:10px 16px;min-width:160px;">
            <div style="display:flex;align-items:center;gap:8px;">
              <div style="flex:1;height:7px;background:#e2e8f0;border-radius:4px;overflow:hidden;">
                <div style="height:100%;width:${efe}%;background:${ec};border-radius:4px;transition:width .8s ease;"></div>
              </div>
              <span style="font-size:12px;font-weight:800;color:${ec};min-width:34px;">${efe}%</span>
            </div>
          </td>
        </tr>`;
      }).join("") || `<tr><td colspan="7" style="padding:28px;text-align:center;color:#94a3b8;font-size:13px;">Sin datos para el período</td></tr>`;
    }

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

// ===== NAVEGACIÓN MÓVIL =====
function isMobile() { return window.innerWidth <= 768; }

function mostrarDetalleMovil(layoutClass) {
  if (!isMobile()) return;
  document.querySelector(layoutClass)?.classList.add('showing-detalle');
}

function volverAListaMovil(layoutClass) {
  document.querySelector(layoutClass)?.classList.remove('showing-detalle');
  state.solicitudSeleccionada = null;
  if (layoutClass === '.sec-layout') {
    renderListaSolicitudes();
    renderFormNueva();
    const header = document.getElementById("form-panel-header");
    if (header) { header.textContent = "➕ Nueva Solicitud"; header.style.cssText = ""; }
  }
  if (layoutClass === '.dir-layout') { renderDirLista(); }
  if (layoutClass === '.uni-layout') { renderSidebarUnidad(); }
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
  mostrarDetalleMovil('.sec-layout');
}

function cargarSolicitudEnFormulario(sol) {
  // ── Header dinámico: modo edición vs solo lectura ──
  const editable = sol.Estado === CONFIG.estados.INGRESADA;
  const header = document.getElementById("form-panel-header");
  if (header) {
    header.style.cssText = `display:flex;flex-direction:column;padding:0;color:white;
      background:${editable ? 'linear-gradient(90deg,#166534,#15803d)' : 'linear-gradient(90deg,#0f2547,#1a3a6b)'};`;
    header.innerHTML = `
      <button class="mobile-back-bar" onclick="volverAListaMovil('.sec-layout')" style="font-size:13px;padding:8px 12px;">
        ← Volver a lista
      </button>
      <div style="display:flex;align-items:center;gap:10px;flex:1;padding:10px 16px;">
        <span style="font-size:16px;">${editable ? '✏️' : '👁️'}</span>
        <span style="font-weight:600;font-size:13px;opacity:0.9;">${editable ? 'Editando' : 'Viendo'}</span>
        <span style="font-weight:800;font-size:16px;letter-spacing:0.5px;color:white;">Solicitud #${sol.NroSolicitud}</span>
        <span style="background:rgba(255,255,255,0.2);color:white;font-size:11px;font-weight:700;
              padding:3px 10px;border-radius:12px;letter-spacing:0.3px;">${sol.Estado}</span>
        <button onclick="limpiarSeleccion()" title="Cerrar"
          style="margin-left:auto;background:rgba(255,255,255,0.15);border:1.5px solid rgba(255,255,255,0.4);
                 color:white;padding:5px 16px;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600;">
          ✕ Cerrar
        </button>
      </div>`;
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

      <!-- ── Respuesta de la Unidad (solo Respondida / Cerrada) ── -->
      ${(sol.Estado === CONFIG.estados.RESPONDIDA || sol.Estado === CONFIG.estados.CERRADA) ? `
      <div style="background:white;border:2px solid #15803d;border-radius:10px;overflow:hidden;">
        <div style="background:linear-gradient(90deg,#14532d,#15803d);color:white;padding:8px 14px;font-size:12px;font-weight:700;letter-spacing:0.3px;">
          🏢 Respuesta de la Unidad
        </div>
        <div id="sec-evidencia-panel" style="padding:0;">
          <div style="text-align:center;color:#9ca3af;font-size:12px;padding:16px;">
            <div class="spinner" style="margin:0 auto 8px;width:20px;height:20px;border-width:2px;"></div>
            Cargando respuesta...
          </div>
        </div>
      </div>` : ''}

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

  // ── Cargar respuesta de la unidad para Respondida/Cerrada ──
  if (sol.Estado === CONFIG.estados.RESPONDIDA || sol.Estado === CONFIG.estados.CERRADA) {
    getEvidenciasBySolicitud(sol.NroSolicitud, sol.id).then(evidencias => {
      const panel = document.getElementById("sec-evidencia-panel");
      if (!panel) return;
      if (!evidencias.length) {
        panel.innerHTML = `<div style="text-align:center;padding:16px;color:#9ca3af;font-size:13px;">Sin respuesta registrada por la unidad</div>`;
        return;
      }
      panel.innerHTML = evidencias.map(ev => `
        <div style="border-bottom:1px solid #e2e8f0;padding:12px 14px;">
          <div style="font-size:12px;font-weight:700;color:#15803d;margin-bottom:6px;">
            🏢 ${ev.Unidad||"Unidad"} &nbsp;·&nbsp; 👤 ${ev.Responsable||""} &nbsp;·&nbsp; 📅 ${formatFecha(ev.FechaCarga)}
          </div>
          <div style="font-size:13px;color:#374151;line-height:1.7;background:#f0fdf4;border-left:3px solid #15803d;padding:10px 12px;border-radius:0 8px 8px 0;">
            ${ev.DescripcionEvidencia||"Sin descripción."}
          </div>
        </div>`).join("");
    }).catch(e => {
      console.warn("Evidencia secretaria:", e?.message);
      const panel = document.getElementById("sec-evidencia-panel");
      if (panel) panel.innerHTML = `<div style="text-align:center;padding:16px;color:#9ca3af;font-size:13px;">Sin respuesta registrada por la unidad</div>`;
    });
  }

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
    hc.innerHTML = hist.map((h,i) => {
      const accion  = h.Accion||h.Title||h.title||"—";
      const usuario = h.UsuarioAccion||"";
      const unidad  = h.Unidad||"";
      const obs     = h.Observaciones||"";
      const fecha   = h.FechaAccion||h.Modified||"";
      const estAnt  = h.EstadoAnterior||"";
      const estNuevo= h.EstadoNuevo||"";
      return `
      <div style="display:flex;gap:12px;padding:10px 0;${i<hist.length-1?'border-bottom:1px solid #f3f4f6':''}">
        <div style="display:flex;flex-direction:column;align-items:center;gap:2px;">
          <div style="width:12px;height:12px;border-radius:50%;background:${dotColor(accion)};flex-shrink:0;margin-top:2px;"></div>
          ${i<hist.length-1?`<div style="width:2px;flex:1;background:#f0f0f0;margin-top:2px;"></div>`:''}
        </div>
        <div style="flex:1;padding-bottom:4px;">
          <div style="font-size:13px;font-weight:700;color:#1a1a1a;">${accion}</div>
          <div style="font-size:11px;color:#888;margin-top:1px;">
            📅 ${formatFechaHora(fecha)}
            ${usuario?`· 👤 ${usuario}`:''}
            ${unidad?`· 🏢 ${unidad}`:''}
          </div>
          ${estAnt?`<div style="font-size:11px;color:#aaa;margin-top:2px;">
            <span class="estado-badge estado-${estAnt}" style="font-size:10px;">${estAnt}</span>
            → <span class="estado-badge estado-${estNuevo}" style="font-size:10px;">${estNuevo}</span>
          </div>`:''}
          ${obs?`<div style="font-size:12px;color:#4b5563;margin-top:4px;padding:6px 8px;background:#f9fafb;border-left:3px solid ${dotColor(accion)};border-radius:0 4px 4px 0;">"${obs}"</div>`:''}
        </div>
      </div>`;}).join('');
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
    }  // cierre if (header)

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
  document.querySelector('.sec-layout')?.classList.remove('showing-detalle');
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

function calcularSemaforo(sol) {
  const activos = [CONFIG.estados.DERIVADA, CONFIG.estados.EN_PROCESO];
  if (!activos.includes(sol.Estado)) return null;
  const inicio      = new Date(sol.FechaDerivacion || sol.FechaRecepcion);
  if (isNaN(inicio)) return null;
  const plazo       = CONFIG.plazoDerivacionDias || 15;
  const vencimiento = new Date(inicio.getTime() + plazo * 864e5);
  const dias        = Math.ceil((vencimiento - new Date()) / 864e5);
  if (dias > 7)  return { color:"#16a34a", bg:"#dcfce7", emoji:"🟢", texto:`${dias}d`,        dias, vencimiento };
  if (dias > 3)  return { color:"#d97706", bg:"#fef3c7", emoji:"🟡", texto:`${dias}d`,        dias, vencimiento };
  if (dias > 0)  return { color:"#dc2626", bg:"#fee2e2", emoji:"🔴", texto:`${dias}d`,        dias, vencimiento };
  if (dias === 0) return { color:"#dc2626", bg:"#fee2e2", emoji:"🔴", texto:"Hoy",            dias, vencimiento };
  return           { color:"#7f1d1d",  bg:"#fecaca",  emoji:"⚫", texto:`+${Math.abs(dias)}d`, dias, vencimiento };
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
    if (b.dataset.tab !== "solicitudes") return;
    let badge = b.querySelector(".badge");
    if (ingresadas > 0) {
      if (!badge) { badge = document.createElement("span"); badge.className = "badge warn"; b.appendChild(badge); }
      badge.textContent = ingresadas;
    } else if (badge) {
      badge.remove();
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
