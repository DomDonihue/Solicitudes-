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
      showLoginPage(msUser, `El usuario ${msUser.mail || msUser.userPrincipalName} no tiene acceso al sistema. Contacta al administrador TI.`);
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

function getDirFiltradas() {
  return state.solicitudes.filter(s => {
    if (state.filtroEstado !== "Todos" && s.Estado !== state.filtroEstado) return false;
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
      return `
        <div class="sol-card ${state.solicitudSeleccionada?.id===s.id?'selected':''}"
          style="border-left-color:${esUrgente?'#ef4444':esAccion?'var(--azul)':'var(--borde)'};"
          onclick="seleccionarSolicitudDirector('${s.id}')">
          <div class="sol-card-top">
            <span class="sol-nro">${s.NroSolicitud}</span>
            <span class="estado-badge estado-${s.Estado}">${s.Estado}</span>
          </div>
          <div class="sol-card-name">${s.Solicitante}</div>
          <div class="sol-card-dir">📍 ${s.Direccion||""}</div>
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

  cont.innerHTML = `<div style="display:flex;flex-direction:column;gap:12px;padding:14px;">

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
    <!-- Panel cerrar — dos opciones -->
    <div class="accion-panel">
      <div class="accion-header cerrar" style="background:linear-gradient(90deg,#14532d,#15803d);">📋 Resolución de Solicitud</div>
      <div class="accion-body">
        <p style="font-size:12px;color:#555;margin-bottom:10px;">
          ${esPendienteCierre
            ? '⏳ Solicitud en plazo de evaluación. Una vez generado el parte final, puede cerrarse formalmente.'
            : 'La unidad ha respondido. Seleccione cómo proceder:'}
        </p>
        <textarea id="dir-cierre-obs" rows="2" placeholder="Observaciones / número de parte final (opcional)..." style="margin-bottom:10px;"></textarea>

        ${!esPendienteCierre ? `
        <div style="background:#fdf4ff;border:1.5px solid #d8b4fe;border-radius:10px;padding:12px;margin-bottom:10px;">
          <div style="font-size:12px;font-weight:700;color:#7e22ce;margin-bottom:6px;">⏳ Pendiente de Cierre</div>
          <p style="font-size:11px;color:#6b21a8;margin:0 0 8px;">Queda en plazo de evaluación hasta que se genere el parte final de inspección.</p>
          <div style="display:flex;gap:8px;align-items:center;margin-bottom:8px;">
            <label style="font-size:11px;font-weight:600;color:#7e22ce;white-space:nowrap;">Plazo hasta:</label>
            <input type="date" id="dir-plazo-cierre"
              style="flex:1;padding:5px 8px;border:1.5px solid #d8b4fe;border-radius:6px;font-size:12px;"
              value="${new Date(new Date().setDate(new Date().getDate()+30)).toISOString().split('T')[0]}">
          </div>
          <button onclick="pendienteCierreSolicitud('${sol.id}')"
            style="width:100%;padding:9px;background:#7e22ce;color:white;border:none;border-radius:8px;cursor:pointer;font-size:13px;font-weight:600;">
            ⏳ Marcar Pendiente de Cierre
          </button>
        </div>` : ""}

        <div style="background:#f0fdf4;border:1.5px solid #86efac;border-radius:10px;padding:12px;">
          <div style="font-size:12px;font-weight:700;color:#15803d;margin-bottom:6px;">🔒 Cierre Definitivo</div>
          <p style="font-size:11px;color:#166534;margin:0 0 8px;">Cierra formalmente la solicitud. El caso queda archivado en el sistema.</p>
          <button class="btn-success" onclick="cerrarSolicitud('${sol.id}')" style="width:100%;padding:9px;">
            🔒 Cerrar Solicitud Formalmente
          </button>
        </div>
      </div>
    </div>` : ""}

    ${esPendienteCierre ? `
    <div style="background:#fdf4ff;border:1.5px solid #d8b4fe;border-radius:10px;padding:12px 14px;display:flex;align-items:center;gap:10px;margin-bottom:8px;">
      <span style="font-size:20px;">⏳</span>
      <div style="flex:1;">
        <div style="font-size:13px;font-weight:700;color:#7e22ce;">Pendiente de Cierre</div>
        <div style="font-size:11px;color:#6b21a8;">Aguardando parte final${sol.PlazoCierre ? ' — Plazo: ' + formatFecha(sol.PlazoCierre) : ''}</div>
      </div>
    </div>` : ""}

    ${sol.Estado === CONFIG.estados.CERRADA ? `
    <!-- Indicador de cierre + opción reabrir -->
    <div style="background:#f0fdf4;border:1.5px solid #86efac;border-radius:10px;padding:12px 14px;margin-bottom:8px;">
      <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;">
        <span style="font-size:20px;">🔒</span>
        <div>
          <div style="font-size:13px;font-weight:700;color:#15803d;">Solicitud Cerrada Formalmente</div>
          <div style="font-size:11px;color:#16a34a;">Caso completado y archivado en el sistema</div>
        </div>
      </div>
      <!-- Reabrir -->
      <div style="border-top:1px solid #bbf7d0;padding-top:10px;">
        <div style="font-size:11px;font-weight:700;color:#374151;margin-bottom:6px;">🔄 ¿Necesita reabrirse?</div>
        <select id="dir-reabrir-estado" style="width:100%;padding:6px 8px;border:1.5px solid var(--borde);border-radius:7px;font-size:12px;margin-bottom:6px;">
          <option value="">— Seleccionar nuevo estado —</option>
          <option value="Ingresada">📥 Ingresada</option>
          <option value="Derivada">📤 Derivada</option>
          <option value="En Proceso">⚙️ En Proceso</option>
          <option value="Respondida">✅ Respondida</option>
          <option value="Pendiente de Cierre">⏳ Pendiente de Cierre</option>
        </select>
        <textarea id="dir-reabrir-obs" rows="2" placeholder="Motivo de reapertura (obligatorio)..." style="margin-bottom:6px;"></textarea>
        <button onclick="reabrirSolicitud('${sol.id}')"
          style="width:100%;padding:8px;background:#1a3a6b;color:white;border:none;border-radius:8px;cursor:pointer;font-size:12px;font-weight:600;">
          🔄 Cambiar Estado
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
      window.open(blobUrl, "_blank");
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

async function cargarHistorialDir(nroSolicitud) {
  const hc = document.getElementById("dir-historial");
  if (!hc) return;
  hc.innerHTML = `<div style="text-align:center;color:#9ca3af;font-size:12px;padding:10px;"><div class="spinner" style="margin:0 auto 8px;width:20px;height:20px;border-width:2px;"></div>Cargando...</div>`;
  try {
    const hist = await getHistorialBySolicitud(nroSolicitud);
    hist.sort((a,b)=>new Date(b.FechaAccion)-new Date(a.FechaAccion));
    if (!hist.length) { hc.innerHTML=`<p style="text-align:center;color:#9ca3af;font-size:12px;">Sin historial</p>`; return; }
    const dot = a=>{a=(a||"").toLowerCase();return a.includes('deriv')?'#f59e0b':a.includes('respond')?'#22c55e':a.includes('devuel')?'#ef4444':a.includes('cerr')?'#6b7280':'#3b82f6';};
    hc.innerHTML = hist.map(h=>{
      // SharePoint puede devolver campos con nombre interno distinto
      const accion = h.Accion || h.Title || h.title || "—";
      const usuario = h.UsuarioAccion || h.UsuarioAccion0 || "";
      const unidad  = h.Unidad || "";
      const obs     = h.Observaciones || "";
      const fecha   = h.FechaAccion || h.Modified || "";
      const estAnt  = h.EstadoAnterior || "";
      const estNuevo= h.EstadoNuevo || "";
      return `
      <div style="display:flex;gap:8px;padding:7px 0;border-bottom:1px solid #f3f4f6;">
        <div style="width:8px;height:8px;border-radius:50%;background:${dot(accion)};margin-top:4px;flex-shrink:0;"></div>
        <div style="flex:1;">
          <div style="font-size:12px;font-weight:600;color:#1a1a1a;">${accion}</div>
          <div style="font-size:11px;color:#888;">${formatFechaHora(fecha)}${usuario?' · '+usuario:''}${unidad?' · '+unidad:''}</div>
          ${estAnt?`<div style="font-size:10px;color:#aaa;">${estAnt} → ${estNuevo}</div>`:""}
          ${obs?`<div style="font-size:11px;color:#555;font-style:italic;">"${obs}"</div>`:""}
        </div>
      </div>`;}).join("");
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
    await actualizarSolicitud(solId, { Estado:CONFIG.estados.DERIVADA, UnidadDerivada:unidad });
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
    await actualizarSolicitud(solId, {
      Estado: CONFIG.estados.PENDIENTE_CIERRE,
      PlazoCierre: plazo || null
    });
    registrarHistorial({
      NroSolicitud:sol.NroSolicitud, Title:"Pendiente de Cierre — en plazo de evaluación",
      EstadoAnterior:sol.Estado, EstadoNuevo:CONFIG.estados.PENDIENTE_CIERRE,
      UsuarioAccion:state.usuario.NombreCompleto, RolUsuario:state.usuario.Rol,
      Unidad:state.usuario.Unidad, FechaAccion:new Date().toISOString(),
      Observaciones: (obs || "") + (plazo ? ` | Plazo: ${plazo}` : "")
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
  mostrarDetalleMovil('.uni-layout');
}

async function renderDetalleUnidad(sol) {
  const cont = document.getElementById("uni-detalle");
  if (!sol) {
    cont.innerHTML = `<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#9ca3af;flex-direction:column;gap:16px;"><span style="font-size:60px">📋</span><p>Selecciona una solicitud</p></div>`;
    return;
  }

  // Cargar adjuntos de la solicitud y evidencias en paralelo
  const [solicitudAtts, evidencias] = await Promise.all([
    getListItemAttachments(CONFIG.lists.solicitudes, sol.id).catch(() => []),
    getEvidenciasBySolicitud(sol.NroSolicitud, sol.id).catch(() => [])
  ]);

  const puedeCerrar = state.usuario.PuedeCerrar && sol.Estado === CONFIG.estados.RESPONDIDA;

  cont.innerHTML = `
    <div style="overflow-y:auto;height:100%;display:flex;flex-direction:column;gap:0;">
      <button class="mobile-back-bar" onclick="volverAListaMovil('.uni-layout')">← Volver a lista</button>
      <div class="panel-header" style="background:#f8fafc;border-bottom:1px solid var(--borde);">📋 ${sol.NroSolicitud} — ${sol.Solicitante}</div>
      <div style="padding:12px;background:#f8fafc;border-bottom:1px solid var(--borde);">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:13px;">
          <div><span style="color:#666">Fecha: </span>${formatFecha(sol.FechaRecepcion)}</div>
          <div><span style="color:#666">Estado: </span><span class="estado-badge estado-${sol.Estado}">${sol.Estado}</span></div>
          <div style="grid-column:1/-1;"><span style="color:#666">Dirección: </span>${sol.Direccion||"-"}</div>
          <div style="grid-column:1/-1;"><span style="color:#666">Solicitud: </span>${sol.Solicitud||"-"}</div>
        </div>
      </div>
      <div id="uni-sol-adjuntos" style="padding:12px;border-bottom:1px solid var(--borde);">
        ${solicitudAtts.length === 0
          ? `<p style="color:#9ca3af;font-size:13px;text-align:center;padding:8px;">Sin documentos adjuntos</p>`
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

      ${evidencias.length > 0 ? `
      <div style="border-top:2px solid #15803d;">
        <div style="background:linear-gradient(90deg,#14532d,#15803d);color:white;padding:8px 14px;font-size:12px;font-weight:700;letter-spacing:0.3px;">
          🏢 Respuesta registrada
        </div>
        ${evidencias.map(e=>`
          <div style="padding:12px 14px;border-bottom:1px solid #e2e8f0;">
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
      <div style="padding:14px;display:flex;flex-direction:column;gap:10px;">
        <div class="form-section">
          <div class="form-section-header" style="font-size:11px;">📝 Descripción de la Solución</div>
          <div class="form-section-body">
            <textarea id="uni-obs" rows="4"
              placeholder="Describe detalladamente la acción realizada, visita, notificación o solución ejecutada..."
              style="width:100%;padding:10px;border:1.5px solid var(--borde);border-radius:8px;font-size:13px;font-family:inherit;resize:vertical;"></textarea>
          </div>
        </div>
        <div class="form-section">
          <div class="form-section-header naranja" style="font-size:11px;">📸 Evidencia Fotográfica / Documentos</div>
          <div class="form-section-body">
            <p style="font-size:12px;color:#888;margin-bottom:8px;">Adjunta fotos de la visita o documentos que corroboren la solución ejecutada.</p>
            <div style="border:2px dashed #f59e0b;border-radius:10px;padding:16px;text-align:center;cursor:pointer;background:#fffbeb;"
              onclick="document.getElementById('uni-ev-files').click()"
              ondragover="event.preventDefault();this.style.background='#fef3c7'"
              ondragleave="this.style.background='#fffbeb'"
              ondrop="event.preventDefault();this.style.background='#fffbeb';handleEvFiles(event.dataTransfer.files)">
              <div style="font-size:28px;">📷</div>
              <div style="font-weight:600;font-size:13px;color:#b45309;">Arrastra fotos aquí</div>
              <div style="font-size:11px;color:#9ca3af;margin-top:2px;">o haz clic — JPG, PNG, PDF</div>
            </div>
            <input type="file" id="uni-ev-files" multiple accept=".pdf,.jpg,.jpeg,.png" onchange="handleEvFiles(this.files)" style="display:none;">
            <div id="uni-ev-preview" style="margin-top:8px;display:grid;grid-template-columns:repeat(3,1fr);gap:6px;"></div>
          </div>
        </div>
        <button class="btn-success" onclick="responderSolicitud('${sol.id}')" style="width:100%;padding:13px;font-size:14px;">✅ Responder — Registrar Solución</button>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
          <button class="btn-primary" style="background:#0e7490;padding:10px;" onclick="enProcesoSolicitud('${sol.id}')">⚙️ En Proceso</button>
          <button class="btn-warning" style="padding:10px;" onclick="devolverSolicitudUnidad('${sol.id}')">↩️ Devolver</button>
        </div>
        <button onclick="verHistorial('${sol.NroSolicitud}')" style="width:100%;padding:9px;border:1.5px solid var(--borde);border-radius:8px;background:white;cursor:pointer;font-size:13px;color:#6b7280;">🕐 Ver Historial</button>
      </div>
    </div>`;

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

    registrarHistorial({
      NroSolicitud: sol.NroSolicitud,
      Title:"Solicitud respondida",
      EstadoAnterior: sol.Estado,
      EstadoNuevo: CONFIG.estados.RESPONDIDA,
      UsuarioAccion: state.usuario.NombreCompleto,
      RolUsuario: state.usuario.Rol,
      Unidad: state.usuario.Unidad,
      FechaAccion: new Date().toISOString(),
      Observaciones: obs
    }).catch(e => console.warn("Historial (no crítico):", e.message));

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

// ===== REPORTES =====
async function renderGraficos() {
  const cont = document.getElementById("view-graficos");
  const hoy = new Date().toISOString().split('T')[0];
  const hace3m = new Date(new Date().setMonth(new Date().getMonth()-3)).toISOString().split('T')[0];

  cont.innerHTML = `
  <div style="display:flex;flex-direction:column;height:calc(100vh - 120px);overflow:hidden;">

    <!-- Barra de controles -->
    <div style="background:white;border-bottom:2px solid var(--borde);padding:12px 20px;display:flex;flex-wrap:wrap;gap:12px;align-items:center;flex-shrink:0;">
      <div style="font-size:15px;font-weight:700;color:var(--azul);margin-right:8px;">📊 Reportes</div>
      <div style="display:flex;align-items:center;gap:8px;">
        <label style="font-size:12px;color:#666;font-weight:600;">Desde</label>
        <input type="date" id="graf-desde" value="${hace3m}"
          style="padding:6px 10px;border:1.5px solid var(--borde);border-radius:7px;font-size:13px;"
          onchange="actualizarGraficos()">
      </div>
      <div style="display:flex;align-items:center;gap:8px;">
        <label style="font-size:12px;color:#666;font-weight:600;">Hasta</label>
        <input type="date" id="graf-hasta" value="${hoy}"
          style="padding:6px 10px;border:1.5px solid var(--borde);border-radius:7px;font-size:13px;"
          onchange="actualizarGraficos()">
      </div>
      <!-- Accesos rápidos de período -->
      <div style="display:flex;gap:6px;">
        <button onclick="setPeriodo(1)" style="padding:5px 10px;border:1.5px solid var(--borde);border-radius:6px;background:white;cursor:pointer;font-size:11px;font-weight:600;color:#666;">Este mes</button>
        <button onclick="setPeriodo(3)" style="padding:5px 10px;border:1.5px solid var(--borde);border-radius:6px;background:white;cursor:pointer;font-size:11px;font-weight:600;color:#666;">3 meses</button>
        <button onclick="setPeriodo(6)" style="padding:5px 10px;border:1.5px solid var(--borde);border-radius:6px;background:white;cursor:pointer;font-size:11px;font-weight:600;color:#666;">6 meses</button>
        <button onclick="setPeriodo(12)" style="padding:5px 10px;border:1.5px solid var(--borde);border-radius:6px;background:white;cursor:pointer;font-size:11px;font-weight:600;color:#666;">1 año</button>
        <button onclick="setPeriodo(0)" style="padding:5px 10px;border:1.5px solid var(--borde);border-radius:6px;background:white;cursor:pointer;font-size:11px;font-weight:600;color:#666;">Todo</button>
      </div>
      <div style="margin-left:auto;display:flex;gap:8px;">
        <button onclick="actualizarGraficos()" class="btn-primary" style="padding:7px 16px;">🔄 Actualizar</button>
        <button onclick="exportarExcel()" style="padding:7px 16px;background:#16a34a;color:white;border:none;border-radius:8px;cursor:pointer;font-size:13px;font-weight:600;">⬇️ Exportar CSV</button>
      </div>
    </div>

    <!-- Contenido scroll -->
    <div style="flex:1;overflow-y:auto;padding:16px 20px;display:flex;flex-direction:column;gap:16px;">

      <!-- KPIs -->
      <div id="graf-kpis" style="display:grid;grid-template-columns:repeat(6,1fr);gap:12px;"></div>

      <!-- Fila 1: doughnut + barras horizontales -->
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;">
        <div class="chart-card">
          <h3>📊 Distribución por Estado</h3>
          <div style="display:flex;align-items:center;gap:16px;">
            <div style="flex:1;max-height:220px;"><canvas id="chart-estado"></canvas></div>
            <div id="legend-estado" style="font-size:12px;display:flex;flex-direction:column;gap:6px;min-width:140px;"></div>
          </div>
        </div>
        <div class="chart-card">
          <h3>🏢 Solicitudes por Unidad</h3>
          <canvas id="chart-unidad" height="220"></canvas>
        </div>
      </div>

      <!-- Fila 2: línea de tendencia (ancho completo) -->
      <div class="chart-card">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
          <h3 style="margin:0;">📅 Evolución Mensual de Solicitudes</h3>
          <div style="display:flex;gap:8px;font-size:11px;">
            <span style="display:flex;align-items:center;gap:4px;"><span style="width:12px;height:3px;background:#1a3a6b;display:inline-block;border-radius:2px;"></span>Total ingresadas</span>
            <span style="display:flex;align-items:center;gap:4px;"><span style="width:12px;height:3px;background:#22c55e;display:inline-block;border-radius:2px;"></span>Cerradas</span>
          </div>
        </div>
        <canvas id="chart-mes" height="100"></canvas>
      </div>

      <!-- Fila 3: tabla resumen por unidad -->
      <div class="chart-card" style="padding:0;overflow:hidden;">
        <div style="padding:14px 18px;border-bottom:1px solid var(--borde);display:flex;justify-content:space-between;align-items:center;">
          <h3 style="margin:0;">🏢 Resumen por Unidad</h3>
          <span id="tabla-periodo" style="font-size:12px;color:#888;"></span>
        </div>
        <div style="overflow-x:auto;">
          <table id="tabla-unidades" style="width:100%;border-collapse:collapse;font-size:13px;">
            <thead>
              <tr style="background:#f8fafc;border-bottom:2px solid var(--borde);">
                <th style="padding:10px 16px;text-align:left;font-weight:700;color:#374151;">Unidad</th>
                <th style="padding:10px 12px;text-align:center;color:#3b82f6;">Derivadas</th>
                <th style="padding:10px 12px;text-align:center;color:#22c55e;">Respondidas</th>
                <th style="padding:10px 12px;text-align:center;color:#6b7280;">Cerradas</th>
                <th style="padding:10px 12px;text-align:center;color:#7e22ce;">Pend. Cierre</th>
                <th style="padding:10px 12px;text-align:center;color:#f59e0b;">En Proceso</th>
                <th style="padding:10px 12px;text-align:center;color:#374151;">Total</th>
                <th style="padding:10px 16px;text-align:left;color:#374151;">Efectividad</th>
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

async function actualizarGraficos() {
  showLoading("Generando reportes...");
  try {
    const desde = document.getElementById("graf-desde")?.value;
    const hasta = document.getElementById("graf-hasta")?.value;
    const all = await getSolicitudes();
    const filtradas = all.filter(s => {
      const f = new Date(s.FechaRecepcion);
      return (!desde || f >= new Date(desde)) && (!hasta || f <= new Date(hasta + "T23:59:59"));
    });

    const COLORES_ESTADO = {
      "Ingresada":           "#3b82f6",
      "Derivada":            "#f59e0b",
      "En Proceso":          "#06b6d4",
      "Respondida":          "#22c55e",
      "Devuelta":            "#ef4444",
      "Pendiente de Cierre": "#7e22ce",
      "Cerrada":             "#6b7280"
    };

    // ── KPIs ──
    const total         = filtradas.length;
    const cerradas      = filtradas.filter(s => s.Estado === "Cerrada").length;
    const respondidas   = filtradas.filter(s => s.Estado === "Respondida").length;
    const pendCierre    = filtradas.filter(s => s.Estado === "Pendiente de Cierre").length;
    const pendientes    = filtradas.filter(s => ["Ingresada","Derivada","En Proceso"].includes(s.Estado)).length;
    const devueltas     = filtradas.filter(s => s.Estado === "Devuelta").length;
    const tasa          = total > 0 ? Math.round(((cerradas + pendCierre) / total) * 100) : 0;

    const kpis = [
      { label:"Total período",      valor: total,       icon:"📋", color:"#1a3a6b", bg:"#eff6ff" },
      { label:"Cerradas",           valor: cerradas,    icon:"🔒", color:"#6b7280", bg:"#f3f4f6" },
      { label:"Respondidas",        valor: respondidas, icon:"✅", color:"#15803d", bg:"#f0fdf4" },
      { label:"Pend. de Cierre",    valor: pendCierre,  icon:"⏳", color:"#7e22ce", bg:"#fdf4ff" },
      { label:"Devueltas",          valor: devueltas,   icon:"↩️", color:"#b91c1c", bg:"#fff1f2" },
      { label:"Tasa resolución",    valor: tasa + "%",  icon:"📈", color:"#0e7490", bg:"#ecfeff" },
    ];
    const kpisCont = document.getElementById("graf-kpis");
    if (kpisCont) kpisCont.innerHTML = kpis.map(k => `
      <div style="background:${k.bg};border-radius:12px;padding:14px 16px;border:1.5px solid ${k.color}22;">
        <div style="font-size:22px;margin-bottom:4px;">${k.icon}</div>
        <div style="font-size:26px;font-weight:800;color:${k.color};line-height:1;">${k.valor}</div>
        <div style="font-size:11px;color:#6b7280;margin-top:4px;font-weight:600;">${k.label}</div>
      </div>`).join("");

    // ── Chart 1: Doughnut por estado ──
    const byEstado = {};
    filtradas.forEach(s => { byEstado[s.Estado] = (byEstado[s.Estado] || 0) + 1; });
    const estadoLabels = Object.keys(byEstado);
    const estadoData   = Object.values(byEstado);
    const estadoColors = estadoLabels.map(e => COLORES_ESTADO[e] || "#94a3b8");

    if (state.chartInstances["chart-estado"]) state.chartInstances["chart-estado"].destroy();
    const ctx1 = document.getElementById("chart-estado")?.getContext("2d");
    if (ctx1) {
      state.chartInstances["chart-estado"] = new Chart(ctx1, {
        type: "doughnut",
        data: { labels: estadoLabels, datasets: [{ data: estadoData, backgroundColor: estadoColors, borderWidth: 2, borderColor: "white" }] },
        options: { responsive:true, cutout:"65%", plugins:{ legend:{ display:false }, tooltip:{ callbacks:{ label: ctx => ` ${ctx.label}: ${ctx.raw} (${Math.round(ctx.raw/total*100)}%)` } } } }
      });
    }
    // Leyenda custom
    const legCont = document.getElementById("legend-estado");
    if (legCont) legCont.innerHTML = estadoLabels.map((e,i) => `
      <div style="display:flex;align-items:center;gap:6px;">
        <span style="width:12px;height:12px;border-radius:3px;background:${estadoColors[i]};flex-shrink:0;"></span>
        <span style="color:#374151;">${e}</span>
        <span style="margin-left:auto;font-weight:700;color:${estadoColors[i]};">${estadoData[i]}</span>
      </div>`).join("");

    // ── Chart 2: Barras horizontales por unidad ──
    const byUnidad = {};
    filtradas.filter(s => s.UnidadDerivada).forEach(s => {
      byUnidad[s.UnidadDerivada] = (byUnidad[s.UnidadDerivada] || 0) + 1;
    });
    const unidadLabels = Object.keys(byUnidad).sort((a,b) => byUnidad[b]-byUnidad[a]);
    const unidadData   = unidadLabels.map(u => byUnidad[u]);

    if (state.chartInstances["chart-unidad"]) state.chartInstances["chart-unidad"].destroy();
    const ctx2 = document.getElementById("chart-unidad")?.getContext("2d");
    if (ctx2) {
      state.chartInstances["chart-unidad"] = new Chart(ctx2, {
        type: "bar",
        data: { labels: unidadLabels, datasets: [{ data: unidadData, backgroundColor: "#1a3a6b", borderRadius: 6, borderSkipped: false }] },
        options: { indexAxis:"y", responsive:true, plugins:{ legend:{ display:false } }, scales:{ x:{ beginAtZero:true, ticks:{ stepSize:1 } }, y:{ ticks:{ font:{ size:11 } } } } }
      });
    }

    // ── Chart 3: Línea mensual (ingresadas vs cerradas) ──
    const byMesIng = {}, byMesCer = {};
    filtradas.forEach(s => {
      const m = s.FechaRecepcion?.substring(0, 7);
      if (m) byMesIng[m] = (byMesIng[m] || 0) + 1;
    });
    filtradas.filter(s => s.Estado === "Cerrada").forEach(s => {
      const m = s.FechaRecepcion?.substring(0, 7);
      if (m) byMesCer[m] = (byMesCer[m] || 0) + 1;
    });
    const meses = [...new Set([...Object.keys(byMesIng), ...Object.keys(byMesCer)])].sort();
    const mesesLabel = meses.map(m => {
      const [y, mo] = m.split("-");
      return `${["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"][parseInt(mo)-1]} ${y.slice(2)}`;
    });

    if (state.chartInstances["chart-mes"]) state.chartInstances["chart-mes"].destroy();
    const ctx3 = document.getElementById("chart-mes")?.getContext("2d");
    if (ctx3) {
      state.chartInstances["chart-mes"] = new Chart(ctx3, {
        type: "line",
        data: {
          labels: mesesLabel,
          datasets: [
            { label:"Ingresadas", data: meses.map(m => byMesIng[m]||0), borderColor:"#1a3a6b", backgroundColor:"rgba(26,58,107,0.08)", fill:true, tension:0.4, pointRadius:4, pointBackgroundColor:"#1a3a6b", borderWidth:2.5 },
            { label:"Cerradas",   data: meses.map(m => byMesCer[m]||0), borderColor:"#22c55e", backgroundColor:"rgba(34,197,94,0.06)",  fill:true, tension:0.4, pointRadius:4, pointBackgroundColor:"#22c55e", borderWidth:2.5, borderDash:[5,3] }
          ]
        },
        options: { responsive:true, plugins:{ legend:{ display:false } }, scales:{ y:{ beginAtZero:true, ticks:{ stepSize:1 } } } }
      });
    }

    // ── Tabla resumen por unidad ──
    const unidades = [...new Set(filtradas.filter(s=>s.UnidadDerivada).map(s=>s.UnidadDerivada))].sort();
    const tbody = document.getElementById("tabla-unidades-body");
    const per   = document.getElementById("tabla-periodo");
    if (per) per.textContent = desde && hasta ? `${desde} — ${hasta}` : "Todo el período";
    if (tbody) {
      tbody.innerHTML = unidades.map(u => {
        const sols = filtradas.filter(s => s.UnidadDerivada === u);
        const der  = sols.length;
        const resp = sols.filter(s => s.Estado === "Respondida").length;
        const cerr = sols.filter(s => s.Estado === "Cerrada").length;
        const penc = sols.filter(s => s.Estado === "Pendiente de Cierre").length;
        const proc = sols.filter(s => s.Estado === "En Proceso").length;
        const efe  = der > 0 ? Math.round(((resp + cerr + penc) / der) * 100) : 0;
        const efeColor = efe >= 80 ? "#15803d" : efe >= 50 ? "#b45309" : "#b91c1c";
        return `
          <tr style="border-bottom:1px solid var(--borde);" onmouseenter="this.style.background='#f8fafc'" onmouseleave="this.style.background=''">
            <td style="padding:10px 16px;font-weight:600;color:#1a3a6b;">${u}</td>
            <td style="padding:10px 12px;text-align:center;"><span style="background:#dbeafe;color:#1d4ed8;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700;">${der}</span></td>
            <td style="padding:10px 12px;text-align:center;"><span style="background:#dcfce7;color:#15803d;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700;">${resp}</span></td>
            <td style="padding:10px 12px;text-align:center;"><span style="background:#f3f4f6;color:#374151;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700;">${cerr}</span></td>
            <td style="padding:10px 12px;text-align:center;"><span style="background:#fdf4ff;color:#7e22ce;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700;">${penc}</span></td>
            <td style="padding:10px 12px;text-align:center;"><span style="background:#fef3c7;color:#b45309;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700;">${proc}</span></td>
            <td style="padding:10px 12px;text-align:center;font-weight:700;">${der}</td>
            <td style="padding:10px 16px;">
              <div style="display:flex;align-items:center;gap:8px;">
                <div style="flex:1;height:6px;background:#e2e8f0;border-radius:3px;overflow:hidden;">
                  <div style="height:100%;width:${efe}%;background:${efeColor};border-radius:3px;transition:width 0.5s;"></div>
                </div>
                <span style="font-size:12px;font-weight:700;color:${efeColor};min-width:36px;">${efe}%</span>
              </div>
            </td>
          </tr>`;
      }).join("") || `<tr><td colspan="7" style="padding:24px;text-align:center;color:#9ca3af;">Sin datos para el período seleccionado</td></tr>`;
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
