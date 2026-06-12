// ===== Cache en memoria con TTL =====
const _cache = {};
const _cacheTTL = {
  [CONFIG.lists?.usuarios]:     5 * 60 * 1000,  // 5 min — cambia poco
  [CONFIG.lists?.unidades]:    10 * 60 * 1000,  // 10 min — muy estable
  [CONFIG.lists?.solicitudes]:  1 * 60 * 1000,  //  1 min — cambia seguido
  [CONFIG.lists?.historial]:    2 * 60 * 1000,  //  2 min
  [CONFIG.lists?.evidencias]:   2 * 60 * 1000,  //  2 min
};
const TTL_DEFAULT = 60 * 1000; // 1 min para cualquier otra lista

function _cacheKey(listName, filter) { return `${listName}||${filter||""}`; }

function _cacheGet(listName, filter) {
  const key = _cacheKey(listName, filter);
  const entry = _cache[key];
  if (!entry) return null;
  if (Date.now() > entry.expires) { delete _cache[key]; return null; }
  return entry.data;
}

function _cacheSet(listName, filter, data) {
  const ttl = _cacheTTL[listName] ?? TTL_DEFAULT;
  _cache[_cacheKey(listName, filter)] = { data, expires: Date.now() + ttl };
}

function _cacheInvalidate(listName) {
  // Elimina todas las entradas de esa lista
  Object.keys(_cache).forEach(k => { if (k.startsWith(listName + "||")) delete _cache[k]; });
}

// ===== SharePoint REST API =====
const SP_BASE = `${CONFIG.sharePointSite}/_api`;

async function spFetch(url, options = {}) {
  const token = await getSharePointToken();
  const method = options.method || "GET";
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: "application/json;odata=nometadata",
    ...(options.body ? { "Content-Type": "application/json;odata=verbose" } : {}),
    ...(method === "PATCH" || method === "DELETE" ? { "IF-MATCH": "*", "X-HTTP-Method": method } : {}),
    ...(options.headers || {})
  };

  const fetchMethod = (method === "PATCH" || method === "DELETE") ? "POST" : method;

  const res = await fetch(url, {
    method: fetchMethod,
    headers,
    body: options.body
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`SharePoint API error ${res.status}: ${err}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

// ===== CRUD listas SharePoint =====

async function getListItems(listName, filter = "") {
  // Revisar cache primero
  const cached = _cacheGet(listName, filter);
  if (cached) return cached;

  let url = `${SP_BASE}/web/lists/getbytitle('${encodeURIComponent(listName)}')/items?$top=5000`;
  if (filter) url += `&$filter=${filter.replace(/ /g, '%20')}`;

  let items = [];
  while (url) {
    const data = await spFetch(url);
    items = items.concat(Array.isArray(data?.value) ? data.value : []);
    url = data?.["odata.nextLink"] || null;
  }
  const result = items.map(item => ({ id: String(item.Id || item.id || ""), ...item }));
  _cacheSet(listName, filter, result);
  return result;
}

async function createListItem(listName, fields) {
  const url = `${SP_BASE}/web/lists/getbytitle('${encodeURIComponent(listName)}')/items`;
  // SP REST necesita __metadata para POST
  const body = JSON.stringify({
    __metadata: { type: await getListItemType(listName) },
    ...fields
  });
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${await getSharePointToken()}`,
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=verbose"
    },
    body
  });
  if (!res.ok) throw new Error(`Error creando item: ${await res.text()}`);
  const data = await res.json();
  _cacheInvalidate(listName); // invalidar cache al escribir
  return { id: String(data.Id || data.id || ""), ...data };
}

async function updateListItem(listName, itemId, fields) {
  const url = `${SP_BASE}/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${itemId})`;
  const body = JSON.stringify({
    __metadata: { type: await getListItemType(listName) },
    ...fields
  });
  const token = await getSharePointToken();
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=verbose",
      "IF-MATCH": "*",
      "X-HTTP-Method": "MERGE"
    },
    body
  });
  if (!res.ok && res.status !== 204) throw new Error(`Error actualizando: ${await res.text()}`);
  _cacheInvalidate(listName); // invalidar cache al escribir
  return null;
}

// Cache de tipos de lista para evitar múltiples llamadas
const _listTypeCache = {};
async function getListItemType(listName) {
  if (_listTypeCache[listName]) return _listTypeCache[listName];
  const data = await spFetch(`${SP_BASE}/web/lists/getbytitle('${encodeURIComponent(listName)}')?$select=ListItemEntityTypeFullName`);
  _listTypeCache[listName] = data.ListItemEntityTypeFullName;
  return _listTypeCache[listName];
}

// ===== Adjuntos via SharePoint REST =====

async function getListItemAttachments(listName, itemId) {
  try {
    const data = await spFetch(`${SP_BASE}/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${itemId})/AttachmentFiles`);
    return (data.value || []).map(a => ({
      name: a.FileName,
      serverRelativeUrl: a.ServerRelativeUrl,
      url: a.ServerRelativeUrl,
      downloadUrl: `${CONFIG.sharePointSite}${a.ServerRelativeUrl}`
    }));
  } catch {
    return [];
  }
}

// Descarga un adjunto vía REST API SharePoint y retorna blob URL
const _blobCache = {};
async function getAttachmentBlobUrl(downloadUrl, serverRelativeUrl) {
  const cacheKey = serverRelativeUrl || downloadUrl;
  if (_blobCache[cacheKey]) return _blobCache[cacheKey];

  const token = await getSharePointToken();

  // Extraer path relativo sin el dominio
  const relUrl = serverRelativeUrl ||
    downloadUrl.replace("https://mdonihue.sharepoint.com", "");

  // IMPORTANTE: NO usar encodeURIComponent — SharePoint necesita el path sin codificar
  // Solo escapar comillas simples si las hay
  const safeRelUrl = relUrl.replace(/'/g, "''");
  const apiUrl = `${SP_BASE}/web/getfilebyserverrelativeurl('${safeRelUrl}')/$value`;

  const res = await fetch(apiUrl, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/octet-stream, */*"
    }
  });

  if (!res.ok) {
    // Intentar con la URL directa como fallback
    const res2 = await fetch(downloadUrl, {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (!res2.ok) throw new Error(`No se pudo descargar el archivo (${res.status})`);
    const blob2 = await res2.blob();
    const blobUrl2 = URL.createObjectURL(blob2);
    _blobCache[cacheKey] = blobUrl2;
    return blobUrl2;
  }

  const blob = await res.blob();
  const blobUrl = URL.createObjectURL(blob);
  _blobCache[cacheKey] = blobUrl;
  return blobUrl;
}

async function uploadAttachment(listName, itemId, file) {
  const token = await getSharePointToken();
  const url = `${SP_BASE}/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(file.name)}')`;
  const buf = await file.arrayBuffer();
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json;odata=nometadata"
    },
    body: buf
  });
  return res.json();
}

// ===== Email via Microsoft Graph =====

async function sendEmail(toEmail, subject, body) {
  const token = await getGraphToken();
  const mail = {
    message: {
      subject,
      body: { contentType: "HTML", content: body },
      toRecipients: [{ emailAddress: { address: toEmail } }]
    },
    saveToSentItems: false
  };
  const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(mail)
  });
  if (!res.ok && res.status !== 202 && res.status !== 204) {
    console.warn("Email error:", await res.text());
  }
}

// ===== Helpers de dominio =====

async function getUserByEmail(email) {
  const items = await getListItems(CONFIG.lists.usuarios, `Correo eq '${email}'`);
  return items[0] || null;
}

async function crearIndicesSharePoint() {
  const indices = [
    // Solicitud_Dom — filtros principales de la app
    { list: CONFIG.lists.solicitudes, field: "Estado" },
    { list: CONFIG.lists.solicitudes, field: "UnidadDerivada" },
    { list: CONFIG.lists.solicitudes, field: "NroSolicitud" },
    { list: CONFIG.lists.solicitudes, field: "FechaRecepcion" },
    { list: CONFIG.lists.solicitudes, field: "FechaDerivacion" },
    // HistorialSolicitud — consultado siempre por NroSolicitud
    { list: CONFIG.lists.historial,   field: "NroSolicitud" },
    // EvidenciaSolicitudes — consultado por NroSolicitud y SolicitudID
    { list: CONFIG.lists.evidencias,  field: "NroSolicitud" },
    { list: CONFIG.lists.evidencias,  field: "SolicitudID" },
  ];
  const token = await getSharePointToken();
  let ok = 0;
  for (const { list, field } of indices) {
    try {
      const url = `${SP_BASE}/web/lists/getbytitle('${encodeURIComponent(list)}')/fields/getbytitle('${encodeURIComponent(field)}')`;
      const res = await fetch(url, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json;odata=nometadata",
          "Content-Type": "application/json;odata=verbose",
          "IF-MATCH": "*",
          "X-HTTP-Method": "MERGE"
        },
        body: JSON.stringify({ __metadata: { type: "SP.Field" }, Indexed: true })
      });
      if (res.ok || res.status === 204) ok++;
      else console.warn(`[DOM] Índice ${list}.${field}: HTTP ${res.status}`);
    } catch(e) {
      console.warn(`[DOM] Índice ${list}.${field}:`, e.message);
    }
  }
  console.info(`[DOM] Índices SharePoint: ${ok}/${indices.length} aplicados`);
}

async function crearCampoFechaDerivacion() {
  const url = `${SP_BASE}/web/lists/getbytitle('${encodeURIComponent(CONFIG.lists.solicitudes)}')/fields`;
  try {
    await spFetch(url, {
      method: "POST",
      body: JSON.stringify({
        __metadata: { type: "SP.FieldDateTime" },
        Title: "FechaDerivacion",
        FieldTypeKind: 4,
        AddToDefaultView: true,
        Required: false
      })
    });
    console.info("[DOM] Campo FechaDerivacion creado en SharePoint.");
  } catch(e) {
    // 400 "duplicate" = ya existe → ignorar; cualquier otro error → solo advertencia
    if (!e.message.includes("400")) console.warn("[DOM] FechaDerivacion:", e.message);
  }
}

async function getConfiguracionDOM() {
  const items = await getListItems(CONFIG.lists.configuracion);
  const cfg = {};
  items.forEach(i => { if (i.Title && i.Valor !== undefined) cfg[i.Title] = i.Valor; });
  return cfg;
}

async function getSolicitudes(extraFilter = "") {
  return getListItems(CONFIG.lists.solicitudes, extraFilter);
}

async function crearSolicitud(fields) {
  return createListItem(CONFIG.lists.solicitudes, fields);
}

async function actualizarSolicitud(itemId, fields) {
  return updateListItem(CONFIG.lists.solicitudes, itemId, fields);
}

async function registrarHistorial(fields) {
  // Nunca debe bloquear el flujo principal — si falla, solo log
  try {
    // El campo en SharePoint se llama "Accion" (no Title)
    // Si viene Title lo mapeamos a Accion; SharePoint requiere Title también (campo obligatorio)
    const mapped = { ...fields };
    if (mapped.Title !== undefined && mapped.Accion === undefined) {
      mapped.Accion = mapped.Title;
    }
    // Title es campo obligatorio en SharePoint — usar Accion como valor
    if (!mapped.Title) mapped.Title = mapped.Accion || "Registro";
    return await createListItem(CONFIG.lists.historial, mapped);
  } catch(e) {
    console.warn("registrarHistorial (no crítico):", e.message);
    return null;
  }
}

async function getHistorialBySolicitud(nroSolicitud) {
  return getListItems(CONFIG.lists.historial, `NroSolicitud eq '${nroSolicitud}'`);
}

async function getEvidenciasBySolicitud(nroSolicitud, solicitudId = null) {
  const nro = String(nroSolicitud || "").trim();
  const idNum = solicitudId ? parseInt(solicitudId) : null;

  // Intentar filtro en servidor (requiere columnas indexadas en SharePoint)
  try {
    const partes = [];
    if (nro) partes.push(`NroSolicitud eq '${nro.replace(/'/g, "''")}'`);
    if (idNum) partes.push(`SolicitudID eq ${idNum}`);
    if (!partes.length) return [];
    const filter = partes.join(" or ");
    return await getListItems(CONFIG.lists.evidencias, filter);
  } catch(e) {
    // Si falla (ej: columna no indexada), fallback a filtro en cliente
    console.warn("Evidencias: filtro servidor falló, usando cliente:", e.message);
    try {
      const todas = await getListItems(CONFIG.lists.evidencias);
      return todas.filter(ev => {
        if (nro && String(ev.NroSolicitud || "").trim() === nro) return true;
        if (idNum && (parseInt(ev.SolicitudID) === idNum || parseInt(ev.SolicitudId) === idNum)) return true;
        return false;
      });
    } catch(e2) {
      console.warn("Error cargando evidencias:", e2.message);
      return [];
    }
  }
}

async function crearEvidencia(fields) {
  return createListItem(CONFIG.lists.evidencias, fields);
}

async function getDirectores() {
  const usuarios = await getListItems(CONFIG.lists.usuarios, `Rol eq 'Director' and Activo eq 1`);
  return usuarios;
}

async function getUsuariosByUnidad(unidad) {
  return getListItems(CONFIG.lists.usuarios, `Unidad eq '${unidad}' and Activo eq 1`);
}

// ===== Unidades (SharePoint list UnidadesDOM) =====
async function getUnidades() {
  try {
    const items = await getListItems(CONFIG.lists.unidades);
    return items.filter(u => u.Activo !== false && u.Activo !== 0);
  } catch(e) {
    console.warn("No se pudo cargar UnidadesDOM, usando lista predeterminada:", e.message);
    return CONFIG.unidades.map((n,i) => ({ id: String(i), Title: n }));
  }
}

async function crearUnidad(nombre) {
  return createListItem(CONFIG.lists.unidades, { Title: nombre, Activo: true });
}

async function actualizarUnidad(itemId, fields) {
  return updateListItem(CONFIG.lists.unidades, itemId, fields);
}

async function eliminarUnidad(itemId) {
  const url = `${SP_BASE}/web/lists/getbytitle('${encodeURIComponent(CONFIG.lists.unidades)}')/items(${itemId})`;
  await spFetch(url, { method: "DELETE" });
}

// ===== Usuarios (SharePoint list UsuarioDom) =====
async function getTodosUsuarios() {
  return getListItems(CONFIG.lists.usuarios);
}

async function crearUsuario(fields) {
  return createListItem(CONFIG.lists.usuarios, fields);
}

async function actualizarUsuario(itemId, fields) {
  return updateListItem(CONFIG.lists.usuarios, itemId, fields);
}

// ===== Notificaciones =====

function notificarDirector(solicitud, accion) {
  // Fire-and-forget: no bloquea el flujo principal
  setTimeout(async () => {
    try {
      const directores = await getDirectores().catch(() => []);
      await Promise.all(directores.map(dir =>
        sendEmail(
          dir.Correo,
          `[SistemaDOM] Solicitud ${solicitud.NroSolicitud} — ${accion}`,
          emailTemplate(solicitud, accion)
        ).catch(e => console.warn("Email director:", e.message))
      ));
    } catch(e) { console.warn("notificarDirector:", e.message); }
  }, 0);
}

function notificarUnidad(solicitud, unidad) {
  // Fire-and-forget: no bloquea el flujo principal
  setTimeout(async () => {
    try {
      const usuarios = await getUsuariosByUnidad(unidad).catch(() => []);
      await Promise.all(usuarios.map(u =>
        sendEmail(
          u.Correo,
          `[SistemaDOM] Solicitud derivada a ${unidad} — ${solicitud.NroSolicitud}`,
          emailTemplate(solicitud, `Derivada a ${unidad}`)
        ).catch(e => console.warn("Email unidad:", e.message))
      ));
    } catch(e) { console.warn("notificarUnidad:", e.message); }
  }, 0);
}

function emailTemplate(sol, accion) {
  return `
    <div style="font-family:Arial;max-width:600px;margin:0 auto;background:#f5f5f5;padding:20px;">
      <div style="background:#1a3a6b;color:white;padding:20px;border-radius:8px 8px 0 0;">
        <h2 style="margin:0;font-size:18px;">Municipalidad de Doñihue</h2>
        <p style="margin:4px 0 0;font-size:13px;opacity:0.85;">Dirección de Obras — Sistema de Solicitudes</p>
      </div>
      <div style="background:white;padding:24px;border-radius:0 0 8px 8px;border:1px solid #ddd;">
        <h3 style="color:#1a3a6b;margin-top:0;">${accion}</h3>
        <table style="width:100%;border-collapse:collapse;font-size:14px;">
          <tr style="border-bottom:1px solid #eee;"><td style="padding:8px;color:#666;width:140px;">Nro Solicitud</td><td style="padding:8px;font-weight:bold;">${sol.NroSolicitud||''}</td></tr>
          <tr style="border-bottom:1px solid #eee;background:#fafafa;"><td style="padding:8px;color:#666;">Solicitante</td><td style="padding:8px;">${sol.Solicitante||''}</td></tr>
          <tr style="border-bottom:1px solid #eee;"><td style="padding:8px;color:#666;">Dirección</td><td style="padding:8px;">${sol.Direccion||''}</td></tr>
          <tr style="border-bottom:1px solid #eee;background:#fafafa;"><td style="padding:8px;color:#666;">Solicitud</td><td style="padding:8px;">${sol.Solicitud||''}</td></tr>
          <tr><td style="padding:8px;color:#666;">Estado</td><td style="padding:8px;"><span style="background:#1a3a6b;color:white;padding:3px 10px;border-radius:12px;font-size:13px;">${sol.Estado||''}</span></td></tr>
        </table>
        <div style="margin-top:20px;text-align:center;">
          <a href="${CONFIG.redirectUri}" style="background:#1a3a6b;color:white;padding:12px 28px;text-decoration:none;border-radius:6px;font-size:14px;font-weight:600;">Ingresar al Sistema</a>
        </div>
      </div>
    </div>`;
}
