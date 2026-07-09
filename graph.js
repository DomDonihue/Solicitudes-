// ===== Cache en memoria con TTL =====
const _cache = {};
const _cacheTTL = {
  [CONFIG.lists?.usuarios]:     5 * 60 * 1000,  // 5 min \u2014 cambia poco
  [CONFIG.lists?.unidades]:    10 * 60 * 1000,  // 10 min \u2014 muy estable
  [CONFIG.lists?.solicitudes]:  1 * 60 * 1000,  //  1 min \u2014 cambia seguido
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

// Cache de tipos de lista para evitar m\u00FAltiples llamadas
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

// Descarga un adjunto v\u00EDa REST API SharePoint y retorna blob URL
const _blobCache = {};
async function getAttachmentBlobUrl(downloadUrl, serverRelativeUrl) {
  const cacheKey = serverRelativeUrl || downloadUrl;
  if (_blobCache[cacheKey]) return _blobCache[cacheKey];

  const token = await getSharePointToken();

  // Extraer path relativo sin el dominio
  const relUrl = serverRelativeUrl ||
    downloadUrl.replace("https://mdonihue.sharepoint.com", "");

  // IMPORTANTE: NO usar encodeURIComponent \u2014 SharePoint necesita el path sin codificar
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

function comprimirImagen(file, maxPx = 1200, quality = 0.82) {
  const comprimibles = ["image/jpeg", "image/jpg", "image/png", "image/webp"];
  if (!comprimibles.includes(file.type)) return Promise.resolve(file);

  return new Promise(resolve => {
    const img = new Image();
    const url = URL.createObjectURL(file);
    img.onload = () => {
      URL.revokeObjectURL(url);
      // Si ya es peque\u00F1a, no tocar
      if (img.width <= maxPx && img.height <= maxPx && file.size < 300 * 1024) {
        resolve(file); return;
      }
      let w = img.width, h = img.height;
      if (w > h) { if (w > maxPx) { h = Math.round(h * maxPx / w); w = maxPx; } }
      else        { if (h > maxPx) { w = Math.round(w * maxPx / h); h = maxPx; } }

      const canvas = document.createElement("canvas");
      canvas.width = w; canvas.height = h;
      canvas.getContext("2d").drawImage(img, 0, 0, w, h);

      const outType = file.type === "image/png" ? "image/png" : "image/jpeg";
      const ext     = file.type === "image/png" ? "png" : "jpg";
      canvas.toBlob(blob => {
        if (!blob || blob.size >= file.size) { resolve(file); return; }
        const nombre = file.name.replace(/\.[^.]+$/, "") + "." + ext;
        console.info(`[DOM] Compresi\u00F3n: ${(file.size/1024).toFixed(0)}KB \u2192 ${(blob.size/1024).toFixed(0)}KB (${nombre})`);
        resolve(new File([blob], nombre, { type: outType }));
      }, outType, quality);
    };
    img.onerror = () => { URL.revokeObjectURL(url); resolve(file); };
    img.src = url;
  });
}

async function uploadAttachment(listName, itemId, file) {
  const compressed = await comprimirImagen(file);
  const token = await getSharePointToken();
  const url = `${SP_BASE}/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(compressed.name)}')`;
  const buf = await compressed.arrayBuffer();
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
    // Solicitud_Dom \u2014 filtros principales de la app
    { list: CONFIG.lists.solicitudes, field: "Estado" },
    { list: CONFIG.lists.solicitudes, field: "UnidadDerivada" },
    { list: CONFIG.lists.solicitudes, field: "NroSolicitud" },
    { list: CONFIG.lists.solicitudes, field: "FechaRecepcion" },
    { list: CONFIG.lists.solicitudes, field: "FechaDerivacion" },
    // HistorialSolicitud \u2014 consultado siempre por NroSolicitud
    { list: CONFIG.lists.historial,   field: "NroSolicitud" },
    // EvidenciaSolicitudes \u2014 consultado por NroSolicitud y SolicitudID
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
      else console.warn(`[DOM] \u00CDndice ${list}.${field}: HTTP ${res.status}`);
    } catch(e) {
      console.warn(`[DOM] \u00CDndice ${list}.${field}:`, e.message);
    }
  }
  console.info(`[DOM] \u00CDndices SharePoint: ${ok}/${indices.length} aplicados`);
}


async function _crearCamposLista(listName, campos) {
  const listUrl = `${SP_BASE}/web/lists/getbytitle('${encodeURIComponent(listName)}')/fields`;
  for (const campo of campos) {
    try {
      await spFetch(listUrl, {
        method: "POST",
        body: JSON.stringify({ ...campo, AddToDefaultView: true, Required: false })
      });
      console.info(`[DOM] Campo ${campo.Title} creado en ${listName}.`);
    } catch(e) {
      if (!e.message.includes("400")) console.warn(`[DOM] ${listName}.${campo.Title}:`, e.message);
    }
  }
}

async function crearCamposEvidencia() {
  await _crearCamposLista(CONFIG.lists.evidencias, [
    { Title: "NroSolicitud",         FieldTypeKind: 2, __metadata: { type: "SP.FieldText" } },
    { Title: "SolicitudID",          FieldTypeKind: 9, __metadata: { type: "SP.FieldNumber" } },
    { Title: "Unidad",               FieldTypeKind: 2, __metadata: { type: "SP.FieldText" } },
    { Title: "DescripcionEvidencia", FieldTypeKind: 3, __metadata: { type: "SP.FieldMultiLineText" } },
    { Title: "FechaCarga",           FieldTypeKind: 4, __metadata: { type: "SP.FieldDateTime" } },
    { Title: "Responsable",          FieldTypeKind: 2, __metadata: { type: "SP.FieldText" } },
  ]);
}

async function crearCamposHistorial() {
  await _crearCamposLista(CONFIG.lists.historial, [
    { Title: "NroSolicitud",   FieldTypeKind: 2, __metadata: { type: "SP.FieldText" } },
    { Title: "Accion",         FieldTypeKind: 2, __metadata: { type: "SP.FieldText" } },
    { Title: "EstadoAnterior", FieldTypeKind: 2, __metadata: { type: "SP.FieldText" } },
    { Title: "EstadoNuevo",    FieldTypeKind: 2, __metadata: { type: "SP.FieldText" } },
    { Title: "UsuarioAccion",  FieldTypeKind: 2, __metadata: { type: "SP.FieldText" } },
    { Title: "RolUsuario",     FieldTypeKind: 2, __metadata: { type: "SP.FieldText" } },
    { Title: "Unidad",         FieldTypeKind: 2, __metadata: { type: "SP.FieldText" } },
    { Title: "FechaAccion",    FieldTypeKind: 4, __metadata: { type: "SP.FieldDateTime" } },
    { Title: "Observaciones",  FieldTypeKind: 3, __metadata: { type: "SP.FieldMultiLineText" } },
  ]);
}

async function crearCamposConfiguracion() {
  // ValorMultilinea permite almacenar contenido extenso (imágenes base64, JSON)
  await _crearCamposLista(CONFIG.lists.configuracion, [
    { Title: "ValorMultilinea", FieldTypeKind: 3, __metadata: { type: "SP.FieldMultiLineText" } },
  ]);
}

// Claves que usan el campo multi-línea (contenido extenso, ej: imágenes base64)
const _CFG_MULTILINEA = new Set(["FirmaDirector"]);

async function getConfiguracionDOM() {
  const items = await getListItems(CONFIG.lists.configuracion);
  const cfg = {};
  items.forEach(i => {
    if (!i.Title) return;
    // Prefiere ValorMultilinea si la clave lo requiere, cae a Valor si no hay
    const v = _CFG_MULTILINEA.has(i.Title)
      ? (i.ValorMultilinea || i.Valor || "")
      : (i.Valor !== undefined ? i.Valor : (i.ValorMultilinea || ""));
    cfg[i.Title] = v;
  });
  return cfg;
}

async function actualizarConfiguracionDOM(key, valor) {
  _cacheInvalidate(CONFIG.lists.configuracion);
  const items = await getListItems(CONFIG.lists.configuracion);
  const existing = items.find(i => i.Title === key);
  const esMultilinea = _CFG_MULTILINEA.has(key);
  const fields = esMultilinea
    ? { ValorMultilinea: String(valor) }
    : { Valor: String(valor) };
  if (existing) {
    await updateListItem(CONFIG.lists.configuracion, existing.id, fields);
  } else {
    await createListItem(CONFIG.lists.configuracion, { Title: key, ...fields });
  }
  _cacheInvalidate(CONFIG.lists.configuracion);
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
  const mapped = { ...fields };
  if (mapped.Title !== undefined && mapped.Accion === undefined) mapped.Accion = mapped.Title;
  if (!mapped.Title) mapped.Title = mapped.Accion || "Registro";
  try {
    return await createListItem(CONFIG.lists.historial, mapped);
  } catch(e) {
    if (e.message.includes("InvalidClientQueryException") || e.message.includes("no es tipo") || e.message.includes("is not of type")) {
      console.warn("[DOM] Columnas faltantes en HistorialSolicitud \u2014 creando y reintentando...");
      await crearCamposHistorial();
      try { return await createListItem(CONFIG.lists.historial, mapped); } catch(e2) { console.warn("registrarHistorial:", e2.message); }
    } else {
      console.warn("registrarHistorial (no cr\u00EDtico):", e.message);
    }
    return null;
  }
}

async function getHistorialBySolicitud(nroSolicitud) {
  return getListItems(CONFIG.lists.historial, `NroSolicitud eq '${nroSolicitud}'`);
}

async function getEvidenciasBySolicitud(nroSolicitud, solicitudId = null) {
  const nro = String(nroSolicitud || "").trim();
  const idNum = solicitudId ? parseInt(solicitudId) : null;

  // Filtrar por SolicitudID (n\u00FAmero, siempre confiable) y como fallback por Title (= NroSolicitud)
  try {
    const partes = [];
    if (idNum) partes.push(`SolicitudID eq ${idNum}`);
    if (nro)   partes.push(`Title eq '${nro.replace(/'/g, "''")}'`);
    if (!partes.length) return [];
    const filter = partes.join(" or ");
    return await getListItems(CONFIG.lists.evidencias, filter);
  } catch(e) {
    // Si falla (ej: columna no indexada), fallback a filtro en cliente
    console.warn("Evidencias: filtro servidor fall\u00F3, usando cliente:", e.message);
    try {
      const todas = await getListItems(CONFIG.lists.evidencias);
      return todas.filter(ev => {
        if (idNum && (parseInt(ev.SolicitudID) === idNum || parseInt(ev.SolicitudId) === idNum)) return true;
        if (nro && (String(ev.Title || "").trim() === nro || String(ev.NroSolicitud || "").trim() === nro)) return true;
        return false;
      });
    } catch(e2) {
      console.warn("Error cargando evidencias:", e2.message);
      return [];
    }
  }
}

async function crearEvidencia(fields) {
  try {
    return await createListItem(CONFIG.lists.evidencias, fields);
  } catch(e) {
    const esFaltaColumna = e.message.includes("InvalidClientQueryException") || e.message.includes("no es tipo") || e.message.includes("is not of type");
    if (!esFaltaColumna) throw e;

    // Intentar crear columnas (solo funciona si el usuario tiene permisos admin)
    await crearCamposEvidencia().catch(() => {});

    // Reintentar con campos completos
    try {
      return await createListItem(CONFIG.lists.evidencias, fields);
    } catch(e2) {
      // Fallback: guardar solo Title \u2014 el adjunto se sube igual como attachment
      console.warn("[DOM] Evidencia sin columnas extra (permisos insuficientes). El administrador debe iniciar sesi\u00F3n para crear las columnas.", e2.message);
      return await createListItem(CONFIG.lists.evidencias, {
        Title: fields.Title || `Evidencia ${fields.SolicitudID || "?"}`
      });
    }
  }
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

async function deleteListItem(listName, itemId) {
  const url = `${SP_BASE}/web/lists/getbytitle('${encodeURIComponent(listName)}')/items(${itemId})`;
  await spFetch(url, { method: "DELETE" });
}

async function limpiarHistorialSolicitud(nroSolicitud) {
  const items = await getHistorialBySolicitud(nroSolicitud);
  if (!items.length) return;
  await Promise.all(items.map(i => deleteListItem(CONFIG.lists.historial, i.id).catch(e => console.warn("Del historial:", e.message))));
  _cacheInvalidate(CONFIG.lists.historial);
}

async function limpiarEvidenciasSolicitud(nroSolicitud, solicitudId) {
  const items = await getEvidenciasBySolicitud(nroSolicitud, solicitudId);
  if (!items.length) return;
  await Promise.all(items.map(i => deleteListItem(CONFIG.lists.evidencias, i.id).catch(e => console.warn("Del evidencia:", e.message))));
  _cacheInvalidate(CONFIG.lists.evidencias);
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
          `[SistemaDOM] Solicitud ${solicitud.NroSolicitud} \u2014 ${accion}`,
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
          `[SistemaDOM] Solicitud derivada a ${unidad} \u2014 ${solicitud.NroSolicitud}`,
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
        <h2 style="margin:0;font-size:18px;">Municipalidad de Do\u00F1ihue</h2>
        <p style="margin:4px 0 0;font-size:13px;opacity:0.85;">Direcci\u00F3n de Obras \u2014 Sistema de Solicitudes</p>
      </div>
      <div style="background:white;padding:24px;border-radius:0 0 8px 8px;border:1px solid #ddd;">
        <h3 style="color:#1a3a6b;margin-top:0;">${accion}</h3>
        <table style="width:100%;border-collapse:collapse;font-size:14px;">
          <tr style="border-bottom:1px solid #eee;"><td style="padding:8px;color:#666;width:140px;">Nro Solicitud</td><td style="padding:8px;font-weight:bold;">${sol.NroSolicitud||''}</td></tr>
          <tr style="border-bottom:1px solid #eee;background:#fafafa;"><td style="padding:8px;color:#666;">Solicitante</td><td style="padding:8px;">${sol.Solicitante||''}</td></tr>
          <tr style="border-bottom:1px solid #eee;"><td style="padding:8px;color:#666;">Direcci\u00F3n</td><td style="padding:8px;">${sol.Direccion||''}</td></tr>
          <tr style="border-bottom:1px solid #eee;background:#fafafa;"><td style="padding:8px;color:#666;">Solicitud</td><td style="padding:8px;">${sol.Solicitud||''}</td></tr>
          <tr><td style="padding:8px;color:#666;">Estado</td><td style="padding:8px;"><span style="background:#1a3a6b;color:white;padding:3px 10px;border-radius:12px;font-size:13px;">${sol.Estado||''}</span></td></tr>
        </table>
        <div style="margin-top:20px;text-align:center;">
          <a href="${CONFIG.redirectUri}" style="background:#1a3a6b;color:white;padding:12px 28px;text-decoration:none;border-radius:6px;font-size:14px;font-weight:600;">Ingresar al Sistema</a>
        </div>
      </div>
    </div>`;
}
