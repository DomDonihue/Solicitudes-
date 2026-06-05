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
  let url = `${SP_BASE}/web/lists/getbytitle('${encodeURIComponent(listName)}')/items?$top=5000`;
  if (filter) url += `&$filter=${encodeURIComponent(filter)}`;

  let items = [];
  while (url) {
    const data = await spFetch(url);
    items = items.concat(data.value);
    url = data["odata.nextLink"] || null;
  }
  // Normalizar: SharePoint devuelve Id (mayúscula), agregamos id minúscula para compatibilidad
  return items.map(item => ({ id: String(item.Id || item.id || ""), ...item }));
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
      url: a.ServerRelativeUrl,
      downloadUrl: `${CONFIG.sharePointSite}${a.ServerRelativeUrl}`
    }));
  } catch {
    return [];
  }
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
  return createListItem(CONFIG.lists.historial, fields);
}

async function getHistorialBySolicitud(nroSolicitud) {
  return getListItems(CONFIG.lists.historial, `NroSolicitud eq '${nroSolicitud}'`);
}

async function getEvidenciasBySolicitud(nroSolicitud) {
  return getListItems(CONFIG.lists.evidencias, `NroSolicitud eq '${nroSolicitud}'`);
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

// ===== Notificaciones =====

async function notificarDirector(solicitud, accion) {
  const directores = await getDirectores().catch(() => []);
  for (const dir of directores) {
    const asunto = `[SistemaDOM] Solicitud ${solicitud.NroSolicitud} — ${accion}`;
    const cuerpo = emailTemplate(solicitud, accion);
    await sendEmail(dir.Correo, asunto, cuerpo).catch(console.error);
  }
}

async function notificarUnidad(solicitud, unidad) {
  const usuarios = await getUsuariosByUnidad(unidad).catch(() => []);
  for (const u of usuarios) {
    const asunto = `[SistemaDOM] Solicitud derivada a ${unidad} — ${solicitud.NroSolicitud}`;
    const cuerpo = emailTemplate(solicitud, `Derivada a ${unidad}`);
    await sendEmail(u.Correo, asunto, cuerpo).catch(console.error);
  }
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
