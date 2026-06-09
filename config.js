const CONFIG = {
  clientId: "c1f8a6c3-95bd-42c6-a7fd-0be2958f35ae",
  tenantId: "1b037852-4fe3-470f-bb9e-c8ba670e4653",
  redirectUri: "https://domdonihue.github.io/Solicitudes-/",
  sharePointSite: "https://mdonihue.sharepoint.com/sites/DOM",
  sitePath: "/sites/DOM",
  lists: {
    usuarios: "UsuarioDom",
    solicitudes: "Solicitud_Dom",
    historial: "HistorialSolicitud",
    evidencias: "EvidenciaSolicitudes"
  },
  estados: {
    INGRESADA: "Ingresada",
    DERIVADA: "Derivada",
    EN_PROCESO: "En Proceso",
    RESPONDIDA: "Respondida",
    DEVUELTA: "Devuelta",
    PENDIENTE_CIERRE: "Pendiente de Cierre",
    CERRADA: "Cerrada"
  },
  unidades: ["Inspección", "Operaciones", "Administración", "Aseo y Ornato", "Dirección de Obras"],
  roles: {
    DIRECTOR: "Director",
    SECRETARIA: "Secretaria",
    UNIDAD: "Unidad"
  }
};
