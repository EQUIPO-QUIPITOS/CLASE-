// Backend del MVP: recibe los formularios y los guarda en la hoja

const HOJAS = {
  registrar_emprendimiento: { hoja: "Emprendimientos", prefijo: "E" },
  registrar_creador:        { hoja: "Creadores",        prefijo: "C" },
};

// Recibe los formularios de la página (POST)
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const config = HOJAS[body.action];
    if (!config) return responder({ ok: false, error: "Acción no válida" });

    // Campo trampa: si un robot lo llenó, se ignora en silencio
    if (body.hp) return responder({ ok: true });

    const data = body.data || {};
    if (!/^\S+@\S+\.\S+$/.test(data.correo || "")) {
      return responder({ ok: false, error: "Correo no válido" });
    }

    guardarFila(config, data);
    avisar(config, data);
    return responder({ ok: true });
  } catch (err) {
    return responder({ ok: false, error: String(err) });
  }
}

// Sirve para comprobar que el servicio está activo (GET)
function doGet(e) {
  return responder({ ok: true, mensaje: "El servicio está activo" });
}

// Escribe una fila nueva usando los títulos de la fila 1 como guía
function guardarFila(config, data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // evita que dos envíos simultáneos se pisen
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.hoja);
    if (!hoja) throw new Error("No existe la pestaña " + config.hoja);

    const titulos = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    const auto = {
      id: config.prefijo + String(hoja.getLastRow()).padStart(3, "0"),
      fecha: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
      estado: "pendiente",
      verificado: "no",
    };

    const fila = titulos.map((t) => {
      const clave = String(t).trim();
      return clave in auto ? auto[clave] : limpiar(data[clave]);
    });

    // Formato texto para que "+57 300..." no se convierta en fórmula
    const destino = hoja.getRange(hoja.getLastRow() + 1, 1, 1, fila.length);
    destino.setNumberFormat("@");
    destino.setValues([fila]);
  } finally {
    lock.releaseLock();
  }
}

function limpiar(valor) {
  return valor === undefined || valor === null ? "" : String(valor).trim().slice(0, 300);
}

// Te manda un correo cada vez que llega un registro nuevo
function avisar(config, data) {
  try {
    MailApp.sendEmail(
      Session.getEffectiveUser().getEmail(),
      "Nuevo registro en " + config.hoja,
      JSON.stringify(data, null, 2)
    );
  } catch (err) {
    // Si el correo falla, el registro ya quedó guardado
  }
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// PRUEBA del formulario de emprendedores
function probar() {
  const e = { postData: { contents: JSON.stringify({
    action: "registrar_emprendimiento",
    hp: "",
    data: {
      negocio: "Negocio de prueba", sector: "Restaurantes y comida", ciudad: "Santa Marta",
      presupuesto: "200.000 a 500.000 COP", objetivo: "Conseguir nuevos clientes",
      tiempo_disponible: "2026-12-01", url_destino: "@negocioprueba",
      whatsapp: "+57 300 000 0000", correo: "prueba@correo.com",
    },
  }) } };
  Logger.log(doPost(e).getContent());
}

// PRUEBA del formulario de creadores
function probarCreador() {
  const e = { postData: { contents: JSON.stringify({
    action: "registrar_creador",
    hp: "",
    data: {
      nombre: "Creador de prueba", red: "Instagram", usuario_red: "@prueba",
      seguidores: "5000", nicho: "Comida y gastronomía", ciudad: "Santa Marta",
      tarifa: "50000", whatsapp: "+57 300 000 0000", correo: "prueba@correo.com",
    },
  }) } };
  Logger.log(doPost(e).getContent());
}
