// Backend del MVP: recibe los formularios, los guarda en la hoja
// y le manda un correo de bienvenida a quien se registró

const HOJAS = {
  registrar_emprendimiento: { hoja: "Emprendimientos", prefijo: "E", tipo: "emprendedor" },
  registrar_creador:        { hoja: "Creadores",        prefijo: "C", tipo: "creador" },
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
    avisar(config, data);          // correo para ti
    darBienvenida(config, data);   // correo para quien se registró
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

/* =====================================================================
   CORREO DE BIENVENIDA (con los colores y letras de la página)
   ===================================================================== */
function darBienvenida(config, data) {
  try {
    const esCreador = config.tipo === "creador";
    const nombre = limpiar(esCreador ? data.nombre : data.negocio) || "";
    const asunto = esCreador
      ? "¡Bienvenido a Hand in Hand! Recibimos tu postulación"
      : "¡Bienvenido a Hand in Hand! Tu negocio quedó registrado";

    MailApp.sendEmail({
      to: data.correo,
      subject: asunto,
      name: "Hand in Hand",
      htmlBody: plantillaBienvenida(esCreador, nombre, data),
      body: textoPlano(esCreador, nombre), // para correos que no muestran diseño
    });
  } catch (err) {
    // Si el correo falla, el registro ya quedó guardado
  }
}

function plantillaBienvenida(esCreador, nombre, data) {
  // Colores de la página
  const C = {
    fondo: "#E3E4EA", tarjeta: "#F6F7FB", oscuro: "#00171F", teal: "#005260",
    verde: "#70A312", lima: "#B5EA3C", cian: "#00E4E8", menta: "#C5DFDF", texto: "#1d2b30", gris: "#4f6a70",
  };
  const portada = esCreador ? "linear-gradient(180deg,#00485C,#001C26)" : "linear-gradient(180deg,#00513F,#002B26)";
  const portadaPlano = esCreador ? "#003646" : "#003D33";
  const acento = esCreador ? "#48CAE4" : C.lima;
  const etiqueta = esCreador ? "CREADORES" : "EMPRENDEDORES";
  const saludo = nombre ? "¡Hola, " + esc(nombre) + "!" : "¡Hola!";

  const intro = esCreador
    ? "Gracias por postularte como creador y por confiar en nosotros. Nos encanta que quieras crecer de la mano de los emprendimientos del Magdalena."
    : "Gracias por registrar tu negocio y por confiar en nosotros. Estamos felices de darte una mano para llegar a más personas.";

  const pasos = esCreador
    ? [
        ["01", "Revisamos tu perfil", "Verificamos tu cuenta para garantizar confianza a los negocios."],
        ["02", "Te proponemos campañas", "Te escribimos por WhatsApp con emprendimientos afines a tu contenido."],
        ["03", "Creas y mides", "Recibes tu código y enlace únicos para mostrar los resultados que generas."],
      ]
    : [
        ["01", "Revisamos tu registro", "Analizamos tu negocio y el objetivo de tu campaña."],
        ["02", "Te emparejamos", "Elegimos creadores afines a tu público y coordinamos contigo por WhatsApp."],
        ["03", "Ves tus resultados", "Cada creador tiene su propio código y enlace para que veas cuántos clientes trajo."],
      ];

  const resumen = esCreador
    ? [["Red", data.red], ["Usuario", data.usuario_red], ["Nicho", data.nicho], ["Municipio", data.ciudad]]
    : [["Negocio", data.negocio], ["Sector", data.sector], ["Objetivo", data.objetivo], ["Municipio", data.ciudad]];

  const fuente = "'Inter',Helvetica,Arial,sans-serif";
  const cursiva = "'Dancing Script','Brush Script MT','Segoe Script',cursive";

  const filasPasos = pasos.map((p) =>
    '<tr><td style="padding:0 0 18px 0;vertical-align:top;width:58px">' +
      '<div style="width:46px;height:46px;border-radius:50%;background:#9BDCD6;color:#003A40;font:700 17px ' + fuente + ';line-height:46px;text-align:center">' + p[0] + '</div></td>' +
    '<td style="padding:0 0 18px 0;vertical-align:top">' +
      '<div style="font:700 17px ' + fuente + ';color:' + C.teal + '">' + p[1] + '</div>' +
      '<div style="font:400 15px/1.5 ' + fuente + ';color:' + C.gris + ';margin-top:2px">' + p[2] + '</div></td></tr>'
  ).join("");

  const filasResumen = resumen.filter((r) => r[1]).map((r) =>
    '<tr><td style="padding:6px 0;font:600 14px ' + fuente + ';color:' + C.gris + ';width:110px">' + r[0] + '</td>' +
    '<td style="padding:6px 0;font:600 15px ' + fuente + ';color:' + C.texto + '">' + esc(r[1]) + '</td></tr>'
  ).join("");

  return '' +
  '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
  '<link href="https://fonts.googleapis.com/css2?family=Dancing+Script:wght@700&family=Inter:wght@400;600;700;900&display=swap" rel="stylesheet"></head>' +
  '<body style="margin:0;padding:0;background:' + C.fondo + '">' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + C.fondo + '"><tr><td align="center" style="padding:28px 12px">' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px">' +

    // Portada
    '<tr><td style="background:' + portadaPlano + ';background-image:' + portada + ';border-radius:24px 24px 0 0;padding:36px 32px 34px">' +
      '<div style="font:800 15px ' + fuente + ';letter-spacing:2px;color:#C8F5DC">' + etiqueta + '</div>' +
      '<div style="font:900 46px/1.05 ' + fuente + ';letter-spacing:-2px;color:' + C.fondo + ';margin-top:6px">Hand<span style="font:700 40px ' + cursiva + ';letter-spacing:0;color:' + acento + '">in</span>Hand</div>' +
      '<div style="font:400 17px ' + fuente + ';color:' + C.fondo + ';margin-top:10px">Estamos aquí para darte una mano</div>' +
    '</td></tr>' +

    // Cuerpo
    '<tr><td style="background:' + C.tarjeta + ';padding:34px 32px 10px">' +
      '<div style="font:700 34px/1.15 ' + cursiva + ';color:' + C.verde + '">' + saludo + '</div>' +
      '<p style="font:400 16px/1.6 ' + fuente + ';color:' + C.texto + ';margin:14px 0 0">' + intro + '</p>' +
      '<p style="font:400 16px/1.6 ' + fuente + ';color:' + C.texto + ';margin:12px 0 0">Tu registro ya está con nosotros. Esto es lo que sigue:</p>' +
      '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="margin-top:22px">' + filasPasos + '</table>' +
    '</td></tr>' +

    // Resumen
    (filasResumen
      ? '<tr><td style="background:' + C.tarjeta + ';padding:0 32px 30px">' +
          '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + C.menta + ';border-radius:16px"><tr><td style="padding:18px 22px">' +
          '<div style="font:700 26px ' + cursiva + ';color:' + C.teal + ';margin-bottom:6px">Tus datos</div>' +
          '<table role="presentation" cellpadding="0" cellspacing="0" width="100%">' + filasResumen + '</table>' +
          '</td></tr></table></td></tr>'
      : '') +

    // Cierre
    '<tr><td style="background:' + C.teal + ';padding:28px 32px;border-radius:0 0 24px 24px;text-align:center">' +
      '<div style="font:700 30px ' + cursiva + ';color:' + C.cian + '">¡Gracias por confiar en nosotros!</div>' +
      '<div style="font:400 15px/1.6 ' + fuente + ';color:#d9ecec;margin-top:8px">Muy pronto te escribiremos por WhatsApp.<br>¿Tienes dudas? Responde este correo o escríbenos a ' +
      '<a href="mailto:equipoquipitosoficial@gmail.com" style="color:' + C.lima + ';font-weight:600;text-decoration:none">equipoquipitosoficial@gmail.com</a></div>' +
    '</td></tr>' +

    '<tr><td style="padding:18px 10px;text-align:center;font:400 12px/1.5 ' + fuente + ';color:' + C.gris + '">' +
      'Recibes este correo porque te registraste en Hand in Hand. Usamos tus datos solo para gestionar campañas y no los compartimos con terceros.' +
    '</td></tr>' +

  '</table></td></tr></table></body></html>';
}

function textoPlano(esCreador, nombre) {
  return (nombre ? "¡Hola, " + nombre + "!" : "¡Hola!") + "\n\n" +
    (esCreador
      ? "Gracias por postularte como creador en Hand in Hand y por confiar en nosotros."
      : "Gracias por registrar tu negocio en Hand in Hand y por confiar en nosotros.") +
    "\n\nTu registro ya está con nosotros. Muy pronto te escribiremos por WhatsApp." +
    "\n\n¿Dudas? Escríbenos a equipoquipitosoficial@gmail.com\n\nHand in Hand: estamos aquí para darte una mano.";
}

// Evita que un texto escrito en el formulario rompa el diseño del correo
function esc(t) {
  return String(t).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =====================================================================
   PRUEBAS (se usan solo desde el editor, con el botón Ejecutar)
   ===================================================================== */

// Te manda a TI los dos correos de bienvenida para que veas cómo se ven
function probarCorreos() {
  const yo = Session.getEffectiveUser().getEmail();
  darBienvenida(HOJAS.registrar_emprendimiento, {
    negocio: "Panadería La Bahía", sector: "Restaurantes y comida", objetivo: "Conseguir nuevos clientes",
    ciudad: "Santa Marta", correo: yo,
  });
  darBienvenida(HOJAS.registrar_creador, {
    nombre: "Ana", red: "Instagram", usuario_red: "@ana.samaria", nicho: "Viajes y turismo",
    ciudad: "Santa Marta", correo: yo,
  });
  Logger.log("Correos de prueba enviados a " + yo);
}

