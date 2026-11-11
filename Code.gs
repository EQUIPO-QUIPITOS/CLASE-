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

// Atiende las consultas de la página (GET):
//  - recomendaciones: lista de creadores o emprendimientos afines
//  - crear_solicitud: el emprendedor elige creador y paquete
//  - mostrar_interes: el creador dice "me interesa"
//  - responder: el creador acepta o rechaza desde el correo
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === "recomendaciones") return responder(recomendar(p));
    if (p.action === "crear_solicitud") return responder(crearSolicitud(p));
    if (p.action === "mostrar_interes") return responder(mostrarInteres(p));
    if (p.action === "responder") return paginaRespuesta(p);
    return responder({ ok: true, mensaje: "El servicio está activo" });
  } catch (err) {
    return responder({ ok: false, error: String(err) });
  }
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
    ? "Gracias por postularte como creador y por confiar en nosotros. Nos encanta que quieras crecer de la mano de los emprendimientos de Santa Marta."
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



/* =====================================================================
   RECOMENDACIONES Y SOLICITUDES (el corazón del MVP)
   ===================================================================== */

// Frase secreta para firmar los enlaces "Aceptar / Rechazar" de los correos.
// Cámbiala por cualquier frase tuya (solo tú debes conocerla).
const SECRETO = "hand-in-hand-cambia-esta-frase";

const CATEGORIAS = ["Comida", "Moda", "Maquillaje"];

// Los precios viven aquí, no en la página: así nadie puede cambiarlos desde el navegador
const PAQUETES = {
  basica:   { nombre: "Campaña básica",   precio: 300000,  comision: 90000,  pago: 210000 },
  feed:     { nombre: "Básico + feed",    precio: 550000,  comision: 165000, pago: 385000 },
  completo: { nombre: "Paquete completo", precio: 1240000, comision: 372000, pago: 868000 },
};

function hojaDe(nombre) {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!h) throw new Error("No existe la pestaña " + nombre);
  return h;
}

// Lee una pestaña como lista de filas con los títulos de la fila 1 como nombres
function leerTabla(nombre) {
  const v = hojaDe(nombre).getDataRange().getDisplayValues();
  const t = v[0].map((x) => String(x).trim());
  return v.slice(1).map((r, i) => {
    const o = { _fila: i + 2 };
    t.forEach((k, j) => { o[k] = r[j]; });
    return o;
  });
}

function categoriaDe(texto) {
  return String(texto || "").split(":")[0].trim();
}

function esSi(v) { return String(v || "").trim().toLowerCase() === "si"; }
function enSantaMarta(v) { return String(v || "").toLowerCase().indexOf("santa marta") >= 0; }
function minusculas(v) { return String(v || "").trim().toLowerCase(); }

// Lista de creadores (si pide un emprendedor) o de emprendimientos (si pide un creador)
function recomendar(p) {
  const categoria = String(p.categoria || "").trim();
  const sub = String(p.sub || "").trim();
  if (CATEGORIAS.indexOf(categoria) < 0) return { ok: true, items: [], sinCategoria: true };

  const paraEmprendedor = p.tipo === "emprendedor";
  const filas = leerTabla(paraEmprendedor ? "Creadores" : "Emprendimientos");
  const campo = paraEmprendedor ? "nicho" : "sector";

  const items = filas
    .filter((f) => esSi(f.verificado) && enSantaMarta(f.ciudad) && categoriaDe(f[campo]) === categoria)
    .map((f) => {
      const afin = String(f[campo]).trim() === sub;
      return paraEmprendedor
        ? { id: f.id, nombre: f.nombre, red: f.red, usuario: f.usuario_red, seguidores: f.seguidores, area: f.nicho, afin: afin }
        : { id: f.id, nombre: f.negocio, area: f.sector, objetivo: f.objetivo, presupuesto: f.presupuesto, afin: afin };
    })
    .sort((a, b) => (b.afin - a.afin) || (numero(b.seguidores) - numero(a.seguidores)))
    .slice(0, 6);

  return { ok: true, items: items };
}

function numero(v) {
  const n = parseInt(String(v || "").replace(/[^0-9]/g, ""), 10);
  return isNaN(n) ? 0 : n;
}

// Agrega una fila a una pestaña usando sus títulos; crea el id (S001, S002…)
function agregarFila(nombre, prefijo, datos) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const h = hojaDe(nombre);
    const t = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map((x) => String(x).trim());
    datos.id = prefijo + String(h.getLastRow()).padStart(3, "0");
    datos.fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    h.appendRow(t.map((k) => (k in datos ? datos[k] : "")));
    return datos;
  } finally {
    lock.releaseLock();
  }
}

function ultimaPorCorreo(nombreHoja, correo) {
  const c = minusculas(correo);
  const filas = leerTabla(nombreHoja).filter((f) => minusculas(f.correo) === c);
  return filas.length ? filas[filas.length - 1] : null;
}

function firmar(id) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, id + SECRETO);
  return Utilities.base64EncodeWebSafe(bytes).slice(0, 16);
}

function dinero(n) { return "$" + String(n).replace(/\B(?=(\d{3})+(?!\d))/g, "."); }

// El emprendedor elige un creador y un paquete
function crearSolicitud(p) {
  const paquete = PAQUETES[p.paquete];
  if (!paquete) return { ok: false, error: "Paquete no válido" };

  const emp = ultimaPorCorreo("Emprendimientos", p.correo);
  if (!emp) return { ok: false, error: "No encontramos tu registro" };

  const cre = leerTabla("Creadores").filter((f) => f.id === p.creador_id && esSi(f.verificado))[0];
  if (!cre) return { ok: false, error: "Ese creador no está disponible" };

  const fila = agregarFila("Solicitudes", "S", {
    emprendimiento: emp.negocio,
    correo_emprendedor: emp.correo,
    creador: cre.nombre + " (" + cre.id + ")",
    paquete: paquete.nombre,
    precio: paquete.precio,
    pago_creador: paquete.pago,
    comision_hih: paquete.comision,
    estado: "enviada",
  });

  const base = ScriptApp.getService().getUrl();
  const enlace = (r) => base + "?action=responder&id=" + fila.id + "&r=" + r + "&t=" + firmar(fila.id);

  // 1) Confirmación para el emprendedor
  enviarCorreo(emp.correo, "Tu solicitud fue enviada: " + cre.nombre, "¡Tu solicitud va en camino!", [
    "Le enviaste una solicitud a <b>" + esc(cre.nombre) + "</b> (" + esc(cre.usuario_red) + ").",
    "Paquete: <b>" + esc(paquete.nombre) + "</b> · " + dinero(paquete.precio) + " COP.",
    "El creador recibe el " + "70 % (" + dinero(paquete.pago) + ") y Hand in Hand el 30 % (" + dinero(paquete.comision) + ").",
    "Te avisamos por correo apenas el creador responda.",
  ]);

  // 2) Aviso al creador con botones para aceptar o rechazar
  if (cre.correo) {
    enviarCorreo(cre.correo, "Nueva solicitud de " + emp.negocio, "¡Tienes una solicitud!", [
      "<b>" + esc(emp.negocio) + "</b> (" + esc(emp.sector) + ") quiere trabajar contigo.",
      "Paquete: <b>" + esc(paquete.nombre) + "</b>. Tú recibirías <b>" + dinero(paquete.pago) + " COP</b> (el 70 %).",
      "Objetivo del negocio: " + esc(emp.objetivo || "—"),
    ], [["Aceptar", enlace("aceptar")], ["Rechazar", enlace("rechazar")]]);
  }

  // 3) Aviso para ti
  avisarEquipo("Nueva solicitud " + fila.id, emp.negocio + " → " + cre.nombre + " · " + paquete.nombre + " · " + dinero(paquete.precio));

  return { ok: true, id: fila.id };
}

// El creador pulsa "Me interesa" en un emprendimiento
function mostrarInteres(p) {
  const cre = ultimaPorCorreo("Creadores", p.correo);
  if (!cre) return { ok: false, error: "No encontramos tu registro" };
  const emp = leerTabla("Emprendimientos").filter((f) => f.id === p.emprendimiento_id && esSi(f.verificado))[0];
  if (!emp) return { ok: false, error: "Ese emprendimiento no está disponible" };

  const marca = cre.nombre + " (" + cre.id + ")";
  const ya = leerTabla("Solicitudes").some((f) => f.estado === "interés del creador" && f.creador === marca && f.emprendimiento === emp.negocio);
  if (ya) return { ok: true, repetido: true };

  agregarFila("Solicitudes", "S", {
    emprendimiento: emp.negocio,
    correo_emprendedor: emp.correo,
    creador: marca,
    estado: "interés del creador",
  });

  if (emp.correo) {
    enviarCorreo(emp.correo, cre.nombre + " quiere trabajar contigo", "¡Un creador se interesó en tu negocio!", [
      "<b>" + esc(cre.nombre) + "</b> (" + esc(cre.usuario_red) + " en " + esc(cre.red) + ", " + esc(cre.seguidores) + " seguidores) quiere crear contenido para <b>" + esc(emp.negocio) + "</b>.",
      "Entra a Hand in Hand, regístrate con este mismo correo y elige a este creador para enviarle tu solicitud.",
    ]);
  }
  avisarEquipo("Interés de creador", cre.nombre + " → " + emp.negocio);
  return { ok: true };
}

// Página que ve el creador al pulsar Aceptar o Rechazar en el correo
function paginaRespuesta(p) {
  const id = String(p.id || "");
  if (p.t !== firmar(id)) return pagina("Enlace no válido", "Este enlace no es válido o fue modificado.");

  const h = hojaDe("Solicitudes");
  const titulos = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map((x) => String(x).trim());
  const fila = leerTabla("Solicitudes").filter((f) => f.id === id)[0];
  if (!fila) return pagina("No encontramos la solicitud", "Puede que ya no exista.");
  if (fila.estado !== "enviada") return pagina("Ya respondiste", "Esta solicitud ya está marcada como «" + esc(fila.estado) + "».");

  const acepta = p.r === "aceptar";
  const estado = acepta ? "aceptada" : "rechazada";
  h.getRange(fila._fila, titulos.indexOf("estado") + 1).setValue(estado);

  if (fila.correo_emprendedor) {
    enviarCorreo(
      fila.correo_emprendedor,
      acepta ? "¡" + fila.creador.replace(/ \(.*$/, "") + " aceptó tu solicitud!" : "Tu solicitud no fue aceptada",
      acepta ? "¡Solicitud aceptada!" : "El creador no pudo esta vez",
      acepta
        ? ["Tu solicitud de <b>" + esc(fila.paquete) + "</b> fue aceptada. Te escribiremos por WhatsApp para coordinar los detalles."]
        : ["El creador no pudo tomar esta campaña. Entra de nuevo a Hand in Hand y elige otro creador de tu categoría."]
    );
  }
  avisarEquipo("Solicitud " + id + " " + estado, fila.emprendimiento + " · " + fila.creador);
  return pagina(acepta ? "¡Aceptaste la solicitud!" : "Solicitud rechazada",
    acepta ? "Avisamos al negocio. Te escribiremos por WhatsApp para coordinar." : "Avisamos al negocio. Gracias por responder.");
}

function pagina(titulo, texto) {
  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0;background:#E3E4EA;font-family:Helvetica,Arial,sans-serif;display:flex;min-height:100vh;align-items:center;justify-content:center">' +
    '<div style="max-width:420px;margin:20px;background:#fff;border-radius:24px;padding:36px;text-align:center">' +
    '<div style="font-size:13px;font-weight:800;letter-spacing:2px;color:#005260">HAND IN HAND</div>' +
    '<h1 style="color:#00473F;margin:14px 0 10px">' + titulo + '</h1><p style="color:#4f6a70;line-height:1.5;margin:0">' + texto + '</p></div></body></html>'
  ).setTitle("Hand in Hand");
}

function avisarEquipo(asunto, texto) {
  try { MailApp.sendEmail(Session.getEffectiveUser().getEmail(), "Hand in Hand: " + asunto, texto); } catch (err) {}
}

// Correo sencillo con los colores de la página (lineas ya vienen con HTML seguro)
function enviarCorreo(para, asunto, titulo, lineas, botones) {
  try {
    const f = "Helvetica,Arial,sans-serif";
    const btn = (botones || []).map((b, i) =>
      '<a href="' + b[1] + '" style="display:inline-block;margin:6px 6px 0 0;padding:12px 22px;border-radius:999px;text-decoration:none;font:700 15px ' + f +
      ';background:' + (i === 0 ? "#B5EA3C" : "#C5DFDF") + ';color:#00171F">' + b[0] + '</a>').join("");
    const html =
      '<div style="background:#E3E4EA;padding:24px 12px;font-family:' + f + '"><div style="max-width:560px;margin:0 auto">' +
      '<div style="background:#003D33;border-radius:24px 24px 0 0;padding:26px 30px"><div style="font:800 13px ' + f + ';letter-spacing:2px;color:#C8F5DC">HAND IN HAND</div>' +
      '<div style="font:800 26px/1.2 ' + f + ';color:#E3E4EA;margin-top:8px">' + titulo + '</div></div>' +
      '<div style="background:#F6F7FB;border-radius:0 0 24px 24px;padding:26px 30px;color:#1d2b30;font-size:16px;line-height:1.6">' +
      lineas.map((l) => '<p style="margin:0 0 12px">' + l + '</p>').join("") + btn +
      '<p style="margin:18px 0 0;font-size:12px;color:#4f6a70">Hand in Hand · Santa Marta · equipoquipitosoficial@gmail.com</p></div></div></div>';
    MailApp.sendEmail({ to: para, subject: asunto, name: "Hand in Hand", htmlBody: html, body: lineas.join("\n").replace(/<[^>]+>/g, "") });
  } catch (err) {
    // Si el correo falla, la solicitud ya quedó guardada
  }
}
