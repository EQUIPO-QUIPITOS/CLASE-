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
const SECRETO = "quipitos 2026";

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

// Lista de creadores (si pide un emprendedor) o de emprendimientos (si pide un creador).
// Los creadores "General" sirven para cualquier categoría: salen después de los especializados.
function recomendar(p) {
  const categoria = String(p.categoria || "").trim();
  const sub = String(p.sub || "").trim();
  const paraEmprendedor = p.tipo === "emprendedor";
  const creadorGeneral = !paraEmprendedor && categoria === "General";
  if (CATEGORIAS.indexOf(categoria) < 0 && !creadorGeneral) return { ok: true, items: [], sinCategoria: true };

  const filas = leerTabla(paraEmprendedor ? "Creadores" : "Emprendimientos");
  const campo = paraEmprendedor ? "nicho" : "sector";

  const items = filas
    .filter((f) => {
      if (!esSi(f.verificado) || !enSantaMarta(f.ciudad)) return false;
      const cat = categoriaDe(f[campo]);
      if (creadorGeneral) return CATEGORIAS.indexOf(cat) >= 0;
      return cat === categoria || (paraEmprendedor && cat === "General");
    })
    .map((f) => {
      const general = categoriaDe(f[campo]) === "General";
      const afin = !general && String(f[campo]).trim() === sub;
      return paraEmprendedor
        ? { id: f.id, nombre: f.nombre, red: f.red, usuario: f.usuario_red, seguidores: f.seguidores, area: f.nicho, afin: afin, general: general }
        : { id: f.id, nombre: f.negocio, area: f.sector, objetivo: f.objetivo, presupuesto: f.presupuesto, afin: afin, general: false };
    })
    .sort((a, b) => (b.afin - a.afin) || (a.general - b.general) || (numero(b.seguidores) - numero(a.seguidores)))
    .slice(0, 8);

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

  const detalles = [
    ["Negocio", emp.negocio],
    ["Creador", cre.nombre + " (" + cre.usuario_red + ")"],
    ["Paquete", paquete.nombre],
    ["Valor", dinero(paquete.precio) + " COP"],
    ["Para el creador (70 %)", dinero(paquete.pago) + " COP"],
    ["Hand in Hand (30 %)", dinero(paquete.comision) + " COP"],
  ];

  // 1) Confirmación para el emprendedor
  enviarCorreo(emp.correo, "Tu solicitud fue enviada: " + cre.nombre, "¡Tu solicitud va en camino!", [
    "Le enviaste una solicitud a <b>" + esc(cre.nombre) + "</b> (" + esc(cre.usuario_red) + ").",
    "Te avisamos por correo apenas el creador responda.",
  ], null, { etiqueta: "EMPRENDEDORES", detalles: detalles });

  // 2) Aviso al creador con botones para aceptar o rechazar
  if (cre.correo) {
    enviarCorreo(cre.correo, "Nueva solicitud de " + emp.negocio, "¡Tienes una solicitud!", [
      "<b>" + esc(emp.negocio) + "</b> quiere trabajar contigo. ¿Aceptas la campaña?",
      "Objetivo del negocio: " + esc(emp.objetivo || "—"),
    ], [["Aceptar solicitud", enlace("aceptar")], ["Rechazar", enlace("rechazar")]],
    { etiqueta: "CREADORES", detalles: [
      ["Negocio", emp.negocio], ["Sector", emp.sector], ["Paquete", paquete.nombre],
      ["Tú recibirías", dinero(paquete.pago) + " COP (70 %)"],
    ] });
  }

  // 3) Aviso para el equipo
  avisarEquipo("Nueva solicitud " + fila.id, emp.negocio + " envió una solicitud a " + cre.nombre + ".", detalles.concat([["Solicitud", fila.id]]));

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
      "<b>" + esc(cre.nombre) + "</b> quiere crear contenido para <b>" + esc(emp.negocio) + "</b>.",
      "Entra a Hand in Hand, regístrate con este mismo correo y elige a este creador para enviarle tu solicitud.",
    ], null, { etiqueta: "EMPRENDEDORES", detalles: [["Creador", cre.nombre], ["Perfil", cre.usuario_red + " · " + cre.red], ["Seguidores", cre.seguidores], ["Categoría", cre.nicho]] });
  }
  // Confirmación para el creador
  if (cre.correo) {
    enviarCorreo(cre.correo, "Registramos tu interés en " + emp.negocio, "¡Listo, avisamos al negocio!", [
      "Le contamos a <b>" + esc(emp.negocio) + "</b> (" + esc(emp.sector) + ") que te interesa trabajar con ellos.",
      "Si el negocio te elige, te llegará un correo con los botones para aceptar o rechazar la campaña.",
    ], null, { etiqueta: "CREADORES", detalles: [["Negocio", emp.negocio], ["Sector", emp.sector]] });
  }
  avisarEquipo("Interés de creador", cre.nombre + " mostró interés en " + emp.negocio + ".", [["Creador", cre.nombre], ["Negocio", emp.negocio]]);
  return { ok: true };
}

// Página que ve el creador al pulsar Aceptar o Rechazar en el correo
function paginaRespuesta(p) {
  const id = String(p.id || "");
  if (p.t !== firmar(id)) return pagina("Enlace no válido", "Este enlace no es válido o fue modificado.", false);

  const h = hojaDe("Solicitudes");
  const titulos = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map((x) => String(x).trim());
  const fila = leerTabla("Solicitudes").filter((f) => f.id === id)[0];
  if (!fila) return pagina("No encontramos la solicitud", "Puede que ya no exista.", false);
  if (fila.estado !== "enviada") return pagina("Ya respondiste", "Esta solicitud ya está marcada como «" + esc(fila.estado) + "».", false);

  const acepta = p.r === "aceptar";
  const estado = acepta ? "aceptada" : "rechazada";
  h.getRange(fila._fila, titulos.indexOf("estado") + 1).setValue(estado);

  if (fila.correo_emprendedor) {
    const nom = fila.creador.replace(/ \(.*$/, "");
    const det = [["Creador", nom], ["Paquete", fila.paquete], ["Estado", acepta ? "Aceptada" : "No aceptada"]];
    if (acepta) {
      enviarCorreo(fila.correo_emprendedor, "¡" + nom + " aceptó tu solicitud!", "¡Solicitud aceptada!", [
        "<b>" + esc(nom) + "</b> aceptó tu solicitud de <b>" + esc(fila.paquete) + "</b>.",
        "Te escribiremos por WhatsApp para coordinar los detalles de la campaña.",
      ], null, { etiqueta: "EMPRENDEDORES", detalles: det });
    } else {
      enviarCorreo(fila.correo_emprendedor, "Tu solicitud no fue aceptada", "El creador no pudo esta vez", [
        "<b>" + esc(nom) + "</b> no pudo tomar esta campaña.",
        "Entra de nuevo a Hand in Hand y elige otro creador de tu categoría.",
      ], null, { etiqueta: "EMPRENDEDORES", detalles: det });
    }
  }
  avisarEquipo("Solicitud " + id + " " + estado, "El creador " + (acepta ? "aceptó" : "rechazó") + " la solicitud.", [["Negocio", fila.emprendimiento], ["Creador", fila.creador], ["Paquete", fila.paquete], ["Estado", estado]]);
  return pagina(acepta ? "¡Aceptaste la solicitud!" : "Solicitud rechazada",
    acepta ? "Avisamos al negocio. Te escribiremos por WhatsApp para coordinar." : "Avisamos al negocio. Gracias por responder.", acepta);
}

const COL = {
  fondo: "#E3E4EA", tarjeta: "#F6F7FB", oscuro: "#00171F", teal: "#005260",
  verde: "#70A312", lima: "#B5EA3C", cian: "#00E4E8", menta: "#C5DFDF", texto: "#1d2b30", gris: "#4f6a70",
};
const FUENTE = "'Inter',Helvetica,Arial,sans-serif";
const CURSIVA = "'Dancing Script','Brush Script MT','Segoe Script',cursive";

// Marca "Hand in Hand" como en la página
function marca(acento) {
  return 'Hand<span style="font:700 0.86em ' + CURSIVA + ';letter-spacing:0;color:' + acento + '">in</span>Hand';
}

// Página que ve el creador al pulsar Aceptar o Rechazar en el correo
function pagina(titulo, texto, ok) {
  const color = ok === false ? "#B3261E" : COL.verde;
  const icono = ok === false ? "✕" : "✓";
  return HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<link href="https://fonts.googleapis.com/css2?family=Dancing+Script:wght@700&family=Inter:wght@400;700;900&display=swap" rel="stylesheet"></head>' +
    '<body style="margin:0;background:' + COL.fondo + ';font-family:' + FUENTE + ';display:flex;min-height:100vh;align-items:center;justify-content:center">' +
    '<div style="width:100%;max-width:460px;margin:20px;border-radius:28px;overflow:hidden;background:' + COL.tarjeta + ';box-shadow:0 18px 40px rgba(0,23,31,.18)">' +
      '<div style="background:#003D33;background-image:linear-gradient(180deg,#00513F,#002B26);padding:30px 32px">' +
        '<div style="font-weight:800;font-size:13px;letter-spacing:2px;color:#C8F5DC">SANTA MARTA</div>' +
        '<div style="font-weight:900;font-size:40px;letter-spacing:-2px;color:' + COL.fondo + ';margin-top:4px">' + marca(COL.lima) + '</div>' +
      '</div>' +
      '<div style="padding:34px 32px 38px;text-align:center">' +
        '<div style="width:72px;height:72px;border-radius:50%;background:' + color + ';color:#fff;font-size:38px;line-height:72px;margin:0 auto 18px;font-weight:700">' + icono + '</div>' +
        '<h1 style="color:' + COL.teal + ';margin:0 0 10px;font-size:26px;line-height:1.2">' + titulo + '</h1>' +
        '<p style="color:' + COL.gris + ';line-height:1.6;margin:0;font-size:16px">' + texto + '</p>' +
      '</div>' +
      '<div style="background:' + COL.teal + ';padding:16px;text-align:center;font:700 22px ' + CURSIVA + ';color:' + COL.cian + '">Gracias por confiar en nosotros</div>' +
    '</div></body></html>'
  ).setTitle("Hand in Hand");
}

// Aviso interno para el equipo (mismo diseño que los demás correos)
function avisarEquipo(asunto, texto, detalles) {
  try {
    enviarCorreo(Session.getEffectiveUser().getEmail(), "Hand in Hand: " + asunto, asunto, [esc(texto)], null,
      { etiqueta: "AVISO PARA EL EQUIPO", detalles: detalles, pie: "Este aviso es solo para el equipo de Hand in Hand." });
  } catch (err) {}
}

// Correo con el diseño de la página. "lineas" ya viene con HTML seguro.
// extra: { etiqueta, detalles: [[nombre, valor], ...], pie }
function enviarCorreo(para, asunto, titulo, lineas, botones, extra) {
  try {
    extra = extra || {};
    const etiqueta = extra.etiqueta || "HAND IN HAND";
    const filas = (extra.detalles || []).filter((d) => d[1] !== "" && d[1] != null).map((d) =>
      '<tr><td style="padding:6px 0;font:600 14px ' + FUENTE + ';color:' + COL.gris + ';width:120px;vertical-align:top">' + esc(d[0]) + '</td>' +
      '<td style="padding:6px 0;font:700 15px ' + FUENTE + ';color:' + COL.texto + '">' + esc(d[1]) + '</td></tr>').join("");

    const caja = filas
      ? '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + COL.menta + ';border-radius:16px;margin:20px 0 4px"><tr><td style="padding:16px 22px">' +
        '<div style="font:700 26px ' + CURSIVA + ';color:' + COL.teal + ';margin-bottom:4px">Detalles</div>' +
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + filas + '</table></td></tr></table>'
      : "";

    const btns = (botones || []).map((b, i) =>
      '<td style="padding:0 12px 12px 0"><a href="' + b[1] + '" style="display:inline-block;padding:15px 34px;border-radius:999px;text-decoration:none;font:800 17px ' + FUENTE + ';' +
      (i === 0 ? 'background:' + COL.lima + ';color:#1F3A07;border:2px solid ' + COL.lima : 'background:#ffffff;color:' + COL.teal + ';border:2px solid ' + COL.teal) +
      '">' + (i === 0 ? "✓ " : "") + b[0] + '</a></td>').join("");
    const zonaBotones = btns
      ? '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:24px 0 4px"><tr>' + btns + '</tr></table>' +
        '<p style="font:400 13px/1.5 ' + FUENTE + ';color:' + COL.gris + ';margin:8px 0 0">Pulsa uno de los botones para responder. No necesitas iniciar sesión.</p>'
      : "";

    const html =
      '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<link href="https://fonts.googleapis.com/css2?family=Dancing+Script:wght@700&family=Inter:wght@400;600;700;800;900&display=swap" rel="stylesheet"></head>' +
      '<body style="margin:0;padding:0;background:' + COL.fondo + '">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + COL.fondo + '"><tr><td align="center" style="padding:28px 12px">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px">' +
        '<tr><td style="background:#003D33;background-image:linear-gradient(180deg,#00513F,#002B26);border-radius:24px 24px 0 0;padding:34px 32px 30px">' +
          '<div style="font:800 14px ' + FUENTE + ';letter-spacing:2px;color:#C8F5DC">' + esc(etiqueta) + '</div>' +
          '<div style="font:900 44px/1.05 ' + FUENTE + ';letter-spacing:-2px;color:' + COL.fondo + ';margin-top:6px">' + marca(COL.lima) + '</div>' +
          '<div style="font:400 16px ' + FUENTE + ';color:' + COL.fondo + ';margin-top:8px">Estamos aquí para darte una mano</div>' +
        '</td></tr>' +
        '<tr><td style="background:' + COL.tarjeta + ';padding:32px 32px 30px">' +
          '<div style="font:700 32px/1.15 ' + CURSIVA + ';color:' + COL.verde + '">' + titulo + '</div>' +
          lineas.map((l) => '<p style="font:400 16px/1.6 ' + FUENTE + ';color:' + COL.texto + ';margin:14px 0 0">' + l + '</p>').join("") +
          caja + zonaBotones +
        '</td></tr>' +
        '<tr><td style="background:' + COL.teal + ';padding:24px 32px;border-radius:0 0 24px 24px;text-align:center">' +
          '<div style="font:700 28px ' + CURSIVA + ';color:' + COL.cian + '">¡Gracias por confiar en nosotros!</div>' +
          '<div style="font:400 14px/1.6 ' + FUENTE + ';color:#d9ecec;margin-top:6px">¿Dudas? Escríbenos a ' +
          '<a href="mailto:equipoquipitosoficial@gmail.com" style="color:' + COL.lima + ';font-weight:600;text-decoration:none">equipoquipitosoficial@gmail.com</a></div>' +
        '</td></tr>' +
        '<tr><td style="padding:16px 10px;text-align:center;font:400 12px/1.5 ' + FUENTE + ';color:' + COL.gris + '">' +
          esc(extra.pie || "Hand in Hand · Santa Marta, Colombia. Usamos tus datos solo para gestionar campañas.") +
        '</td></tr>' +
      '</table></td></tr></table></body></html>';

    const plano = lineas.join("\n").replace(/<[^>]+>/g, "") +
      (botones || []).map((b) => "\n" + b[0] + ": " + b[1]).join("");
    MailApp.sendEmail({ to: para, subject: asunto, name: "Hand in Hand", htmlBody: html, body: plano });
  } catch (err) {
    // Si el correo falla, la solicitud ya quedó guardada
  }
}
