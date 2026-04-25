var CONFIG = {
  STORE_ID:        '7396246',
  API_TOKEN:       'e3d744d94ddbf13317bef0082c53e2c46fb50631',
  SECRET:          'retiro2026',
  TIENDA_NOMBRE:   'Diente de León',
  STAFF_TOKEN:     'test',
  VENTAS_SHEET_ID: '1-57n6RmTFQjwNFVYxNPzll8MMuV4NvXXx5v0wTvR_6g',

  STAFF_PINS: {
    'Yuliana.Longhi ':   'PYL26',   // Jardín
    'Maria.Martini ':    'PMM26',   // Jardín
    'Ines.Robertson':    'PIR26',   // 4to
    'Delfina.Salvetti':  'PDS26',   // 6to
    'Dolo.Dominguez':    'PDD26',   // 2do
    'Gaby.Dominguez':    'PGD26',   // 5to
    'Juan.Diaz':         'PJD26',   // 8vo
    'Karina':            'PK26',    // 7mo
    'Monica.Chesta':     'PMC26',   // —
    'Paola':             'PP26',    // 2do
    'Pia':               'PPI26',   // 11vo
    'Sabi':              'PSA26',   // 9no
    'Yana':              'PYA26',   // 5to
    'Elina':             'PEL26',   // 3ro
    'Claudio':           'PCL26'    // Tiendita
  },

  // PINs con permiso para confirmar entregas (Retiro Mercadería)
  RETIRO_PINS: [
    'PYL26',  // Yuliana Longhi   — Jardín
    'PMM26',  // Maria Martini    — Jardín
    'PIR26',  // Ines Robertson   — 4to
    'PDS26',  // Delfina Salvetti — 6to
    'PDD26',  // Dolo Dominguez   — 2do
    'PGD26',  // Gaby Dominguez   — 5to
    'PJD26',  // Juan Diaz        — 8vo
    'PK26',   // Karina           — 7mo
    'PMC26',  // Monica Chesta    — —
    'PP26',   // Paola            — 2do
    'PPI26',  // Pia              — 11vo
    'PSA26',  // Sabi             — 9no
    'PYA26',  // Yana             — 5to
    'PEL26',  // Elina            — 3ro
    'PCL26'   // Claudio          — Tiendita
  ],

  // Solo dashboard — no pueden confirmar entregas
  DASHBOARD_PINS: {
    'Maxi.Contreras':      'PMAC26',
    'Lucas.DiStefano':     'PLDS26',
    'Esteban.Prospero':    'PEP26',
    'Jose.Brizzi':         'PJB26',
    'Anabella.Gargiulo':   'PAG26',   // Comisión de Recursos — sin acceso QR
    'Hector.Larrea':       'PHL26'    // Prof. Héctor Larrea — solo Dashboard General
  },

};

var LOGO_B64 = "https://escuela-dandelion.github.io/Comision-Recursos/Logo_Diente_de_Leon.png";

// ═══════════════════════════════════════════════════════════
//  PUNTO DE RETIRO — Tiendanube + QR automático
// ═══════════════════════════════════════════════════════════


// ──────────────────────────────────────────────────────────
// ENTRY POINTS
// ──────────────────────────────────────────────────────────

function doPost(e) {
  try {
    var body    = JSON.parse(e.postData.contents);
    var evento  = body.event || '';
    var orderId = String(body.id);

    if (evento === 'order/paid' || evento === 'order/updated') {
      var order = fetchOrder(orderId);
      if (order && order.payment_status === 'paid' && !yaFueEnviado(orderId)) {
        procesarPedido(order);
      }
    }
    return ContentService.createTextOutput('OK');
  } catch(err) {
    Logger.log('doPost error: ' + err);
    return ContentService.createTextOutput('error');
  }
}

function doGet(e) {
  try {
    return doGetInterno(e);
  } catch(err) {
    Logger.log('doGet error: ' + err + ' | stack: ' + err.stack);
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:24px;color:red">' +
      '<strong>Error:</strong> ' + err.message + '</div>'
    );
  }
}

function doGetInterno(e) {
  var action    = e.parameter.action    || '';
  var orderId   = e.parameter.id        || '';
  var token     = e.parameter.token     || '';
  var pass      = e.parameter.pass      || '';
  var staffName = e.parameter.staff     || '';
  var pin       = e.parameter.pin       || '';
  var entregado = e.parameter.entregado || '';

  if (action === 'version') {
    return ContentService.createTextOutput('Version_1.5');
  }

  if (action === 'portal') {
    if (pass !== CONFIG.STAFF_TOKEN) {
      return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px;color:red">Acceso no autorizado.</p>');
    }
    return paginaPortalStaff(entregado === '1');
  }

  // Endpoint JSON para verificación de PIN desde el cliente (AJAX)
  if (action === 'api_entregar') {
    var notasParam = e.parameter.notas || '';
    var resultado = apiEntregar(orderId, staffName, pin, pass, token, notasParam);
    return ContentService
      .createTextOutput(JSON.stringify(resultado))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Entrega directa desde el portal (admin, sin QR) — muestra el formulario de PIN
  if (action === 'entregar_admin') {
    if (pass !== CONFIG.STAFF_TOKEN) {
      return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px;color:red">Acceso no autorizado.</p>');
    }
    return paginaConfirmar(orderId, generarToken(orderId), null, pass);
  }

  if (action === 'addNotasRetiro') {
    var ordId  = e.parameter.id    || '';
    var notas  = e.parameter.notas || '';
    var sheet  = getSheet();
    var data   = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(ordId)) {
        sheet.getRange(i + 1, 10).setValue(notas);
        return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Pedido no encontrado' })).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getRetiros') {
    var estadoFiltro = e.parameter.estado || '';  // 'Pendiente', 'Entregado', o '' = todos
    var diasFiltro   = parseInt(e.parameter.dias || '0');
    var retiros = getRetirosData(estadoFiltro, diasFiltro);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, retiros: retiros }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'api_dashboard') {
    var res = apiDashboard(pin, e.parameter.email || '');
    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  }

  // ── AUTH para admin.html ──────────────────────────────────

  if (action === 'verificarEmail') {
    var res = apiVerificarEmail(e.parameter.email || '');
    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'sendOTP') {
    var res = apiSendOTP(e.parameter.tel || '');
    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'verifyOTP') {
    var res = apiVerifyOTP(e.parameter.tel || '', e.parameter.code || '');
    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'api_verificar') {
    var res = apiVerificar(orderId, token);
    return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
  }

  if (!orderId || !verificarToken(orderId, token)) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px;color:red">QR inválido o expirado.</p>');
  }

  if (action === 'verificar')  return paginaVerificar(orderId, token);
  if (action === 'confirmar')  return paginaConfirmar(orderId, token);

  return HtmlService.createHtmlOutput('<p>URL inválida.</p>');
}


// ──────────────────────────────────────────────────────────
// LÓGICA PRINCIPAL
// ──────────────────────────────────────────────────────────

function procesarPedido(order) {
  var orderId    = String(order.id);
  var buyerEmail = order.contact_email;
  var buyerName  = (order.contact_name || 'comprador').split(' ')[0];
  var productos  = order.products.map(function(p) {
    return p.quantity + 'x ' + p.name;
  }).join('\n');

  var webAppUrl  = ScriptApp.getService().getUrl();
  var token      = generarToken(orderId);
  var verifyUrl  = webAppUrl + '?action=verificar&id=' + orderId + '&token=' + token;
  var qrImageUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&data='
                   + encodeURIComponent(verifyUrl);

  enviarEmail(buyerEmail, buyerName, order, qrImageUrl, productos);
  registrarEnSheet(order, token);
  registrarEnVentas(order);

  Logger.log('Pedido #' + order.number + ' procesado OK — email enviado a ' + buyerEmail);
}

// ──────────────────────────────────────────────────────────
// TIENDANUBE API
// ──────────────────────────────────────────────────────────

function fetchOrder(orderId) {
  var url  = 'https://api.tiendanube.com/v1/' + CONFIG.STORE_ID + '/orders/' + orderId;
  var resp = UrlFetchApp.fetch(url, {
    headers: {
      'Authentication': 'bearer ' + CONFIG.API_TOKEN,
      'User-Agent': CONFIG.TIENDA_NOMBRE + ' (retiro-automatico)',
    },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    Logger.log('fetchOrder error ' + resp.getResponseCode() + ': ' + resp.getContentText());
    return null;
  }
  return JSON.parse(resp.getContentText());
}

// ──────────────────────────────────────────────────────────
// EMAIL AL COMPRADOR
// ──────────────────────────────────────────────────────────

function enviarEmail(email, nombre, order, qrUrl, productos) {
  var asunto = CONFIG.TIENDA_NOMBRE + ' — Tu código QR para retirar el pedido #' + order.number;

  var html =
    '<div style="font-family:sans-serif;max-width:520px;margin:auto;padding:24px">' +
    '<div style="text-align:center;padding:16px 0 24px 0">' +
    '<img src="' + LOGO_B64 + '" style="max-width:220px;width:60%" alt="' + CONFIG.TIENDA_NOMBRE + '">' +
    '</div>' +
    '<h2 style="color:#1a1a1a">¡Hola ' + nombre + '!</h2>' +
    '<p>Tu pago fue confirmado. Guardá este QR para retirar tu pedido en el punto de retiro.</p>' +
    '<div style="background:#f5f5f5;padding:16px;border-radius:8px;margin:20px 0">' +
    '<p style="margin:0 0 6px 0"><strong>Pedido #' + order.number + '</strong></p>' +
    '<p style="margin:0;color:#555;white-space:pre-line">' + productos + '</p>' +
    '<p style="margin:10px 0 0 0;color:#555"><strong>Total: ' + order.total + '</strong></p>' +
    '</div>' +
    '<div style="text-align:center;margin:28px 0">' +
    '<img src="' + qrUrl + '" width="240" height="240"' +
    ' style="border:1px solid #eee;padding:10px;border-radius:8px" alt="QR Retiro"/>' +
    '<p style="color:#888;font-size:13px;margin-top:8px">Mostrá este QR en el punto de retiro</p>' +
    '</div>' +
    '<p style="font-size:13px;color:#aaa">¿Preguntas? Respondé este email.</p>' +
    '</div>';

  GmailApp.sendEmail(email, asunto, '', { htmlBody: html, from: 'tiendavirtual.dientedeleon@gmail.com' });
}

// ──────────────────────────────────────────────────────────
// GOOGLE SHEETS
// ──────────────────────────────────────────────────────────

function getSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Retiros');
  if (!sheet) {
    sheet = ss.insertSheet('Retiros');
    sheet.appendRow(['ID interno','Pedido #','Nombre','Email','Productos','Total','Fecha','Estado']);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,8).setFontWeight('bold');
  }
  return sheet;
}

function registrarEnSheet(order, token) {
  var sheet     = getSheet();
  var productos = order.products.map(function(p) {
    return p.quantity + 'x ' + p.name;
  }).join(' | ');

  sheet.appendRow([
    String(order.id),
    order.number,
    order.contact_name,
    order.contact_email,
    productos,
    order.total,
    new Date(),
    'PENDIENTE'
  ]);
}

function getVentasSheet() {
  var ss    = SpreadsheetApp.openById(CONFIG.VENTAS_SHEET_ID);
  var sheet = ss.getSheetByName('Ventas');
  if (!sheet) {
    sheet = ss.insertSheet('Ventas');
    sheet.appendRow([
      'Fecha','Pedido #','ID Interno','Nombre','Email',
      'Producto','SKU','Cantidad','Precio Unit.','Total Línea','Total Pedido','Comentarios','Marca'
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 13).setFontWeight('bold');
  }
  return sheet;
}

// Cache por "productId:variantId" → { marca, costo }
var _prodCache = {};

function fetchDatosDeProducto(productId, variantId) {
  if (!productId) return { marca: '', costo: 0 };
  var cacheKey = productId + ':' + (variantId || '');
  if (_prodCache[cacheKey] !== undefined) return _prodCache[cacheKey];
  try {
    var url  = 'https://api.tiendanube.com/v1/' + CONFIG.STORE_ID + '/products/' + productId;
    var resp = UrlFetchApp.fetch(url, {
      headers: {
        'Authentication': 'bearer ' + CONFIG.API_TOKEN,
        'User-Agent': CONFIG.TIENDA_NOMBRE + ' (retiro-automatico)'
      },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) { _prodCache[cacheKey] = { marca: '', costo: 0 }; return _prodCache[cacheKey]; }
    var prod = JSON.parse(resp.getContentText());
    var marca = String(prod.brand || '').trim().toUpperCase();

    // El costo está en la variante, no en el producto raíz
    var costo = 0;
    var variants = prod.variants || [];
    if (variantId) {
      for (var i = 0; i < variants.length; i++) {
        if (String(variants[i].id) === String(variantId)) {
          costo = parseFloat(variants[i].cost || 0);
          break;
        }
      }
    }
    // Fallback: primera variante si no hubo match
    if (!costo && variants.length > 0) {
      costo = parseFloat(variants[0].cost || 0);
    }

    var datos = { marca: marca, costo: costo };
    _prodCache[cacheKey] = datos;
    return datos;
  } catch(e) {
    Logger.log('fetchDatosDeProducto error: ' + e);
    _prodCache[cacheKey] = { marca: '', costo: 0 };
    return _prodCache[cacheKey];
  }
}

function fetchMarcaDeProducto(productId) {
  return fetchDatosDeProducto(productId, null).marca;
}

function registrarEnVentas(order) {
  try {
    var sheet  = getVentasSheet();
    var fecha  = new Date();
    var nota   = order.note || order.notes || '';

    order.products.forEach(function(p) {
      var nombre = (typeof p.name === 'object')
        ? (p.name.es || p.name.pt || Object.values(p.name)[0] || '')
        : (p.name || '');
      var sku        = p.sku || '';
      var qty        = p.quantity || 1;
      var precioUnit = parseFloat(p.price || 0);
      var totalLinea = precioUnit * qty;
      var datos      = fetchDatosDeProducto(p.product_id || '', p.variant_id || '');

      sheet.appendRow([
        fecha,
        order.number,
        String(order.id),
        order.contact_name  || '',
        order.contact_email || '',
        nombre,
        sku,
        qty,
        precioUnit,
        totalLinea,
        parseFloat(order.total || 0),
        nota,
        datos.marca,
        datos.costo
      ]);
    });
    Logger.log('Ventas: ' + order.products.length + ' fila(s) para pedido #' + order.number);
  } catch(err) {
    Logger.log('Error registrarEnVentas: ' + err);
  }
}

function yaFueEnviado(orderId) {
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(orderId)) return true;
  }
  return false;
}

function marcarEntregado(orderId, staffName, notas) {
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();
  var buyerEmail  = '';
  var buyerName   = '';
  var orderNumber = '';
  var productos   = '';

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(orderId)) {
      sheet.getRange(i + 1, 8).setValue('ENTREGADO ✓ — ' + new Date().toLocaleString('es-AR'));
      sheet.getRange(i + 1, 9).setValue(staffName);
      if (notas) sheet.getRange(i + 1, 10).setValue(notas);
      buyerEmail  = data[i][3];
      buyerName   = data[i][2];
      orderNumber = data[i][1];
      productos   = data[i][4];
      break;
    }
  }

  marcarFulfillmentTiendanube(orderId);

  if (buyerEmail) {
    enviarEmailEntrega(buyerEmail, buyerName, orderNumber, productos);
  }

}

// ──────────────────────────────────────────────────────────
// PÁGINAS HTML
// ──────────────────────────────────────────────────────────

// JSON endpoint: verifica token, consulta Tiendanube y devuelve datos del pedido
function apiVerificar(orderId, token) {
  if (!orderId || !verificarToken(orderId, token)) {
    return { ok: false, error: 'QR inválido o expirado.' };
  }

  // ¿Ya fue entregado?
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(orderId)) {
      if (String(data[i][7]).indexOf('ENTREGADO') !== -1) {
        return {
          ok:            true,
          entregado:     true,
          pedido_numero: data[i][1],
          estado:        String(data[i][7])
        };
      }
      break;
    }
  }

  var order = fetchOrder(orderId);
  if (!order) return { ok: false, error: 'No se encontró el pedido.' };

  if (order.status === 'cancelled') {
    return { ok: false, error: 'El pedido #' + order.number + ' fue cancelado y no puede entregarse.' };
  }
  if (order.payment_status === 'voided' || order.payment_status === 'refunded') {
    return { ok: false, error: 'El pedido #' + order.number + ' tiene un pago revertido o reembolsado.' };
  }

  return {
    ok:       true,
    entregado: false,
    order: {
      id:             String(order.id),
      number:         order.number,
      contact_name:   order.contact_name  || '',
      contact_email:  order.contact_email || '',
      total:          order.total || '',
      payment_status: order.payment_status,
      products: (order.products || []).map(function(p) {
        return { name: p.name, quantity: p.quantity };
      })
    }
  };
}

// Página inicial del QR — carga rápido con spinner, luego llama api_verificar
function paginaVerificar(orderId, token) {
  var webAppUrl    = ScriptApp.getService().getUrl();
  var apiUrl       = webAppUrl + '?action=api_verificar&id=' + encodeURIComponent(orderId) + '&token=' + encodeURIComponent(token);
  var portalUrl    = webAppUrl + '?action=portal&pass=' + CONFIG.STAFF_TOKEN;
  var extraParam   = '&token=' + encodeURIComponent(token);
  var staffOptions = Object.keys(CONFIG.STAFF_PINS).map(function(name) {
    return '<option value="' + name.trim() + '">' + name.trim() + '</option>';
  }).join('');

  var html =
    '<!DOCTYPE html><html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>' +
    'body{font-family:sans-serif;padding:20px;max-width:420px;margin:auto;color:#1a1a1a}' +
    'h2{margin-bottom:4px}' +
    '#loading{text-align:center;padding:60px 20px;color:#aaa}' +
    '.spinner{width:40px;height:40px;border:4px solid #eee;border-top-color:#00A650;border-radius:50%;animation:spin .8s linear infinite;margin:0 auto 16px}' +
    '.sc{width:22px;height:22px;border:3px solid #eee;border-top-color:#00A650;border-radius:50%;animation:spin .8s linear infinite;display:inline-block;vertical-align:middle;margin-right:8px}' +
    '@keyframes spin{to{transform:rotate(360deg)}}' +
    '.card{background:#f5f5f5;border-radius:10px;padding:16px;margin:16px 0}' +
    '.btn{display:block;width:100%;padding:18px;background:#00A650;color:#fff;border:none;border-radius:10px;font-size:20px;font-weight:bold;margin-top:24px;cursor:pointer;box-sizing:border-box}' +
    '.btn:disabled{background:#aaa;cursor:not-allowed}' +
    '.btn-portal{display:inline-block;padding:14px 28px;background:#1a1a6e;color:#fff;border-radius:10px;text-decoration:none;font-size:16px;font-weight:bold}' +
    'select,input{width:100%;padding:14px;font-size:16px;border:1px solid #ccc;border-radius:8px;margin:8px 0;box-sizing:border-box}' +
    '#err{color:red;font-weight:bold;text-align:center;display:none;margin:8px 0}' +
    '#sp-confirm{display:none;text-align:center;margin-top:12px;color:#555}' +
    '</style></head><body>' +
    '<div style="text-align:center;padding:16px 0 8px">' +
    '<img src="' + LOGO_B64 + '" style="max-width:200px;width:60%" alt="' + CONFIG.TIENDA_NOMBRE + '">' +
    '</div>' +
    '<div id="loading"><div class="spinner"></div><p>Cargando pedido...</p></div>' +
    '<div id="content" style="display:none"></div>' +
    '<script>' +
    'var WEB_URL="' + webAppUrl + '";' +
    'var ORDEN_ID="' + orderId + '";' +
    'var EXTRA="' + extraParam + '";' +
    'var PORTAL_URL="' + portalUrl + '";' +
    'function esc(s){return String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}' +
    'fetch("' + apiUrl + '",{redirect:"follow"})' +
    '  .then(function(r){return r.json();})' +
    '  .then(function(d){' +
    '    document.getElementById("loading").style.display="none";' +
    '    var c=document.getElementById("content");c.style.display="block";' +
    '    if(!d.ok){' +
    '      c.innerHTML=\'<div style="text-align:center"><p style="font-size:60px;margin:0">&#x274C;</p><h2 style="color:#E63946">Error</h2><p style="color:#555">\'+esc(d.error||"Error desconocido")+\'</p></div>\';' +
    '      return;' +
    '    }' +
    '    if(d.entregado){' +
    '      c.innerHTML=\'<div style="text-align:center">\'' +
    '        +\'<p style="font-size:60px;margin:0">&#x2705;</p>\'' +
    '        +\'<h2 style="color:#00A650">Pedido #\'+d.pedido_numero+\'</h2>\'' +
    '        +\'<p style="font-size:18px;color:#555">Este pedido ya fue entregado.</p>\'' +
    '        +\'<p style="color:#aaa;font-size:13px">\'+esc(d.estado)+\'</p>\'' +
    '        +\'<br><a href="\'+PORTAL_URL+\'" class="btn-portal">&#x2190; Volver al portal</a></div>\';' +
    '      return;' +
    '    }' +
    '    var o=d.order;' +
    '    var pagoOk=o.payment_status==="paid";' +
    '    var est=pagoOk' +
    '      ?\'<span style="color:#00A650;font-weight:bold;font-size:18px">&#x2713; PAGO CONFIRMADO</span>\'' +
    '      :\'<span style="color:#E63946;font-weight:bold;font-size:18px">&#x26A0; PAGO NO CONFIRMADO</span>\';' +
    '    var items=o.products.map(function(p){return\'<li style="padding:4px 0">\'+p.quantity+\'x <strong>\'+esc(p.name)+\'</strong></li>\';}).join("");' +
    '    var formHtml=pagoOk' +
    '      ?\'<label>Tu nombre:</label><select id="staff">' + staffOptions + '</select><label>Tu PIN:</label><input type="password" id="pin" placeholder="&bull;&bull;&bull;&bull;"><button id="btn" class="btn" onclick="confirmar()">&#x2713; Confirmar entrega</button><p id="sp-confirm"><span class="sc"></span>Registrando entrega...</p>\'' +
    '      :\'<p style="color:#E63946;text-align:center">No se puede entregar: pago pendiente.</p>\';' +
    '    c.innerHTML=\'<h2>Pedido #\'+o.number+\'</h2><p>\'+est+\'</p>\'' +
    '      +\'<div class="card"><p style="margin:0 0 4px"><strong>\'+esc(o.contact_name)+\'</strong></p>\'' +
    '      +\'<p style="margin:0;color:#888;font-size:14px">\'+esc(o.contact_email)+\'</p></div>\'' +
    '      +\'<div class="card"><ul style="margin:0;padding-left:18px">\'+items+\'</ul>\'' +
    '      +\'<p style="margin:12px 0 0"><strong>Total: \'+esc(o.total)+\'</strong></p></div>\'' +
    '      +\'<p id="err"></p>\'' +
    '      +formHtml;' +
    '  })' +
    '  .catch(function(){' +
    '    document.getElementById("loading").style.display="none";' +
    '    document.getElementById("content").innerHTML=\'<p style="color:red;text-align:center">Error de conexi\\u00f3n. Recarg\\u00e1 la p\\u00e1gina.</p>\';' +
    '    document.getElementById("content").style.display="block";' +
    '  });' +
    'function confirmar(){' +
    '  var staff=document.getElementById("staff").value;' +
    '  var pin=document.getElementById("pin").value;' +
    '  if(!pin){var el=document.getElementById("err");el.textContent="Ingres\\u00e1 tu PIN.";el.style.color="red";el.style.display="block";return;}' +
    '  var btn=document.getElementById("btn");' +
    '  btn.disabled=true;' +
    '  document.getElementById("sp-confirm").style.display="block";' +
    '  document.getElementById("err").style.display="none";' +
    '  var url=WEB_URL+"?action=api_entregar&id="+encodeURIComponent(ORDEN_ID)+"&staff="+encodeURIComponent(staff)+"&pin="+encodeURIComponent(pin)+EXTRA;' +
    '  fetch(url,{redirect:"follow"})' +
    '    .then(function(r){return r.json();})' +
    '    .then(function(data){' +
    '      if(data.ok){' +
    '        document.body.innerHTML=\'<div style="font-family:sans-serif;text-align:center;padding:48px 24px;max-width:420px;margin:auto">\'' +
    '          +\'<p style="font-size:72px;margin:0">&#x2705;</p>\'' +
    '          +\'<h2 style="color:#00A650;margin:16px 0 8px">\\u00a1Entrega registrada!</h2>\'' +
    '          +\'<p style="color:#555;margin-bottom:32px">El pedido fue marcado como entregado.</p>\'' +
    '          +\'<a href="\'+PORTAL_URL+\'" style="display:inline-block;padding:14px 28px;background:#1a1a6e;color:#fff;border-radius:10px;text-decoration:none;font-size:16px;font-weight:bold">\\u2190 Volver al portal</a></div>\';' +
    '      }else{' +
    '        var el=document.getElementById("err");' +
    '        el.textContent=data.error||"Error desconocido.";' +
    '        el.style.color="red";el.style.display="block";' +
    '        btn.disabled=false;' +
    '        document.getElementById("sp-confirm").style.display="none";' +
    '      }' +
    '    })' +
    '    .catch(function(){' +
    '      var el=document.getElementById("err");' +
    '      el.textContent="Error de red. Intent\\u00e1 de nuevo.";' +
    '      el.style.color="red";el.style.display="block";' +
    '      btn.disabled=false;' +
    '      document.getElementById("sp-confirm").style.display="none";' +
    '    });' +
    '}' +
    '<\/script>' +
    '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function paginaPortalStaff(mostrarExito) {
  var sheet     = getSheet();
  var data      = sheet.getDataRange().getValues();
  var webAppUrl = ScriptApp.getService().getUrl();

  var pendientes = [];
  var entregados = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var entregado = String(row[7]).indexOf('ENTREGADO') !== -1;
    var item = {
      id:        String(row[0]),
      numero:    row[1],
      nombre:    row[2],
      productos: row[4],
      total:     row[5],
      fecha:     row[6] ? new Date(row[6]).toLocaleString('es-AR') : '',
      estado:    row[7],
      urlVer:    webAppUrl + '?action=verificar&id=' + row[0] + '&token=' + generarToken(String(row[0])),
      urlAdmin:  webAppUrl + '?action=entregar_admin&id=' + row[0] + '&pass=' + CONFIG.STAFF_TOKEN
    };
    if (entregado) entregados.push(item); else pendientes.push(item);
  }

  function filaHtml(item, esEntregado) {
    var bg    = esEntregado ? '#f5fff8' : '#fff';
    var badge = esEntregado
      ? '<span style="background:#e0f5e0;color:#00A650;padding:2px 8px;border-radius:12px;font-size:12px">✓ entregado</span>'
      : '<span style="background:#fff3cd;color:#856404;padding:2px 8px;border-radius:12px;font-size:12px">pendiente</span>';

    var acciones = '';
    if (!esEntregado) {
      acciones =
        '<a href="' + item.urlVer + '" style="color:#3483FA;font-size:13px;text-decoration:none;margin-right:10px">Ver →</a>' +
        '<a href="' + item.urlAdmin + '" style="color:#fff;background:#00A650;font-size:12px;text-decoration:none;padding:4px 10px;border-radius:6px">Entregar</a>';
    }

    return '<tr style="background:' + bg + ';border-bottom:1px solid #eee">' +
      '<td style="padding:12px 8px">' +
        '<strong>Pedido #' + item.numero + '</strong><br>' +
        '<span style="color:#555;font-size:13px">' + item.nombre + '</span><br>' +
        '<span style="color:#888;font-size:12px">' + item.productos + '</span>' +
      '</td>' +
      '<td style="padding:12px 8px;text-align:right;white-space:nowrap">' +
        '<strong>' + item.total + '</strong><br>' +
        badge + '<br>' +
        '<div style="margin-top:6px">' + acciones + '</div>' +
      '</td>' +
      '</tr>';
  }

  var filasPendientes = pendientes.length > 0
    ? pendientes.map(function(i) { return filaHtml(i, false); }).join('')
    : '<tr><td colspan="2" style="padding:20px;text-align:center;color:#aaa">Sin pedidos pendientes</td></tr>';

  var filasEntregados = entregados.slice(-10).reverse()
    .map(function(i) { return filaHtml(i, true); }).join('');

  var html =
    '<html><head>' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>' +
    'body{font-family:sans-serif;padding:0;margin:0;background:#f5f5f5;color:#1a1a1a}' +
    '.header{background:#1a1a1a;color:#fff;padding:16px 20px;position:sticky;top:0;z-index:10}' +
    '.header h1{margin:0;font-size:20px}' +
    '.header p{margin:4px 0 0;font-size:13px;color:#aaa}' +
    '.section{padding:16px 12px 4px;font-size:11px;font-weight:bold;color:#888;text-transform:uppercase;letter-spacing:1px}' +
    'table{width:100%;border-collapse:collapse;background:#fff}' +
    '.badge-count{background:#E63946;color:#fff;border-radius:12px;padding:2px 8px;font-size:13px;margin-left:8px}' +
    '</style></head><body>' +
    '<div class="header">' +
    '<h1>' + CONFIG.TIENDA_NOMBRE + '</h1>' +
    '<p>Portal de retiros — ' + new Date().toLocaleDateString('es-AR') + '</p>' +
    '</div>' +
    (mostrarExito ? '<div style="background:#00A650;color:#fff;padding:14px 20px;text-align:center;font-weight:bold;font-size:16px">&#x2705; Pedido marcado como entregado</div>' : '') +
    '<div class="section">Pendientes de entrega' +
    (pendientes.length > 0 ? '<span class="badge-count">' + pendientes.length + '</span>' : '') +
    '</div>' +
    '<table>' + filasPendientes + '</table>' +
    '<div class="section">Últimas entregas</div>' +
    '<table>' + (filasEntregados || '<tr><td style="padding:20px;text-align:center;color:#aaa">Sin entregas registradas</td></tr>') + '</table>' +
    '<div style="height:40px"></div>' +
    '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function paginaConfirmar(orderId, token, errorMsg, adminPass) {
  var webAppUrl    = ScriptApp.getService().getUrl();
  var portalUrl    = webAppUrl + '?action=portal&pass=' + CONFIG.STAFF_TOKEN + '&entregado=1';
  var staffOptions = Object.keys(CONFIG.STAFF_PINS).map(function(name) {
    return '<option value="' + name.trim() + '">' + name.trim() + '</option>';
  }).join('');

  // El endpoint AJAX que verifica PIN y registra la entrega, devuelve JSON
  var extraParam = adminPass
    ? '&pass=' + encodeURIComponent(adminPass)
    : '&token=' + encodeURIComponent(token);

  var html =
    '<!DOCTYPE html><html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>' +
    'body{font-family:sans-serif;padding:24px;max-width:420px;margin:auto;color:#1a1a1a}' +
    'h2{text-align:center}' +
    'select,input{width:100%;padding:14px;font-size:16px;border:1px solid #ccc;border-radius:8px;margin:8px 0;box-sizing:border-box}' +
    '.btn{display:block;width:100%;padding:16px;background:#00A650;color:#fff;border:none;border-radius:10px;font-size:18px;font-weight:bold;margin-top:16px;cursor:pointer}' +
    '.btn:disabled{background:#aaa;cursor:not-allowed}' +
    '#error{color:red;font-weight:bold;text-align:center;display:none;margin:8px 0}' +
    '#loading{text-align:center;display:none;color:#555;margin-top:16px}' +
    '</style></head><body>' +
    '<div style="text-align:center;padding:16px 0 8px 0">' +
    '<img src="' + LOGO_B64 + '" style="max-width:200px;width:60%" alt="' + CONFIG.TIENDA_NOMBRE + '">' +
    '</div>' +
    '<h2>Confirmar entrega</h2>' +
    '<p style="text-align:center;color:#555">Pedido #' + getOrderNumber(orderId) + '</p>' +
    '<p id="error">' + (errorMsg || '') + '</p>' +
    '<label>Tu nombre:</label>' +
    '<select id="staff">' + staffOptions + '</select>' +
    '<label>Tu PIN:</label>' +
    '<input type="password" id="pin" placeholder="••••">' +
    '<button id="btn" class="btn" onclick="confirmar()">&#x2713; Confirmar entrega</button>' +
    '<p id="loading">Registrando entrega...</p>' +
    '<script>' +
    'var WEB_URL="' + webAppUrl + '";' +
    'var PORTAL_URL="' + portalUrl + '";' +
    'var ORDEN_ID="' + orderId + '";' +
    'var EXTRA="' + extraParam + '";' +
    (errorMsg ? 'document.getElementById("error").style.display="block";' : '') +
    'function confirmar(){' +
    '  var staff=document.getElementById("staff").value;' +
    '  var pin=document.getElementById("pin").value;' +
    '  if(!pin){var el=document.getElementById("error");el.textContent="Ingres\\u00e1 tu PIN.";el.style.display="block";return;}' +
    '  var btn=document.getElementById("btn");' +
    '  btn.disabled=true;' +
    '  document.getElementById("loading").style.display="block";' +
    '  document.getElementById("error").style.display="none";' +
    '  var url=WEB_URL+"?action=api_entregar&id="+encodeURIComponent(ORDEN_ID)+"&staff="+encodeURIComponent(staff)+"&pin="+encodeURIComponent(pin)+EXTRA;' +
    '  fetch(url,{redirect:"follow"})' +
    '    .then(function(r){return r.json();})' +
    '    .then(function(data){' +
    '      if(data.ok){' +
    '        document.body.innerHTML=' +
    '          \'<div style="font-family:sans-serif;text-align:center;padding:48px 24px;max-width:420px;margin:auto">\'' +
    '          +\'<p style="font-size:72px;margin:0">\\u2705</p>\'' +
    '          +\'<h2 style="color:#00A650;margin:16px 0 8px">\\u00a1Entrega registrada!</h2>\'' +
    '          +\'<p style="color:#555;margin-bottom:32px">El pedido fue marcado como entregado.</p>\'' +
    '          +\'<a href="\'+PORTAL_URL+\'" style="display:inline-block;padding:14px 28px;background:#00A650;color:#fff;border-radius:10px;text-decoration:none;font-size:16px;font-weight:bold">\\u2190 Volver al portal</a>\'' +
    '          +\'</div>\';' +
    '      } else {' +
    '        var el=document.getElementById("error");' +
    '        el.textContent=data.error||"Error desconocido.";' +
    '        el.style.display="block";' +
    '        btn.disabled=false;' +
    '        document.getElementById("loading").style.display="none";' +
    '      }' +
    '    })' +
    '    .catch(function(e){' +
    '      document.getElementById("error").textContent="Error de red. Intent\\u00e1 de nuevo.";' +
    '      document.getElementById("error").style.display="block";' +
    '      btn.disabled=false;' +
    '      document.getElementById("loading").style.display="none";' +
    '    });' +
    '}' +
    '<\/script>' +
    '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function paginaEntregado(orderId, staffName) {
  var webAppUrl = ScriptApp.getService().getUrl();
  var portalUrl = webAppUrl + '?action=portal&pass=' + CONFIG.STAFF_TOKEN + '&entregado=1';
  var html =
    '<!DOCTYPE html><html><head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<meta http-equiv="refresh" content="0;url=' + portalUrl + '">' +
    '</head><body>' +
    '<p style="font-family:sans-serif;padding:24px;text-align:center">' +
    '&#x2705; Entregado! <a href="' + portalUrl + '">Volver al portal</a>' +
    '</p>' +
    '</body></html>';
  return HtmlService.createHtmlOutput(html);
}

// ──────────────────────────────────────────────────────────
// TOKEN DE SEGURIDAD
// ──────────────────────────────────────────────────────────

function generarToken(orderId) {
  var raw  = String(orderId) + CONFIG.SECRET;
  var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw);
  return hash.slice(0, 8)
    .map(function(b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); })
    .join('')
    .substring(0, 12);
}

function verificarToken(orderId, token) {
  return generarToken(orderId) === token;
}

function getOrderNumber(orderId) {
  var data = getSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(orderId)) return data[i][1];
  }
  return orderId; // fallback: si no lo encuentra, muestra el ID interno
}

// ──────────────────────────────────────────────────────────
// ENTREGA
// ──────────────────────────────────────────────────────────

// Lógica de verificación de PIN + entrega — usada por api_entregar (AJAX)
function apiEntregar(orderId, staffName, pin, adminPass, token, notas) {
  // Validar acceso: o viene del portal (adminPass) o tiene un token QR válido
  if (adminPass) {
    if (adminPass !== CONFIG.STAFF_TOKEN) return { ok: false, error: 'Acceso no autorizado.' };
  } else {
    if (!orderId || !verificarToken(orderId, token)) return { ok: false, error: 'QR inválido o expirado.' };
  }
  var pins = CONFIG.STAFF_PINS;
  var staffKey = staffName ? staffName.trim() : '';
  // Buscar clave en STAFF_PINS ignorando espacios en los keys
  var pinEsperado = null;
  var keys = Object.keys(pins);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].trim() === staffKey) { pinEsperado = pins[keys[i]]; break; }
  }
  if (!pinEsperado || pinEsperado !== pin) {
    return { ok: false, error: 'PIN incorrecto. Intentá de nuevo.' };
  }
  if (CONFIG.RETIRO_PINS.indexOf(pin) === -1) {
    return { ok: false, error: 'Tu acceso es solo al dashboard de ventas. No tenés permiso para confirmar entregas.' };
  }
  try {
    marcarEntregado(orderId, staffKey, notas || '');
  } catch(err) {
    Logger.log('Error al marcar entregado: ' + err);
    return { ok: false, error: 'No se pudo registrar la entrega: ' + err.message };
  }
  return { ok: true };
}

function procesarEntrega(orderId, staffName, pin) {
  var res = apiEntregar(orderId, staffName, pin, null, generarToken(orderId));
  if (!res.ok) return paginaConfirmar(orderId, generarToken(orderId), res.error);
  return paginaPortalStaff(true);
}

function paginaError(mensaje) {
  return HtmlService.createHtmlOutput(
    '<html><head><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{font-family:sans-serif;padding:32px 24px;max-width:420px;margin:auto;text-align:center;color:#1a1a1a}</style>' +
    '</head><body>' +
    '<p style="font-size:60px;margin:0">&#x26A0;</p>' +
    '<h2 style="color:#E63946">Error</h2>' +
    '<p style="color:#555">' + mensaje + '</p>' +
    '<p style="font-size:13px;color:#aaa">Contactá al administrador.</p>' +
    '</body></html>'
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function marcarFulfillmentTiendanube(orderId) {
  try {
    var url = 'https://api.tiendanube.com/v1/' + CONFIG.STORE_ID + '/orders/' + orderId + '/fulfill';
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ notify_customer: true }),
      headers: {
        'Authentication': 'bearer ' + CONFIG.API_TOKEN,
        'User-Agent': CONFIG.TIENDA_NOMBRE + ' (retiro-automatico)',
      },
      muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    Logger.log('TiendaNube fulfill → HTTP ' + code + ' | ' + resp.getContentText().substring(0, 300));
  } catch(err) {
    Logger.log('Error fulfillment TiendaNube: ' + err);
  }
}

function enviarEmailEntrega(email, nombre, numeroPedido, productos) {
  try {
    var asunto = '✅ Tu pedido #' + numeroPedido + ' fue entregado — Diente de León';
    var cuerpo =
      'Hola ' + nombre + ',\n\n' +
      'Tu pedido #' + numeroPedido + ' fue entregado correctamente. 🎉\n\n' +
      'Productos:\n' + productos + '\n\n' +
      'Muchas gracias por tu compra en la tienda Diente de León.\n\n' +
      'Comisión de Recursos\nEscuela Diente de León';
    GmailApp.sendEmail(email, asunto, cuerpo, { from: 'tiendavirtual.dientedeleon@gmail.com' });
    Logger.log('Email entrega enviado a ' + email);
  } catch(err) {
    Logger.log('Error email entrega: ' + err);
  }
}

// ──────────────────────────────────────────────────────────
// UTILIDADES
// ──────────────────────────────────────────────────────────

function testConPedido() {
  var ORDER_ID_REAL = 'PEGA_UN_ID_DE_PEDIDO_REAL'; // ← reemplazá
  var order = fetchOrder(ORDER_ID_REAL);
  if (!order) { Logger.log('No se encontró el pedido'); return; }
  Logger.log('Pedido encontrado: #' + order.number + ' — ' + order.contact_email);
  procesarPedido(order);
}

// Trae todas las órdenes pagadas de Tiendanube con paginación
function fetchAllPaidOrders() {
  var allOrders = [];
  var page = 1;
  while (true) {
    var url = 'https://api.tiendanube.com/v1/' + CONFIG.STORE_ID +
              '/orders?payment_status=paid&per_page=200&page=' + page;
    var resp = UrlFetchApp.fetch(url, {
      headers: {
        'Authentication': 'bearer ' + CONFIG.API_TOKEN,
        'User-Agent': CONFIG.TIENDA_NOMBRE + ' (retiro-automatico)'
      },
      muteHttpExceptions: true
    });
    var batch = JSON.parse(resp.getContentText());
    if (!batch || batch.length === 0) break;
    allOrders = allOrders.concat(batch);
    Logger.log('Página ' + page + ': ' + batch.length + ' pedidos');
    if (batch.length < 200) break;
    page++;
  }
  // Ordenar de más viejo a más nuevo por ID (IDs son secuenciales en TiendaNube)
  allOrders.sort(function(a, b) { return a.id - b.id; });
  return allOrders;
}

// Borra la pestaña Ventas y la rellena con todos los pedidos pagados históricos
function reprocesarVentas() {
  Logger.log('Iniciando reproceso de Ventas...');

  // Limpiar hoja (dejando el encabezado)
  var sheet = getVentasSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  var orders = fetchAllPaidOrders();
  Logger.log('Total pedidos pagados: ' + orders.length);

  // Diagnóstico: loguear el primer pedido para ver la estructura de precios
  if (orders.length > 0) {
    var sample = orders[0];
    Logger.log('SAMPLE pedido #' + sample.number + ' total=' + sample.total + ' subtotal=' + sample.subtotal);
    (sample.products || []).forEach(function(p, i) {
      Logger.log('  producto[' + i + ']: price=' + p.price + ' unit_price=' + p.unit_price + ' qty=' + p.quantity + ' name=' + JSON.stringify(p.name));
    });
  }

  var registrados = 0;
  orders.forEach(function(order) {
    if (order.status === 'cancelled') return;
    registrarEnVentas(order);
    registrados++;
  });

  Logger.log('Listo. ' + registrados + ' pedidos → ' + registrados + '+ filas en Ventas.');
}

function reprocesarPendientes() {
  var url  = 'https://api.tiendanube.com/v1/' + CONFIG.STORE_ID + '/orders?payment_status=paid&per_page=50';
  var resp = UrlFetchApp.fetch(url, {
    headers: {
      'Authentication': 'bearer ' + CONFIG.API_TOKEN,
      'User-Agent': CONFIG.TIENDA_NOMBRE + ' (retiro-automatico)'
    },
    muteHttpExceptions: true
  });
  var orders = JSON.parse(resp.getContentText());
  Logger.log('Pedidos encontrados: ' + orders.length);
  orders.forEach(function(order) {
    if (order.status !== 'cancelled' && !yaFueEnviado(String(order.id))) {
      Logger.log('Procesando pedido #' + order.number + ' — ' + order.contact_email);
      procesarPedido(order);
    } else {
      Logger.log('Saltando pedido #' + order.number + ' (cancelado o ya enviado)');
    }
  });
  Logger.log('Listo.');
}

// ──────────────────────────────────────────────────────────
// REPROCESO — Marcar en TiendaNube todos los pedidos ya entregados en Sheets
// Ejecutar UNA SOLA VEZ manualmente desde el editor de GAS
// ──────────────────────────────────────────────────────────
function reprocesarEntregasTiendaNube() {
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();
  var procesados = 0;
  var errores    = 0;

  for (var i = 1; i < data.length; i++) {
    var orderId = String(data[i][0]);
    var estado  = String(data[i][7]);

    if (estado.indexOf('ENTREGADO') === -1) continue;
    if (!orderId || orderId === '') continue;

    Logger.log('Procesando orden ' + orderId + ' (fila ' + (i+1) + ')...');
    try {
      marcarFulfillmentTiendanube(orderId);
      procesados++;
      Utilities.sleep(500); // pausa entre llamadas para no saturar la API
    } catch(err) {
      Logger.log('ERROR en orden ' + orderId + ': ' + err);
      errores++;
    }
  }

  Logger.log('=== REPROCESO COMPLETO === Procesados: ' + procesados + ' | Errores: ' + errores);
}

function registrarWebhook() {
  var url = 'https://api.tiendanube.com/v1/' + CONFIG.STORE_ID + '/webhooks';
  var payload = JSON.stringify({
    event: 'order/paid',
    url: 'https://script.google.com/macros/s/AKfycbwt8mRYjjFZtsmFPw0JYcMpCkZcOt9J7ALFVkPByqQ8NM2NvCaYF1onagU_ag0a2ziksg/exec'
  });
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: payload,
    headers: {
      'Authentication': 'bearer ' + CONFIG.API_TOKEN,
      'User-Agent': CONFIG.TIENDA_NOMBRE + ' (retiro-automatico)'
    },
    muteHttpExceptions: true
  });
  Logger.log(resp.getResponseCode() + ' → ' + resp.getContentText());
}

// ──────────────────────────────────────────────────────────
// WHATSAPP — ENVÍO DE PINs A REFERENTES
// ──────────────────────────────────────────────────────────

function enviarPinsReferentes() {
  var GREEN_INSTANCE = '7107556439';
  var GREEN_TOKEN    = '7f7e2a357d2a4f0faa7edc4729f4c89f491b1420ef19462783';
  var DASHBOARD_URL  = 'https://escuela-dandelion.github.io/Comision-Recursos/dashboard-referentes.html';

  var referentes = [
    { nombre: 'Ines',    grado: '4to',    pin: 'PIR26',  tel: '5493512112050' },
    { nombre: 'Yuliana', grado: 'Jardín', pin: 'PYL26',  tel: '5493513645612' },
    { nombre: 'Maria',   grado: 'Jardín', pin: 'PMM26',  tel: '5493516865078' },
    { nombre: 'Delfina', grado: '6to',    pin: 'PDS26',  tel: '5493543512538' },
    { nombre: 'Dolo',    grado: '2do',    pin: 'PDD26',  tel: '5493516523262' },
    { nombre: 'Gaby',    grado: '5to',    pin: 'PGD26',  tel: '5493517076221' },
    { nombre: 'Juan',    grado: '8vo',    pin: 'PJD26',  tel: '5493515514893' },
    { nombre: 'Karina',  grado: '7mo',    pin: 'PK26',   tel: '5493512130302' },
    { nombre: 'Monica',  grado: null,     pin: 'PMC26',  tel: '5493516808115' },
    { nombre: 'Paola',   grado: '2do',    pin: 'PP26',   tel: '5493516166134' },
    { nombre: 'Pia',     grado: '11vo',   pin: 'PPI26',  tel: '5493516602546' },
    { nombre: 'Sabi',    grado: '9no',    pin: 'PSA26',  tel: '5493517888479' },
    { nombre: 'Yana',    grado: '5to',    pin: 'PYA26',  tel: '5493512114209' },
    { nombre: 'Elina',   grado: '3ro',    pin: 'PEL26',  tel: '5493512032487' },
    { nombre: 'Claudio', grado: null,     pin: 'PCL26',  tel: '5493516319266' }
  ];

  referentes.forEach(function(r) {
    var saludo = r.grado
      ? 'Hola ' + r.nombre + '! Sos *Referente de ' + r.grado + ' grado* en la tienda Diente de León. Te comparto tu PIN:\n\n'
      : 'Hola ' + r.nombre + '! Te comparto tu PIN para la tienda:\n\n';

    var mensaje =
      '🌼 *Diente de León — Tu PIN personal*\n\n' +
      saludo +
      '🔐 *Tu PIN: ' + r.pin + '*\n\n' +
      'Lo vas a usar para dos cosas:\n\n' +
      '📦 *Confirmar entregas:* cuando una familia venga a retirar su pedido, escaneás el QR que te muestra el comprador y confirmás la entrega con tu PIN. De esta forma tendremos control sobre quién entregó qué mercadería.\n\n' +
      '📊 *Dashboard de ventas:* podés ver en tiempo real las ventas de tu grado en:\n' +
      DASHBOARD_URL + '\n' +
      '¡Animate a impulsar y promover más compras!\n\n' +
      '¡Guardá este mensaje! 🙌';

    if (!r.tel) {
      Logger.log('Sin teléfono — saltando ' + r.nombre);
      return;
    }

    UrlFetchApp.fetch(
      'https://api.green-api.com/waInstance' + GREEN_INSTANCE + '/sendMessage/' + GREEN_TOKEN, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({ chatId: r.tel + '@c.us', message: mensaje })
    });

    Logger.log('Enviado a ' + r.nombre + ' (' + r.tel + ')');
    Utilities.sleep(1500); // pausa entre mensajes para no saturar la API
  });

  Logger.log('Listo. ' + referentes.length + ' mensajes enviados.');
}

function enviarPinsStaff() {
  var staff = [
    { nombre: 'Claudio',   pin: 'PCL26',  tel: '5493516319266' },
    { nombre: 'Anabella',  pin: 'PAG26',  tel: '5493513915935' }
  ];
  staff.forEach(function(r) { _enviarPinDashboard(r.nombre, r.pin, r.tel); });
  Logger.log('Listo. ' + staff.length + ' mensajes enviados.');
}

// ──────────────────────────────────────────────────────────
// Envío individual — para cuando se suma una persona nueva
// Ejemplo de uso: enviarPinNuevoMiembro('Anabella', 'PAG26', '5493513915935')
// ──────────────────────────────────────────────────────────
function enviarPinNuevoMiembro() {
  // ← Modificar estos tres valores antes de ejecutar
  var nombre = 'Anabella';
  var pin    = 'PAG26';
  var tel    = '5493513915935';
  _enviarPinDashboard(nombre, pin, tel);
  Logger.log('PIN enviado a ' + nombre);
}

function _enviarPinDashboard(nombre, pin, tel) {
  var GREEN_INSTANCE  = '7107556439';
  var GREEN_TOKEN     = '7f7e2a357d2a4f0faa7edc4729f4c89f491b1420ef19462783';
  var URL_GENERAL     = 'https://escuela-dandelion.github.io/Comision-Recursos/dashboard-general.html';
  var URL_REFERENTES  = 'https://escuela-dandelion.github.io/Comision-Recursos/dashboard-referentes.html';
  var URL_WIKI        = 'https://escuela-dandelion.github.io/Comision-Recursos/wiki-comision-recursos.html';

  var mensaje =
    '🌼 *Bienvenida a la Comisión de Recursos — Diente de León!*\n\n' +
    'Hola ' + nombre + '! Te compartimos tu acceso al sistema de gestión de la tienda.\n\n' +
    '🔐 *Tu PIN personal: ' + pin + '*\n' +
    '_Guardalo, lo vas a necesitar para ingresar a todo._\n\n' +
    '📊 *Dashboard General*\n' +
    'Métricas de ventas, evolución mensual, productos y rentabilidad:\n' +
    URL_GENERAL + '\n\n' +
    '🎓 *Dashboard Referentes*\n' +
    'Ventas por grado, quién compró qué y lista de retiro:\n' +
    URL_REFERENTES + '\n\n' +
    '📖 *Wiki de la Comisión*\n' +
    'Procesos, roles, circuito logístico y más:\n' +
    URL_WIKI + '\n\n' +
    '¡Guardá este mensaje! 🙌';

  UrlFetchApp.fetch(
    'https://api.green-api.com/waInstance' + GREEN_INSTANCE + '/sendMessage/' + GREEN_TOKEN, {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify({ chatId: tel + '@c.us', message: mensaje })
  });
  Utilities.sleep(1500);
}

// ──────────────────────────────────────────────────────────
// DASHBOARD API
// ──────────────────────────────────────────────────────────

function apiDashboard(pin, email) {
  var pinValido = false;

  // Bypass por email (sesión Google del admin)
  if (email) {
    var authUser = AUTH_EMAILS[email.toLowerCase().trim()];
    if (authUser) pinValido = true;
  }

  // Validar PIN contra STAFF_PINS y DASHBOARD_PINS
  if (!pinValido) {
    var allPins = [CONFIG.STAFF_PINS, CONFIG.DASHBOARD_PINS];
    for (var g = 0; g < allPins.length; g++) {
      var keys = Object.keys(allPins[g]);
      for (var i = 0; i < keys.length; i++) {
        if (allPins[g][keys[i]] === pin) { pinValido = true; break; }
      }
      if (pinValido) break;
    }
  }

  if (!pinValido) return { ok: false, error: 'PIN incorrecto.' };

  var sheet = getVentasSheet();
  var data  = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { ok: true, rows: [], meses_todos: [], total_pedidos: 0, total_pesos: 0, total_costo: 0, actualizado: new Date().toLocaleString('es-AR') };
  }

  // Columnas: 0=Fecha,1=Pedido#,2=IDInt,3=Nombre,4=Email,5=Producto,6=SKU,7=Cantidad,8=PrecioU,9=TotLinea,10=TotPedido,11=Comentarios,12=Marca,13=CostoUnitario
  var rows       = [];
  var mesesSet   = {};
  var pedidosSet = {};
  var totalPesos = 0;
  var totalCosto = 0;

  for (var r = 1; r < data.length; r++) {
    var row        = data[r];
    var fecha      = row[0] ? new Date(row[0]) : null;
    var pedido     = String(row[1] || '');
    var nombre     = String(row[3] || 'Sin nombre');
    var email      = String(row[4] || '');
    var producto   = String(row[5] || '');
    var cantidad   = parseInt(row[7]) || 1;
    var totalLinea  = parseFloat(row[9])  || 0;
    var totalPedido = parseFloat(row[10]) || 0;
    var comentario  = String(row[11] || '').trim();
    var marca       = String(row[12] || '');
    var mes         = fecha ? (fecha.getFullYear() + '-' + ('0' + (fecha.getMonth()+1)).slice(-2)) : '';
    var grado       = comentario || '(Sin observaciones)';

    var costoUnit  = parseFloat(row[13]) || 0;
    var costoLinea = costoUnit * cantidad;
    totalCosto    += costoLinea;

    if (mes) mesesSet[mes] = true;
    if (pedido && !pedidosSet[pedido]) {
      totalPesos += totalPedido;
    }
    if (pedido) pedidosSet[pedido] = true;

    rows.push({
      fecha:       fecha ? fecha.toLocaleDateString('es-AR') : '',
      mes:         mes,
      pedido:      pedido,
      familia:     nombre,
      email:       email,
      producto:    producto,
      marca:       marca,
      cantidad:    cantidad,
      total_linea: Math.round(totalLinea * 100) / 100,
      costo_linea: Math.round(costoLinea * 100) / 100,
      grado:       grado
    });
  }

  return {
    ok:            true,
    rows:          rows,
    meses_todos:   Object.keys(mesesSet).sort(),
    total_pedidos: Object.keys(pedidosSet).length,
    total_pesos:   Math.round(totalPesos * 100) / 100,
    total_costo:   Math.round(totalCosto * 100) / 100,
    actualizado:   new Date().toLocaleString('es-AR')
  };
}

// ═══════════════════════════════════════════════════════════
//  AUTH — admin.html
// ═══════════════════════════════════════════════════════════

// Lista blanca de emails con su rol
var AUTH_EMAILS = {
  'robertson.ine@gmail.com':    { role: 'ADMIN', name: 'Inés' },
  'longhi.yuliana@gmail.com':        { role: 'ADMIN', name: 'Yuliana' },
  'yuliana.longhi@dandelion.edu.ar': { role: 'ADMIN', name: 'Yuliana' },
  'martinimaria39@gmail.com':   { role: 'FGP',   name: 'Maria' },
  'lourdelcastillo@gmail.com':       { role: 'FGP',   name: 'Luli' },
  'belen.ravaza@dandelion.edu.ar':   { role: 'FGP',   name: 'Belén' }
  // Agregar más emails de Staff acá con role: 'STAFF'
};

// Lista blanca de teléfonos de referentes (sin el +, formato internacional)
var AUTH_REFERENTES = {
  '5493513645612': { role: 'REFERENTE', name: 'Yuliana' },
  '5493516865078': { role: 'REFERENTE', name: 'Maria'   },
  '5493512112050': { role: 'REFERENTE', name: 'Inés'    },
  '5493543512538': { role: 'REFERENTE', name: 'Delfina' },
  '5493516523262': { role: 'REFERENTE', name: 'Dolo'    },
  '5493517076221': { role: 'REFERENTE', name: 'Gaby'    },
  '5493515514893': { role: 'REFERENTE', name: 'Juan'    },
  '5493512130302': { role: 'REFERENTE', name: 'Karina'  },
  '5493516808115': { role: 'REFERENTE', name: 'Monica'  },
  '5493516166134': { role: 'REFERENTE', name: 'Paola'   },
  '5493516602546': { role: 'REFERENTE', name: 'Pia'     },
  '5493517888479': { role: 'REFERENTE', name: 'Sabi'    },
  '5493512114209': { role: 'REFERENTE', name: 'Yana'    },
  '5493512032487': { role: 'REFERENTE', name: 'Elina'   },
  '5493516319266': { role: 'REFERENTE', name: 'Claudio' }
};

// Verificar Google Sign-In — devuelve {ok, role, name}
function apiVerificarEmail(email) {
  email = email.toLowerCase().trim();
  var user = AUTH_EMAILS[email];
  if (!user) return { ok: false, error: 'Email no autorizado' };
  return { ok: true, role: user.role, name: user.name };
}

// Generar y enviar OTP por WhatsApp
function apiSendOTP(tel) {
  tel = tel.replace(/\D/g, ''); // solo dígitos
  if (!AUTH_REFERENTES[tel]) return { ok: false, error: 'Número no registrado' };

  var code = String(Math.floor(100000 + Math.random() * 900000));
  var expira = new Date().getTime() + 5 * 60 * 1000; // 5 minutos

  // Guardar en PropertiesService (temporal, expira)
  PropertiesService.getScriptProperties().setProperty('otp_' + tel, code + '|' + expira);

  // Enviar por WhatsApp
  try {
    UrlFetchApp.fetch(
      'https://api.green-api.com/waInstance' + CONFIG.GREEN_INSTANCE + '/sendMessage/' + CONFIG.GREEN_TOKEN,
      {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify({
          chatId: tel + '@c.us',
          message: '🌼 *Diente de León*\n\nTu código de acceso es: *' + code + '*\n\nExpira en 5 minutos.'
        })
      }
    );
  } catch(err) {
    Logger.log('Error enviando OTP: ' + err);
    return { ok: false, error: 'Error enviando WhatsApp' };
  }

  return { ok: true };
}

// Verificar OTP ingresado
function apiVerifyOTP(tel, code) {
  tel = tel.replace(/\D/g, '');
  var stored = PropertiesService.getScriptProperties().getProperty('otp_' + tel);
  if (!stored) return { ok: false, error: 'Código no encontrado o expirado' };

  var parts  = stored.split('|');
  var savedCode  = parts[0];
  var expira     = Number(parts[1]);

  if (new Date().getTime() > expira) {
    PropertiesService.getScriptProperties().deleteProperty('otp_' + tel);
    return { ok: false, error: 'Código expirado' };
  }

  if (code !== savedCode) return { ok: false, error: 'Código incorrecto' };

  // Limpiar OTP usado
  PropertiesService.getScriptProperties().deleteProperty('otp_' + tel);

  var user = AUTH_REFERENTES[tel];
  return { ok: true, role: user.role, name: user.name };
}

// ── RETIROS CON ESTADO ENTREGADO/PENDIENTE ───────────────────
function getRetirosData(estadoFiltro, diasFiltro) {
  var ss    = SpreadsheetApp.openById(CONFIG.VENTAS_SHEET_ID);
  var sheet = ss.getSheetByName('Retiros');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row    = data[i];
    if (!row[0]) continue;
    var estado        = String(row[7] || '');
    var entregado     = estado.indexOf('ENTREGADO') !== -1;
    var fechaEntrega  = entregado ? estado.replace('ENTREGADO ✓ — ', '').trim() : '';
    var staff         = String(row[8] || '').replace('.', ' ').trim();
    var fechaPedido   = '';
    try { fechaPedido = row[6] ? Utilities.formatDate(new Date(row[6]), 'America/Argentina/Cordoba', 'dd/MM/yyyy') : ''; } catch(e) {}
    // Filtro por estado
    if (estadoFiltro && estadoFiltro !== entregado ? 'Entregado' : 'Pendiente') {
      if (estadoFiltro === 'Pendiente' && entregado) continue;
      if (estadoFiltro === 'Entregado' && !entregado) continue;
    }
    // Filtro por días (solo para Entregado)
    if (diasFiltro > 0 && entregado) {
      var fechaRow = row[6] ? new Date(row[6]) : null;
      if (fechaRow) {
        var limite = new Date();
        limite.setDate(limite.getDate() - diasFiltro);
        if (fechaRow < limite) continue;
      }
    }
    rows.push({
      id:           String(row[0]),
      numero:       row[1],
      nombre:       String(row[2] || ''),
      email:        String(row[3] || ''),
      productos:    String(row[4] || ''),
      total:        String(row[5] || ''),
      fecha:        fechaPedido,
      estado:       entregado ? 'Entregado' : 'Pendiente',
      fechaEntrega: fechaEntrega,
      staff:        staff,
      notas:        String(row[9] || '')
    });
  }
  return rows.reverse();
}
