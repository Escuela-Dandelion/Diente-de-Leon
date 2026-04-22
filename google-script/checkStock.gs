// ============================================================
// DIENTE DE LEÓN — Alerta de Stock Bajo + Orden de Pedido
// ============================================================

const CONFIG_STOCK = {
  STORE_ID:       '7396246',
  API_TOKEN:      'e3d744d94ddbf13317bef0082c53e2c46fb50631',
  GREEN_INSTANCE: '7107556439',
  GREEN_TOKEN:    '7f7e2a357d2a4f0faa7edc4729f4c89f491b1420ef19462783',
  STOCK_UMBRAL:   10,
  FORM_URL:       'https://escuela-dandelion.github.io/Comision-Recursos/orden-de-pedido.html',
  RETIRO_URL:     'https://escuela-dandelion.github.io/Comision-Recursos/retiro-proveedor.html',
  SHEET_ID:       '1NmjnYWllrXrFpJI8GYJOjkGz90lPtvvDi0IQEYZl7-I',
  TEST_MODE:      false,
  TEST_TELEFONO:  '5493512112050'         // ← usado solo si TEST_MODE: true
};

const FGP_POR_MARCA = {
  'LA YAYA':                  { nombre: 'Yuliana Longhi',   telefono: '5493513645612', email: 'longhi.yuliana@gmail.com', tel_proveedor: '' },
  'ODDIS':                    { nombre: 'Luli Del Castillo', telefono: '5493513003789', email: null,                       tel_proveedor: '' },
  'CABALLO NEGRO':            { nombre: 'Maria Martini',     telefono: '5493516865078', email: 'martinimaria39@gmail.com', tel_proveedor: '' },
  'YEMARI':                   { nombre: 'Maria Martini',     telefono: '5493516865078', email: 'martinimaria39@gmail.com', tel_proveedor: '' },
  'GROEN':                    { nombre: 'Maria Martini',     telefono: '5493516865078', email: 'martinimaria39@gmail.com', tel_proveedor: '' },
  // PARAISA excluida — productos con stock infinito (null), no requieren alerta
  'EL MAITEN':                { nombre: 'Maria Martini',     telefono: '5493516865078', email: 'martinimaria39@gmail.com', tel_proveedor: '' },
  'GUARDIANES DE LA COLMENA': { nombre: 'Maria Martini',     telefono: '5493516865078', email: 'martinimaria39@gmail.com', tel_proveedor: '' }
};

// ── FUNCIÓN PRINCIPAL ──────────────────────────────────────
function checkStockBajo() {
  const productos = obtenerProductos();
  Logger.log(`Productos encontrados: ${productos.length}`);

  const alertasEnviadas = obtenerAlertasEnviadas();
  Logger.log(`Alertas enviadas previamente: ${JSON.stringify(alertasEnviadas)}`);
  const nuevasAlertas = {};

  // Agrupar alertas por FGP + marca (= un mensaje por proveedor por FGP)
  const alertasPorGrupo = {};

  productos.forEach(producto => {
    const marca = (producto.brand || '').toUpperCase();
    const fgp = FGP_POR_MARCA[marca];
    Logger.log(`Producto: ${JSON.stringify(producto.name)} | Marca: "${marca}" | FGP encontrado: ${!!fgp}`);

    if (!fgp) return;

    const nombreProducto = producto.name && producto.name.es
      ? producto.name.es
      : String(producto.name || 'Producto sin nombre');

    producto.variants.forEach(variante => {
      if (variante.stock === null) return;
      const stock = parseInt(variante.stock);
      const clave = `${producto.id}_${variante.id}`;
      Logger.log(`  Variante: ${clave} | Stock: ${stock} | Ya enviada: ${!!alertasEnviadas[clave]}`);

      if (stock <= CONFIG_STOCK.STOCK_UMBRAL && !alertasEnviadas[clave]) {
        const nombreVariante = variante.values && variante.values.length > 0
          ? variante.values.map(v => v.es || v).join(' / ')
          : null;
        const descripcion = nombreVariante
          ? `${nombreProducto} — ${nombreVariante}`
          : nombreProducto;

        const grupoKey = fgp.telefono + '|' + marca;
        if (!alertasPorGrupo[grupoKey]) {
          alertasPorGrupo[grupoKey] = { fgp: fgp, marca: marca, items: [] };
        }
        alertasPorGrupo[grupoKey].items.push({
          clave:       clave,
          descripcion: descripcion,
          stock:       stock,
          precio:      variante.price || null
        });
        nuevasAlertas[clave] = new Date().toISOString();
      }
    });
  });

  // Enviar UN mensaje por FGP + proveedor
  Object.values(alertasPorGrupo).forEach(function(grupo) {
    const fgp   = grupo.fgp;
    const marca = grupo.marca;
    const items = grupo.items;

    const listaProductos = items
      .map(function(item) {
        return '• *' + item.descripcion + '* — ' + item.stock + ' unidades';
      })
      .join('\n');

    // Generar N° de orden y ambas URLs
    const urls = generarUrls(items, marca, fgp);

    const prefijo = CONFIG_STOCK.TEST_MODE ? '🧪 *[PRUEBA]* \n\n' : '';
    const intro   = items.length === 1
      ? 'El siguiente producto de *' + marca + '* tiene stock bajo:'
      : 'Los siguientes productos de *' + marca + '* tienen stock bajo:';

    const mensaje =
      prefijo + '🌼 *Diente de León — Stock Bajo*\n\n' +
      'Hola ' + fgp.nombre + '! ' + intro + '\n\n' +
      listaProductos + '\n\n' +
      '━━━━━━━━━━━━━━━━\n' +
      '📋 *1. Hacer el pedido al proveedor:*\n' + urls.orden + '\n\n' +
      '📅 *2. Agendar el retiro (cuando confirme):*\n' + urls.retiro;

    enviarWhatsApp(fgp.telefono, mensaje);
    Logger.log('Mensaje enviado a ' + fgp.nombre + ' | Proveedor: ' + marca + ' | ' + items.length + ' producto(s)');
  });

  if (Object.keys(nuevasAlertas).length > 0) {
    const actualizadas = Object.assign(obtenerAlertasEnviadas(), nuevasAlertas);
    PropertiesService.getScriptProperties()
      .setProperty('ALERTAS_ENVIADAS', JSON.stringify(actualizadas));
  }
}

// ── LEER TEL DE PROVEEDOR DESDE EL SHEET ───────────────────
function leerTelProveedor(marca) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG_STOCK.SHEET_ID);
    const sheet = ss.getSheetByName('Proveedores');
    if (!sheet) return '';
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toUpperCase() === marca.toUpperCase()) return String(data[i][1] || '');
    }
  } catch(e) { Logger.log('Error leyendo tel proveedor: ' + e.message); }
  return '';
}

// ── GENERAR AMBAS URLs (orden + retiro) ────────────────────
function generarUrls(items, marca, fgp) {
  // Registrar orden una sola vez → obtener N° correlativo
  const nOrden = registrarNuevaOrden(marca, fgp, items);

  const productos  = items.map(function(i) { return i.descripcion; }).join('|');
  const cantidades = items.map(function(i) { return i.stock; }).join('|');
  const precios    = items.map(function(i) { return i.precio || ''; }).join('|');

  // Leer tel del proveedor (si fue guardado en algún pedido anterior)
  const telProveedor = fgp.tel_proveedor || leerTelProveedor(marca);

  // URL orden-de-pedido.html
  const paramsOrden = {
    producto:  productos,
    cantidad:  cantidades,
    precio:    precios,
    proveedor: marca,
    fgp:       fgp.nombre,
    norden:    nOrden
  };
  if (telProveedor) paramsOrden.tel = telProveedor;

  // URL retiro-proveedor.html
  const paramsRetiro = {
    proveedor: marca,
    producto:  productos.replace(/\|/g, '\n'),
    fgp:       fgp.nombre,
    norden:    nOrden
  };

  return {
    orden:  buildUrl(CONFIG_STOCK.FORM_URL,   paramsOrden),
    retiro: buildUrl(CONFIG_STOCK.RETIRO_URL, paramsRetiro)
  };
}

function buildUrl(base, params) {
  const query = Object.keys(params)
    .map(function(k) { return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]); })
    .join('&');
  return base + '?' + query;
}

// ── GOOGLE SHEETS — REGISTRO DE ÓRDENES ───────────────────
function registrarNuevaOrden(marca, fgp, items) {
  const ss     = SpreadsheetApp.openById(CONFIG_STOCK.SHEET_ID);
  const config = ss.getSheetByName('Config');
  const pedidos = ss.getSheetByName('Pedidos');

  // Leer y actualizar correlativo
  const ultimo = parseInt(config.getRange('B2').getValue()) || 0;
  const nuevo  = ultimo + 1;
  config.getRange('B2').setValue(nuevo);

  // Formatear número: ORD-2026-001
  const anio   = new Date().getFullYear();
  const nOrden = 'ORD-' + anio + '-' + String(nuevo).padStart(3, '0');

  // Registrar fila en Pedidos
  const descripcionProductos = items.map(function(i) { return i.descripcion; }).join('\n');
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  // Columnas: N°Orden · Fecha · Proveedor · FGP · Productos · Estado · Fecha Estado · Quien · Comprobante · Notas · Total
  pedidos.appendRow([
    nOrden,
    fecha,
    marca,
    fgp.nombre,
    descripcionProductos,
    'Solicitado',
    fecha,
    '',  // Quien (asignado al agendar retiro)
    '',  // Comprobante (link al pagar)
    '',  // Notas
    ''   // Total (se actualiza desde orden-de-pedido.html)
  ]);

  Logger.log('Orden registrada: ' + nOrden + ' | ' + marca + ' | ' + fgp.nombre);
  return nOrden;
}

// ── INICIALIZAR SHEET (correr una sola vez) ────────────────
function inicializarSheet() {
  const ss = SpreadsheetApp.openById(CONFIG_STOCK.SHEET_ID);

  // Tab Config
  let config = ss.getSheetByName('Config');
  if (!config) config = ss.insertSheet('Config');
  config.clearContents();
  config.getRange('A1:B1').setValues([['Parámetro', 'Valor']]);
  config.getRange('A2:B2').setValues([['ULTIMO_CORRELATIVO', 0]]);
  config.getRange('A1:B1').setFontWeight('bold').setBackground('#e8f5e9');

  // Tab Pedidos
  let pedidos = ss.getSheetByName('Pedidos');
  if (!pedidos) pedidos = ss.insertSheet('Pedidos');
  pedidos.clearContents();
  const headers = ['N° Orden', 'Fecha', 'Marca/Proveedor', 'FGP', 'Producto(s)', 'Estado', 'Fecha Estado', 'Notas'];
  pedidos.getRange(1, 1, 1, headers.length).setValues([headers]);
  pedidos.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f5e9').setFontColor('#3a7d44');
  pedidos.setColumnWidth(1, 130);
  pedidos.setColumnWidth(5, 280);
  pedidos.setColumnWidth(8, 200);

  // Validación de estado con dropdown
  const estadosValidos = 'Solicitado,Confirmado,En proceso,Lista para Recibir,Recibido,Pagado';
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(estadosValidos.split(','))
    .setAllowInvalid(false)
    .build();
  pedidos.getRange('F2:F1000').setDataValidation(regla);

  Logger.log('Sheet inicializada correctamente.');
}

// ── TIENDANUBE API ─────────────────────────────────────────
function obtenerProductos() {
  const response = UrlFetchApp.fetch(
    'https://api.tiendanube.com/v1/' + CONFIG_STOCK.STORE_ID + '/products?per_page=200', {
    method: 'GET',
    headers: {
      'Authentication': 'bearer ' + CONFIG_STOCK.API_TOKEN,
      'User-Agent': 'DienteDeLeon (dientedeleon-admin@googlegroups.com)'
    }
  });
  return JSON.parse(response.getContentText());
}

// ── GREEN-API ──────────────────────────────────────────────
function enviarWhatsApp(telefono, mensaje) {
  const destino = CONFIG_STOCK.TEST_MODE ? CONFIG_STOCK.TEST_TELEFONO : telefono;
  UrlFetchApp.fetch(
    'https://api.green-api.com/waInstance' + CONFIG_STOCK.GREEN_INSTANCE + '/sendMessage/' + CONFIG_STOCK.GREEN_TOKEN, {
    method: 'POST',
    contentType: 'application/json',
    payload: JSON.stringify({ chatId: destino + '@c.us', message: mensaje })
  });
}

// ── ALERTAS ────────────────────────────────────────────────
function obtenerAlertasEnviadas() {
  const raw = PropertiesService.getScriptProperties().getProperty('ALERTAS_ENVIADAS');
  return raw ? JSON.parse(raw) : {};
}

function resetearAlertas() {
  PropertiesService.getScriptProperties().deleteProperty('ALERTAS_ENVIADAS');
  Logger.log('Alertas reseteadas.');
}

// ── TRIGGER ────────────────────────────────────────────────
function crearTrigger() {
  ScriptApp.newTrigger('checkStockBajo')
    .timeBased().everyHours(6).create();
  Logger.log('Trigger creado: cada 6 horas.');
}
