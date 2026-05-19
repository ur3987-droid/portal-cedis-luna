// ═══════════════════════════════════════════════════════════════
// LUNA MEDIA CAFÉ — APPS SCRIPT COMPLETO v5
// Portal CEDIS: Pedidos + Inventario + Evaluaciones + Asignaciones
// ═══════════════════════════════════════════════════════════════

const CEDIS_EMAIL = 'cedislunamediacafe@gmail.com';
const ADMIN_EMAIL = 'lunamediacafedo@gmail.com';
const CEO_EMAIL   = 'ceo@lunamediacafe.com';
const HORAS_VENCIMIENTO = 72;

const SUCURSALES_INFO = {
  'Avila Camacho':  { correo: 'lunamediacentro@gmail.com',       gerente: 'Luis Enrique Velazquez Hernandez' },
  'Ávila Camacho':  { correo: 'lunamediacentro@gmail.com',       gerente: 'Luis Enrique Velazquez Hernandez' },
  'Paseo Jardines': { correo: 'lunamediapaseojardines@gmail.com', gerente: 'Ingrid Arely Martinez Diaz'      },
  'Postreria':      { correo: 'postrerialunamediacafe@gmail.com', gerente: 'Eduardo Hernandez Cancela'       },
  'Postrería':      { correo: 'postrerialunamediacafe@gmail.com', gerente: 'Eduardo Hernandez Cancela'       }
};

// ═══════════════════════════════════════════════════════════════
// doPost
// ═══════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const p    = e.parameter;
    const tipo = p.tipo || '';
    Logger.log('doPost recibido - tipo: [' + tipo + '] keys: ' + Object.keys(p).join(','));
    if (tipo === 'pedido')     return handlePedido(p);
    if (tipo === 'inventario') return handleInventario(p);
    if (tipo === 'evaluacion')    return handleEvaluacion(p);
    if (tipo === 'transferencia') return handleTransferencia(p);
    if (tipo === 'eliminar_sdp')  return eliminarSDP(e.parameter);
    if (tipo === 'eliminar_inv')  return eliminarInv(e.parameter);
    if (tipo === 'cancelar_sdp')  return cancelarSDP(e.parameter);
    if (tipo === 'espejo_pub')    return espejoPub(p);
    if (tipo === 'cancelar_sdp')   return cancelarSDP(p);
    if (tipo === 'eliminar_sdp')   return eliminarSDP(p);
    if (tipo === 'eliminar_inv')   return eliminarInv(p);
    return ok('tipo no reconocido');
  } catch(err) {
    return error(err.toString());
  }
}

// ═══════════════════════════════════════════════════════════════
// doGet
// ═══════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const tipo     = e.parameter.tipo     || '';
    Logger.log('doGet recibido - tipo: [' + tipo + ']');
    const sucursal = e.parameter.sucursal || '';
    const folio    = e.parameter.folio    || '';
    if (tipo === 'pedidos_pendientes') return getPedidosPendientes(sucursal);
    if (tipo === 'detalle_pedido')     return getDetallePedido(folio);
    if (tipo === 'catalogo')           return getCatalogo();
    if (tipo === 'ping')               return ok('portal activo');
    if (tipo === 'pedido')             return handlePedido(e.parameter);
    if (tipo === 'inventario')         return handleInventario(e.parameter);
    if (tipo === 'evaluacion')         return handleEvaluacion(e.parameter);
    if (tipo === 'historial')          return getHistorial();
    if (tipo === 'historial_limpiar')  return limpiarHistorial();
    if (tipo === 'inventario_existencia') return getInventarioExistencia(sucursal);
    if (tipo === 'borrador_save')   return guardarBorrador(e.parameter);
    if (tipo === 'borrador_get')    return getBorrador(sucursal);
    if (tipo === 'borrador_delete') return eliminarBorrador(sucursal);
    if (tipo === 'espejo_update')   return actualizarEspejo(e.parameter);
    if (tipo === 'espejo_list')       return listarEspejo();

    if (tipo === 'conciliacion_list') return getConciliacion(e.parameter);
    if (tipo === 'historial_sdp')     return getHistorialSDP(e.parameter);
    if (tipo === 'get_pdf_url')       return getPdfUrl(e.parameter);
    if (tipo === 'eliminar_sdp')      return eliminarSDP(e.parameter);
    if (tipo === 'check_entrega')     return registrarCheckEntrega(e.parameter);

    if (tipo === 'get_check')         return getCheckEntrega(e.parameter);
    if (tipo === 'sesion_ping')       return sesionPing(e.parameter);
    if (tipo === 'sesion_check')      return sesionCheck(e.parameter);
    if (tipo === 'sesion_cerrar')     return sesionCerrar(e.parameter);
    if (tipo === 'espejo_get')        return espejoGet(e.parameter);
    if (tipo === 'transferencia')     return handleTransferencia(e.parameter);
    if (tipo === 'eliminar_sdp')      return eliminarSDP(e.parameter);
    if (tipo === 'eliminar_inv')      return eliminarInv(e.parameter);
    if (tipo === 'cancelar_sdp')      return cancelarSDP(e.parameter);
    return ok('tipo no reconocido');
  } catch(err) {
    return error(err.toString());
  }
}

// ═══════════════════════════════════════════════════════════════
// PEDIDOS Y ASIGNACIONES
// ═══════════════════════════════════════════════════════════════
function handlePedido(p) {
  Logger.log('PEDIDO RECIBIDO - folio: ' + p.folio + ' sucursal: ' + p.sucursal);
  const ss  = SpreadsheetApp.openById('1hqpPHPr9msdhqjN3WBMYGrIH9XxRVmdoayhcscipNJg');
  const tab = normalizarTab(p.sucursal);
  let sheet = ss.getSheetByName(tab);
  if (!sheet) sheet = ss.insertSheet(tab);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Folio','Fecha SDP','Sucursal','Correo','Gerente/Encargado',
      'Entrega Requerida','# Productos','Total Estimado',
      'Observaciones','Productos Detalle',
      'Estado','Fecha Vencimiento','Fecha Evaluacion',
      'Total Real Recibido','Productos Recibidos','Calificacion','Comentarios Evaluacion'
    ]);
    sheet.getRange(1,1,1,17).setFontWeight('bold').setBackground('#1D4F73').setFontColor('#FFFFFF');
  }

  const ahora       = new Date();
  const vencimiento = new Date(ahora.getTime() + HORAS_VENCIMIENTO * 60 * 60 * 1000);

  sheet.appendRow([
    p.folio, p.timestamp, p.sucursal, p.correoSuc, p.gerente,
    p.fechaEntrega, p.numProductos, p.totalEstimado,
    p.obsGenerales || '', p.productos,
    'PENDIENTE', vencimiento, '', '', '', '', ''
  ]);

  try {
    const esAsi   = (p.(folio.indexOf('-ASI') >= 0 || folio.indexOf('ASI-') === 0) || p.folio.indexOf('ASI-') === 0);
    const esExt   = p.folio.indexOf('-EXT') >= 0;
    const tipoSDP = esAsi ? 'Asignacion' : (esExt ? 'Extraordinaria' : 'Ordinaria');
    const asunto  = esAsi
      ? 'Nueva Asignacion CEDIS - ' + p.sucursal + ' | ' + p.folio
      : 'Nueva SDP (' + tipoSDP + ') - ' + p.sucursal + ' | ' + p.folio;
    const htmlCaratula = generarHTMLCaratula(p);
    const pdfBlob      = generarPDF(htmlCaratula, p.folio, 'Pedidos', p.sucursal);
    const cuerpoHtml   = generarCorreoPedido(p);
    // Sucursal/area destino siempre recibe
    if (p.correoSuc) enviarCorreoConPDF(p.correoSuc, asunto, cuerpoHtml, pdfBlob);
    // CEDIS siempre recibe (incluso en asignaciones ASI)
    enviarCorreoConPDF(CEDIS_EMAIL, asunto, cuerpoHtml, pdfBlob);
    enviarCorreoConPDF(ADMIN_EMAIL, asunto, cuerpoHtml, pdfBlob);
    enviarCorreoConPDF(CEO_EMAIL,   asunto, cuerpoHtml, pdfBlob);
    Logger.log('Correos pedido enviados OK - tipo: ' + tipoSDP);

    // Notificacion especial a Postreria si hay productos POSTRERIA PRODUCTO TERMINADO
    try {
      const lineas = (p.productos || '').split('\n').filter(function(l){ return l.trim(); });
      const lineasPost = lineas.filter(function(l){
        // productos de postreria tienen prefijo RPT
        const partes = l.split('|');
        return (partes[0] || '').trim().indexOf('RPT') === 0;
      });
      if (lineasPost.length > 0) {
        let filasPost = '';
        lineasPost.forEach(function(linea, idx) {
          const partes = linea.split('|');
          filasPost += '<tr style="background:' + (idx%2===0?'white':'#f5f8fc') + '">' +
            '<td style="padding:7px 12px;font-size:12px;font-family:monospace;">' + (partes[0]||'') + '</td>' +
            '<td style="padding:7px 12px;font-size:12px;">' + (partes[1]||'') + '</td>' +
            '<td style="padding:7px 12px;font-size:12px;">' + (partes[2]||'') + '</td>' +
            '<td style="padding:7px 12px;font-size:12px;text-align:right;">' + (partes[3]||'') + '</td>' +
            '<td style="padding:7px 12px;font-size:12px;text-align:right;font-weight:700;">' + (partes[4]||'').replace('x','') + '</td>' +
            '<td style="padding:7px 12px;font-size:12px;text-align:right;font-weight:700;color:#1D4F73;">' + (partes[5]||'') + '</td>' +
            '</tr>';
        });
        const asuntoPost = 'Solicitud de Produccion — ' + p.sucursal + ' | ' + p.folio;
        const cuerpoPost = '<div style="font-family:sans-serif;max-width:680px;margin:0 auto;background:#F5F8FC;">' +
          '<div style="background:linear-gradient(135deg,#1D4F73,#3C8EC8);padding:24px 28px;border-radius:12px 12px 0 0;">' +
          '<div style="color:white;font-size:22px;font-weight:700;">Solicitud de Produccion</div>' +
          '<div style="color:rgba(255,255,255,.7);font-size:12px;margin-top:4px;">Copia automatica — Solo productos Postreria Producto Terminado</div>' +
          '<div style="color:rgba(255,255,255,.8);font-size:13px;margin-top:4px;">' + p.timestamp + '</div></div>' +
          '<div style="background:white;padding:20px 28px;">' +
          '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:20px;">' +
          '<tr><td style="padding:8px;color:#6B7E93;font-weight:700;width:40%">FOLIO REFERENCIA</td><td style="padding:8px;font-size:16px;font-weight:700;color:#3C8EC8;">' + p.folio + '</td></tr>' +
          '<tr style="background:#f5f8fc"><td style="padding:8px;color:#6B7E93;font-weight:700">SUCURSAL SOLICITANTE</td><td style="padding:8px;font-weight:600">' + p.sucursal + '</td></tr>' +
          '<tr><td style="padding:8px;color:#6B7E93;font-weight:700">ENTREGA REQUERIDA</td><td style="padding:8px;font-weight:600">' + p.fechaEntrega + '</td></tr>' +
          '</table>' +
          '<div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#6B7E93;margin-bottom:8px;">Productos a producir</div>' +
          '<table style="width:100%;border-collapse:collapse;">' +
          '<thead><tr style="background:#1D4F73;">' +
          '<th style="padding:8px 12px;color:white;font-size:11px;text-align:left;">Codigo</th>' +
          '<th style="padding:8px 12px;color:white;font-size:11px;text-align:left;">Producto</th>' +
          '<th style="padding:8px 12px;color:white;font-size:11px;text-align:left;">Presentacion</th>' +
          '<th style="padding:8px 12px;color:white;font-size:11px;text-align:right;">Precio</th>' +
          '<th style="padding:8px 12px;color:white;font-size:11px;text-align:right;">Cantidad</th>' +
          '<th style="padding:8px 12px;color:white;font-size:11px;text-align:right;">Subtotal</th>' +
          '</tr></thead><tbody>' + filasPost + '</tbody></table>' +
          '<div style="background:#FEF3E8;border:1px solid #E07B2A;border-radius:8px;padding:12px 16px;margin-top:16px;font-size:12px;color:#8B4A10;">' +
          'Este es un aviso automatico. La entrega debe estar lista para la fecha indicada.' +
          '</div></div></div>';
        MailApp.sendEmail({ to: 'postrerialunamediacafe@gmail.com', subject: asuntoPost, htmlBody: cuerpoPost, name: 'LunaMedia Ops' });
        Logger.log('Notificacion Postreria enviada: ' + lineasPost.length + ' productos RPT');
      }
    } catch(errPost) {
      Logger.log('Error notificacion Postreria: ' + errPost);
    }
  } catch(err) {
    Logger.log('Error enviando correo pedido: ' + err);
  }

  return ok('pedido guardado');
}

// ═══════════════════════════════════════════════════════════════
// INVENTARIO
// ═══════════════════════════════════════════════════════════════
function handleInventario(p) {
  if (!p || !p.sucursal) return error('sin datos');
  Logger.log('Inventario recibido - Sucursal: ' + p.sucursal);

  const folio = p.folio || 'INV-' + Date.now().toString().slice(-6);
  const fecha = p.timestamp || new Date().toLocaleString('es-MX');
  const etiquetaGerente = (p.sucursal === 'Postreria' || p.sucursal === 'Postrería') ? 'Encargado' : 'Gerente';

  const sucInfo = {
    'Avila Camacho':  'lunamediacentro@gmail.com',
    'Ávila Camacho':  'lunamediacentro@gmail.com',
    'Paseo Jardines': 'lunamediapaseojardines@gmail.com',
    'Postreria':      'postrerialunamediacafe@gmail.com',
    'Postrería':      'postrerialunamediacafe@gmail.com'
  };
  const correoSuc = sucInfo[p.sucursal] || '';
  const asunto = 'Nuevo INV — ' + p.sucursal + ' | ' + fecha;

  const resumenItems = (p.resumen || '').split(';').filter(function(i){ return i.trim(); });
  let filas = '';
  resumenItems.forEach(function(item, idx){
    const partes = item.split('|');
    filas += '<tr style="background:' + (idx%2===0?'white':'#f5f8fc') + '">' +
      '<td style="padding:7px 12px;font-size:12px;font-family:monospace;">' + (partes[0]||'') + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;">' + (partes[1]||'') + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;text-align:right;">' + (partes[2]||'') + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;text-align:right;">' + (partes[3]||'') + '</td>' +
      '</tr>';
  });

  const cuerpo = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>' +
    '<div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;background:#F5F8FC;">' +
    '<div style="background:linear-gradient(135deg,#1D4F73,#3C8EC8);padding:24px 28px;border-radius:12px 12px 0 0;">' +
    '<div style="color:white;font-size:11px;letter-spacing:2px;text-transform:uppercase;opacity:.7;margin-bottom:4px;">Luna Media Cafe — CEDIS</div>' +
    '<div style="color:white;font-size:22px;font-weight:700;">Inventario Registrado</div>' +
    '<div style="color:rgba(255,255,255,.8);font-size:13px;margin-top:4px;">' + fecha + '</div></div>' +
    '<div style="background:white;padding:20px 28px;">' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:20px;">' +
    '<tr><td style="padding:8px 0;color:#6B7E93;font-weight:700;width:40%">FOLIO</td><td style="padding:8px 0;font-size:18px;font-weight:700;color:#3C8EC8;">' + folio + '</td></tr>' +
    '<tr style="background:#f5f8fc"><td style="padding:8px;color:#6B7E93;font-weight:700">SUCURSAL</td><td style="padding:8px;font-weight:600">' + p.sucursal + '</td></tr>' +
    '<tr><td style="padding:8px;color:#6B7E93;font-weight:700">' + etiquetaGerente.toUpperCase() + '</td><td style="padding:8px">' + p.gerente + '</td></tr>' +
    '<tr style="background:#f5f8fc"><td style="padding:8px;color:#6B7E93;font-weight:700">FECHA</td><td style="padding:8px">' + fecha + '</td></tr>' +
    '<tr><td style="padding:8px;color:#6B7E93;font-weight:700">TOTAL ITEMS</td><td style="padding:8px;font-weight:700">' + p.totalItems + '</td></tr>' +
    '<tr style="background:#f5f8fc"><td style="padding:8px;color:#6B7E93;font-weight:700">COSTO TOTAL</td><td style="padding:8px;font-size:16px;font-weight:700;color:#1D4F73;">$' + p.totalCosto + '</td></tr>' +
    '</table>' +
    '<div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#6B7E93;margin-bottom:8px;">Productos con existencia</div>' +
    '<table style="width:100%;border-collapse:collapse;border-radius:8px;overflow:hidden;">' +
    '<thead><tr style="background:#1D4F73;">' +
    '<th style="padding:8px 12px;color:white;font-size:11px;text-align:left;">Codigo</th>' +
    '<th style="padding:8px 12px;color:white;font-size:11px;text-align:left;">Producto</th>' +
    '<th style="padding:8px 12px;color:white;font-size:11px;text-align:right;">Cantidad</th>' +
    '<th style="padding:8px 12px;color:white;font-size:11px;text-align:right;">Precio</th>' +
    '</tr></thead><tbody>' + filas + '</tbody>' +
    '<tfoot><tr style="background:#EBF4FB;">' +
    '<td colspan="2" style="padding:10px 12px;font-weight:700;text-align:right;color:#1D4F73;">TOTAL</td>' +
    '<td style="padding:10px 12px;font-weight:700;color:#1D4F73;text-align:right;">' + p.totalItems + '</td>' +
    '<td style="padding:10px 12px;font-weight:700;color:#1D4F73;font-size:15px;text-align:right;">$' + p.totalCosto + '</td>' +
    '</tr></tfoot></table>' +
    '</div>' +
    '<div style="padding:14px 28px;text-align:center;font-size:10px;color:#bbb;border-top:1px solid #EEF2F7;">Luna Media Cafe · Sistema CEDIS · Generado automaticamente</div>' +
    '</div></body></html>';

  try {
    // Generar PDF inventario, guardar en Drive y adjuntar en correos
    let pdfInvBlob = null;
    try {
      pdfInvBlob = generarPDF(cuerpo, folio, 'Inventarios', p.sucursal);
      Logger.log('PDF inventario guardado: ' + folio);
    } catch(errPdf) {
      Logger.log('Error generando PDF inventario: ' + errPdf);
    }
    // Enviar correos con PDF adjunto si se generó
    const correosInv = [correoSuc, CEDIS_EMAIL, ADMIN_EMAIL, CEO_EMAIL].filter(Boolean);
    correosInv.forEach(function(to) {
      try {
        const mailOpts = { to: to, subject: asunto, htmlBody: cuerpo, name: 'LunaMedia Ops' };
        if (pdfInvBlob) mailOpts.attachments = [pdfInvBlob];
        MailApp.sendEmail(mailOpts);
      } catch(errMail) { Logger.log('Error correo inv a ' + to + ': ' + errMail); }
    });
    Logger.log('Correos inventario enviados OK');
  } catch(e) {
    Logger.log('Error enviando correo inventario: ' + e);
  }

  // Get PDF URL from Drive
  var pdfUrlInv = '';
  try {
    var sucNormInv = (p.sucursal||'').replace('Ávila Camacho','Avila Camacho').replace('Postrería','Postreria');
    var invFolder = getPDFSubFolder('Inventarios', sucNormInv);
    var invFiles = invFolder.getFilesByName(folio + '.pdf');
    if(invFiles.hasNext()){ var invFile = invFiles.next(); pdfUrlInv = 'https://drive.google.com/file/d/' + invFile.getId() + '/view'; }
  } catch(ePdfInv){}

  saveHistorialEntry({
    folio: folio,
    sucursal: p.sucursal,
    gerente: p.gerente,
    dia: p.dia,
    fecha: fecha,
    conteo: p.conteo || '',
    pdfUrl: pdfUrlInv,
    timestamp: new Date().getTime()
  });

  // Alerta si hay productos en cero
  try {
    const items = (p.resumen || '').split(';').filter(function(i){ return i.trim(); });
    // resumen only has products WITH quantity - zeros are in conteo
    const conteoItems = (p.conteo || '').split(';').filter(function(i){ return i.trim(); });
    const ceros = conteoItems.filter(function(item){
      const parts = item.split('|');
      return parseFloat(parts[1] || '0') === 0;
    }).map(function(item){ return item.split('|')[0]; });
    if (ceros.length > 0) {
      alertaInventarioCeros(p.sucursal, ceros, fecha, folio);
    }
  } catch(errCeros) { Logger.log('Error alertaInventarioCeros: ' + errCeros); }

  return ok('inventario guardado');
}

// ═══════════════════════════════════════════════════════════════
// EVALUACION
// ═══════════════════════════════════════════════════════════════
function handleEvaluacion(p) {
  Logger.log('EVALUACION RECIBIDA - folio: ' + p.folio + ' sucursal: [' + p.sucursal + '] keys: ' + Object.keys(p).join(','));
  // Fallback: if sucursal missing, try to find it from sheets
  if (!p.sucursal) {
    var tabs = ['ÁVILA CAMACHO','PASEO JARDINES','POSTRERÍA'];
    var ssBak = SpreadsheetApp.openById('1hqpPHPr9msdhqjN3WBMYGrIH9XxRVmdoayhcscipNJg');
    for (var t = 0; t < tabs.length; t++) {
      var sh = ssBak.getSheetByName(tabs[t]);
      if (!sh) continue;
      var d = sh.getDataRange().getValues();
      for (var r = 1; r < d.length; r++) {
        if (String(d[r][0]) === String(p.folio)) {
          p.sucursal = String(d[r][2]);
          Logger.log('Sucursal recuperada del sheet: ' + p.sucursal);
          break;
        }
      }
      if (p.sucursal) break;
    }
  }
  const ss    = SpreadsheetApp.openById('1hqpPHPr9msdhqjN3WBMYGrIH9XxRVmdoayhcscipNJg');
  const tab   = normalizarTab(p.sucursal);
  const sheet = ss.getSheetByName(tab);
  if (!sheet) return error('No se encontro la hoja: ' + p.sucursal);

  const data = sheet.getDataRange().getValues();
  let encontrado = false;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.folio)) {
      const estado = data[i][10];
      if (estado === 'VENCIDA')  return error('Folio vencido: ' + p.folio);
      if (estado === 'EVALUADA') return error('Folio ya evaluado: ' + p.folio);
      const row = i + 1;
      sheet.getRange(row, 11).setValue('EVALUADA');
      sheet.getRange(row, 13).setValue(p.timestamp);
      sheet.getRange(row, 14).setValue(p.totalReal);
      sheet.getRange(row, 15).setValue(p.productosRecibidos);
      sheet.getRange(row, 16).setValue(p.calificacion);
      sheet.getRange(row, 17).setValue(p.comentarios);
      encontrado = true;
      break;
    }
  }

  if (!encontrado) return error('Folio no encontrado: ' + p.folio);

  const info   = SUCURSALES_INFO[p.sucursal] || {};
  const asunto = 'Evaluacion registrada - ' + p.sucursal + ' | ' + p.folio;
  // Defensive: ensure p has all required fields
  if (!p.sucursal) { Logger.log('ERROR: p.sucursal is undefined. p = ' + JSON.stringify(p)); return error('Datos incompletos: falta sucursal'); }
  const cuerpo = generarCorreoEvaluacion(p);

  try {
    var correosEv = [CEDIS_EMAIL, info.correo, ADMIN_EMAIL, CEO_EMAIL].filter(function(c){ return c && c.trim() !== ''; });
    Logger.log('Enviando correos evaluacion a: ' + correosEv.join(', '));
    correosEv.forEach(function(to){
      try {
        MailApp.sendEmail({ to: to, subject: asunto, htmlBody: cuerpo, name: 'LunaMedia Ops' });
        Logger.log('OK correo evaluacion: ' + to);
      } catch(eM){ Logger.log('Error correo a ' + to + ': ' + eM); }
    });
  } catch(err) {
    Logger.log('Error enviando correo evaluacion: ' + err);
  }

  // Alerta solo si entrega fue INCOMPLETA o RECHAZADA segun estado evaluado
  try {
    const estadoEv = (p.estado || p.estadoEntrega || '').toLowerCase();
    const esIncompleta = estadoEv === 'incompleto' || estadoEv === 'incompleta';
    const esRechazada  = estadoEv === 'rechazado'  || estadoEv === 'rechazada';
    if (esIncompleta || esRechazada) {
      alertaEntrega(p);
    }
  } catch(errAl) { Logger.log('Error alertaEntrega: ' + errAl); }

  return ok('evaluacion guardada');
}

// ═══════════════════════════════════════════════════════════════
// PEDIDOS PENDIENTES
// ═══════════════════════════════════════════════════════════════
function getPedidosPendientes(sucursal) {
  const ss    = SpreadsheetApp.openById('1hqpPHPr9msdhqjN3WBMYGrIH9XxRVmdoayhcscipNJg');
  const tab   = normalizarTab(sucursal);
  const sheet = ss.getSheetByName(tab);
  if (!sheet || sheet.getLastRow() <= 1) return jsonR([]);

  const data   = sheet.getDataRange().getValues();
  const ahora  = new Date();
  const result = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const estado = String(data[i][10] || '');
    let vencimiento = data[i][11];
    if (!(vencimiento instanceof Date)) vencimiento = new Date(vencimiento);

    if ((estado === 'PENDIENTE' || estado === 'EN_RUTA') && ahora < vencimiento) {
      const msRestantes  = vencimiento - ahora;
      const hrsRestantes = Math.floor(msRestantes / (1000 * 60 * 60));
      result.push({
        folio:         String(data[i][0]),
        fecha:         String(data[i][1]),
        sucursal:      String(data[i][2]),
        gerente:       String(data[i][4]),
        fechaEntrega:  String(data[i][5]),
        numProductos:  String(data[i][6]),
        totalEstimado: String(data[i][7]),
        hrsRestantes:  hrsRestantes
      });
    }
  }
  return jsonR(result);
}

// ═══════════════════════════════════════════════════════════════
// DETALLE DE PEDIDO
// ═══════════════════════════════════════════════════════════════
function getDetallePedido(folio) {
  const ss     = SpreadsheetApp.openById('1hqpPHPr9msdhqjN3WBMYGrIH9XxRVmdoayhcscipNJg');
  const sheets = ss.getSheets();

  for (const sheet of sheets) {
    if (sheet.getName().indexOf('INV') === 0) continue;
    if (sheet.getLastRow() <= 1) continue;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(folio)) {
        const vencimiento  = new Date(data[i][11]);
        const ahora        = new Date();
        const msRestantes  = vencimiento - ahora;
        const hrsRestantes = Math.max(0, Math.floor(msRestantes / (1000 * 60 * 60)));
        const estado       = String(data[i][10] || 'PENDIENTE');

        if ((estado === 'PENDIENTE' || estado === 'EN_RUTA') && ahora > vencimiento) {
          sheet.getRange(i + 1, 11).setValue('VENCIDA');
          return jsonR({ error: 'Folio vencido' });
        }

        return jsonR({
          folio:         String(data[i][0]),
          fecha:         String(data[i][1]),
          sucursal:      String(data[i][2]),
          correoSuc:     String(data[i][3]),
          gerente:       String(data[i][4]),
          fechaEntrega:  String(data[i][5]),
          numProductos:  String(data[i][6]),
          totalEstimado: String(data[i][7]),
          obsGenerales:  String(data[i][8] || ''),
          productos:     String(data[i][9] || ''),
          estado:        estado,
          hrsRestantes:  hrsRestantes
        });
      }
    }
  }
  return jsonR({ error: 'Folio no encontrado' });
}

// ═══════════════════════════════════════════════════════════════
// CATALOGO
// ═══════════════════════════════════════════════════════════════
function getCatalogo() {
  const SHEET_ID = '1dfRsQ07VAPqO5v_gZi5mAVvpYHhIRz4YR79WRC0bY0A';
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheets()[0];
    const data  = sheet.getDataRange().getValues();
    const productos = [];
    let catActual = '';

    for (let i = 1; i < data.length; i++) {
      const col1 = String(data[i][1] || '').trim();
      const col2 = String(data[i][2] || '').trim();
      const col3 = String(data[i][3] || '').trim();
      const col4 = String(data[i][4] || '').trim();

      if (col1 === 'CODIGO' || col1 === 'CÓDIGO') { catActual = col2; continue; }
      if (!col1 || !col2) continue;
      if (catActual.indexOf('INSUMOS POR PAQUETE') >= 0) continue;

      const precio = parseFloat(col4.replace('$','').replace(',','').trim()) || 0;
      productos.push({ code: col1, name: col2, pres: col3, precio: precio, cat: catActual });
    }
    return jsonR(productos);
  } catch(err) {
    return jsonR({ error: err.toString() });
  }
}

// ═══════════════════════════════════════════════════════════════
// HISTORIAL (15 días)
// ═══════════════════════════════════════════════════════════════
const HISTORIAL_FILE   = 'historial_inventarios.json';
const CEDIS_TELEFONO   = '2288365983'; // Telefono autorizado CEDIS — cambiar aqui si cambia el equipo
const PDF_PARENT_FOLDER = '1rxYDJoA_Ycl1lYbbiX8VTvcIiSNuQklo';
const PDF_FOLDER_NAME   = 'PDFs Portal OpsLMC';
const HISTORIAL_FOLDER = '1lZBhWcEX-BhMnXUBw95nWRPv1XasWSTa';

function getHistorialFile() {
  const folder = DriveApp.getFolderById(HISTORIAL_FOLDER);
  const files  = folder.getFilesByName(HISTORIAL_FILE);
  if (files.hasNext()) return files.next();
  return folder.createFile(HISTORIAL_FILE, '[]', 'application/json');
}

function getHistorial() {
  try {
    const file = getHistorialFile();
    const data = JSON.parse(file.getBlob().getDataAsString());
    const hace100dias = new Date().getTime() - (100 * 24 * 60 * 60 * 1000);
    const filtrado = data.filter(function(d){ return (d.timestamp || 0) >= hace100dias; });
    file.setContent(JSON.stringify(filtrado));
    return jsonR(filtrado);
  } catch(e) {
    return jsonR([]);
  }
}

function saveHistorialEntry(entry) {
  try {
    const file = getHistorialFile();
    const data = JSON.parse(file.getBlob().getDataAsString());
    data.unshift(entry);
    file.setContent(JSON.stringify(data));
  } catch(e) {
    Logger.log('Error saving historial: ' + e);
  }
}


function getInventarioExistencia(sucursal) {
  try {
    const file = getHistorialFile();
    const data = JSON.parse(file.getBlob().getDataAsString());
    // Find most recent entry for this sucursal that has conteo
    const entry = data.find(function(d){
      var dSuc = (d.sucursal || '').trim();
      var sSuc = (sucursal || '').trim();
      return (dSuc === sSuc || dSuc.normalize('NFD').replace(/[\u0300-\u036f]/g,'') === sSuc.normalize('NFD').replace(/[\u0300-\u036f]/g,'')) && d.conteo;
    });
    if (!entry) return jsonR({ conteo: '' });
    return jsonR({ folio: entry.folio, fecha: entry.fecha, conteo: entry.conteo });
  } catch(e) {
    Logger.log('Error getInventarioExistencia: ' + e);
    return jsonR({ conteo: '' });
  }
}

function limpiarHistorial() {
  try {
    const file = getHistorialFile();
    file.setContent('[]');
    Logger.log('Historial limpiado OK');
    return ok('historial limpiado');
  } catch(e) {
    Logger.log('Error limpiando historial: ' + e);
    return error(e.toString());
  }
}

// ═══════════════════════════════════════════════════════════════
// TRIGGER — revisar vencimientos cada hora
// ═══════════════════════════════════════════════════════════════
function revisarVencimientos() {
  const ss     = SpreadsheetApp.openById('1hqpPHPr9msdhqjN3WBMYGrIH9XxRVmdoayhcscipNJg');
  const sheets = ss.getSheets();
  const ahora  = new Date();

  sheets.forEach(function(sheet) {
    if (sheet.getName().indexOf('INV') === 0) return;
    if (sheet.getLastRow() <= 1) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const estado      = String(data[i][10] || '');
      const vencimiento = new Date(data[i][11]);
      if ((estado === 'PENDIENTE' || estado === 'EN_RUTA') && ahora > vencimiento) {
        sheet.getRange(i + 1, 11).setValue('VENCIDA');
        notificarVencimiento(data[i]);
      }
      if (estado === 'PENDIENTE') {
        const hrsRestantes = (vencimiento - ahora) / (1000 * 60 * 60);
        if (hrsRestantes > 0 && hrsRestantes <= 24 && !data[i][12]) {
          enviarAlertaPreVencimiento(data[i]);
        }
      }
    }
  });
}

function notificarVencimiento(row) {
  const folio    = String(row[0]);
  const sucursal = String(row[2]);
  const correo   = String(row[3]);
  const total    = String(row[7]);
  const asunto   = 'ALERTA: SDP vencida sin evaluar - ' + folio + ' | ' + sucursal;
  const cuerpo   = '<div style="font-family:sans-serif;max-width:600px;">' +
    '<div style="background:#E07B2A;padding:20px;border-radius:8px 8px 0 0;"><h2 style="color:white;margin:0;">Evaluacion VENCIDA</h2></div>' +
    '<div style="background:#f5f8fc;padding:20px;border-radius:0 0 8px 8px;">' +
    '<p>El plazo de 72 horas para evaluar el folio <strong>' + folio + '</strong> ha vencido.</p>' +
    '<p>Total considerado: <strong>' + total + '</strong></p></div></div>';
  try {
    MailApp.sendEmail({ to: correo,      subject: asunto, htmlBody: cuerpo, name: 'LunaMedia Ops' });
    MailApp.sendEmail({ to: CEDIS_EMAIL, subject: asunto, htmlBody: cuerpo, name: 'LunaMedia Ops' });
    MailApp.sendEmail({ to: ADMIN_EMAIL, subject: asunto, htmlBody: cuerpo, name: 'LunaMedia Ops' });
    MailApp.sendEmail({ to: CEO_EMAIL,   subject: asunto, htmlBody: cuerpo, name: 'LunaMedia Ops' });
  } catch(err) { Logger.log('Error notificando vencimiento: ' + err); }
}

function enviarAlertaPreVencimiento(row) {
  const folio    = String(row[0]);
  const sucursal = String(row[2]);
  const correo   = String(row[3]);
  const asunto   = 'URGENTE: Solo 24hrs para evaluar SDP - ' + folio + ' | ' + sucursal;
  const cuerpo   = '<div style="font-family:sans-serif;max-width:600px;">' +
    '<div style="background:#E07B2A;padding:20px;border-radius:8px 8px 0 0;"><h2 style="color:white;margin:0;">Quedan menos de 24 horas</h2></div>' +
    '<div style="background:#f5f8fc;padding:20px;border-radius:0 0 8px 8px;">' +
    '<p>El folio <strong>' + folio + '</strong> de ' + sucursal + ' esta proximo a vencer. Evalua la entrega ahora.</p></div></div>';
  try {
    MailApp.sendEmail({ to: correo,      subject: asunto, htmlBody: cuerpo, name: 'LunaMedia Ops' });
    MailApp.sendEmail({ to: CEDIS_EMAIL, subject: asunto, htmlBody: cuerpo, name: 'LunaMedia Ops' });
  } catch(err) { Logger.log('Error alerta pre-vencimiento: ' + err); }
}

// ═══════════════════════════════════════════════════════════════
// PDF Y CORREOS
// ═══════════════════════════════════════════════════════════════

function getOrCreateFolder(parent, nombre) {
  const existing = parent.getFoldersByName(nombre);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(nombre);
}

function getPDFSubFolder(tipo, sucursal) {
  // tipo: 'Pedidos' o 'Inventarios'
  const root    = DriveApp.getFolderById(PDF_PARENT_FOLDER);
  const mainDir = getOrCreateFolder(root, PDF_FOLDER_NAME);
  const tipoDir = getOrCreateFolder(mainDir, tipo);
  // Normalize sucursal name for folder
  const sucNorm = sucursal
    .replace('ÁVILA CAMACHO','Avila Camacho')
    .replace('PASEO JARDINES','Paseo Jardines')
    .replace('POSTRERÍA','Postreria')
    .replace('Ávila Camacho','Avila Camacho')
    .replace('Postrería','Postreria') || 'General';
  return getOrCreateFolder(tipoDir, sucNorm);
}

function guardarPDFEnDrive(pdfBlob, nombreArchivo, tipo, sucursal) {
  try {
    const folder = getPDFSubFolder(tipo || 'General', sucursal || '');
    const file   = folder.createFile(pdfBlob.setName(nombreArchivo + '.pdf'));
    // Make accessible via link for portal viewing
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log('PDF guardado: ' + tipo + '/' + sucursal + '/' + nombreArchivo);
    return file;
  } catch(e) {
    Logger.log('Error guardando PDF en Drive: ' + e);
    return null;
  }
}

function generarPDF(htmlContent, nombreArchivo, tipo, sucursal) {
  const blob     = Utilities.newBlob(htmlContent, 'text/html', nombreArchivo + '.html');
  const tempFile = DriveApp.createFile(blob);
  const pdf      = tempFile.getAs(MimeType.PDF).setName(nombreArchivo + '.pdf');
  tempFile.setTrashed(true);
  // Guardar copia organizada en Drive
  guardarPDFEnDrive(pdf, nombreArchivo, tipo || 'General', sucursal || '');
  return pdf;
}

function enviarCorreoConPDF(destinatario, asunto, htmlBody, pdfBlob) {
  if (!destinatario) return;
  try {
    if (pdfBlob) {
      MailApp.sendEmail({ to: destinatario, subject: asunto, htmlBody: htmlBody, attachments: [pdfBlob], name: 'LunaMedia Ops' });
    } else {
      MailApp.sendEmail({ to: destinatario, subject: asunto, htmlBody: htmlBody, name: 'LunaMedia Ops' });
    }
  } catch(e) {
    Logger.log('Error enviando correo a ' + destinatario + ': ' + e);
  }
}

function generarHTMLCaratula(p) {
  const esPostreria  = p.sucursal === 'Postreria' || p.sucursal === 'Postrería';
  const esAsi        = (p.(folio.indexOf('-ASI') >= 0 || folio.indexOf('ASI-') === 0) || p.folio.indexOf('ASI-') === 0);
  const esExt        = p.folio.indexOf('-EXT') >= 0;
  const etiquetaFirma = esPostreria ? 'Encargado de Area' : 'Gerente / Encargado de Sucursal';
  const etiquetaInfo  = esPostreria ? 'Encargado' : 'Gerente';
  const tipoSDP      = esAsi ? 'ASIGNACION CEDIS' : (esExt ? 'EXTRAORDINARIA' : 'ORDINARIA');
  const colorTipo    = esAsi ? '#3C8EC8' : (esExt ? '#E07B2A' : '#2E9E6B');

  const lineas = (p.productos || '').split('\n').filter(function(l){ return l.trim(); });
  let filasTabla = '';
  let n = 1;
  lineas.forEach(function(linea) {
    const partes = linea.split('|');
    if (partes.length >= 5) {
      filasTabla += '<tr style="background:' + (n%2===0?'#f5f8fc':'white') + '">' +
        '<td style="padding:6px 10px;font-size:11px;">' + n + '</td>' +
        '<td style="padding:6px 10px;font-size:11px;font-family:monospace;">' + (partes[0]||'') + '</td>' +
        '<td style="padding:6px 10px;font-size:11px;">' + (partes[1]||'') + '</td>' +
        '<td style="padding:6px 10px;font-size:11px;">' + (partes[2]||'') + '</td>' +
        '<td style="padding:6px 10px;font-size:11px;text-align:right;">' + (partes[3]||'') + '</td>' +
        '<td style="padding:6px 10px;font-size:11px;text-align:right;font-weight:700;">' + (partes[4]||'').replace('x','') + '</td>' +
        '<td style="padding:6px 10px;font-size:11px;text-align:right;">' + (partes[5]||'') + '</td>' +
        '<td style="padding:6px 10px;font-size:11px;color:#6B7E93;">' + (partes[6]||'') + '</td>' +
        '</tr>';
      n++;
    }
  });

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>body{font-family:Arial,sans-serif;margin:0;padding:16px;color:#182533;line-height:1.2;}' +
    '.hdr{display:flex;align-items:center;justify-content:space-between;border-bottom:3px solid #3C8EC8;padding-bottom:8px;margin-bottom:10px;}' +
    '.brand-name{font-size:20px;font-weight:700;color:#1D4F73;}.brand-sub{font-size:9px;color:#6B7E93;text-transform:uppercase;letter-spacing:1.5px;}' +
    '.folio-num{font-size:20px;font-weight:700;color:#3C8EC8;text-align:right;}' +
    '.info-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:10px;}' +
    '.info-box{border:1px solid #EEF2F7;border-radius:6px;padding:5px 10px;}' +
    '.info-lbl{font-size:8px;text-transform:uppercase;letter-spacing:1px;color:#6B7E93;margin-bottom:1px;}' +
    '.info-val{font-size:12px;font-weight:600;line-height:1.2;}' +
    'table{width:100%;border-collapse:collapse;margin-bottom:10px;}' +
    'th{background:#1D4F73;color:white;padding:4px 8px;text-align:left;font-size:9px;line-height:1.2;}' +
    'td{padding:3px 8px;font-size:10px;line-height:1.2;}' +
    'tfoot td{font-weight:700;background:#EBF4FB;color:#1D4F73;padding:5px 8px;}' +
    '.firmas{display:grid;grid-template-columns:1fr 1fr;gap:40px;margin-top:20px;}' +
    '.firma{text-align:center;}.firma-line{border-top:2px solid #182533;margin-bottom:6px;margin-top:36px;}' +
    '.firma-name{font-size:11px;font-weight:600;}.firma-rol{font-size:9px;color:#6B7E93;}' +
    '</style></head><body>' +
    '<div class="hdr">' +
    '<div><div class="brand-name">Luna Media Cafe</div>' +
    '<div class="brand-sub">' + (esAsi ? 'Asignacion CEDIS' : 'Solicitud de Pedido - CEDIS') + '</div></div>' +
    '<div>' +
    '<div style="font-size:10px;color:#6B7E93;text-align:right;">Folio</div>' +
    '<div class="folio-num">' + p.folio + '</div>' +
    '<div style="display:inline-block;background:' + colorTipo + ';color:white;padding:2px 10px;border-radius:10px;font-size:9px;font-weight:700;margin-top:4px;">' + tipoSDP + '</div>' +
    '<div style="font-size:10px;color:#999;text-align:right;margin-top:2px;">' + p.timestamp + '</div>' +
    '</div></div>' +
    '<div class="info-grid">' +
    '<div class="info-box"><div class="info-lbl">' + (esAsi ? 'Sucursal Destino' : 'Sucursal') + '</div><div class="info-val">' + p.sucursal + '</div></div>' +
    '<div class="info-box"><div class="info-lbl">Correo</div><div class="info-val">' + p.correoSuc + '</div></div>' +
    '<div class="info-box"><div class="info-lbl">' + etiquetaInfo + '</div><div class="info-val">' + p.gerente + '</div></div>' +
    '<div class="info-box"><div class="info-lbl">Entrega Requerida</div><div class="info-val">' + p.fechaEntrega + '</div></div>' +
    '</div>' +
    '<table><thead><tr><th>#</th><th>Codigo</th><th>Producto</th><th>Presentacion</th><th>Precio</th><th>Cantidad</th><th>Subtotal</th><th>Observaciones</th></tr></thead>' +
    '<tbody>' + filasTabla + '</tbody>' +
    '<tfoot><tr><td colspan="6" style="text-align:right;">TOTAL ESTIMADO</td><td>' + fmtMoneyAS(p.totalEstimado) + '</td><td></td></tr></tfoot></table>' +
    (p.obsGenerales ? '<div style="background:#F5F8FC;border:1px solid #EEF2F7;border-radius:8px;padding:10px 14px;margin-bottom:14px;"><div style="font-size:9px;text-transform:uppercase;color:#6B7E93;margin-bottom:4px;">Observaciones</div><div style="font-size:12px;">' + p.obsGenerales + '</div></div>' : '') +
    '<div class="firmas">' +
    '<div class="firma"><div class="firma-line"></div><div class="firma-name">' + p.gerente + '</div><div class="firma-rol">' + etiquetaFirma + '</div></div>' +
    '<div class="firma"><div class="firma-line"></div><div class="firma-name">José Eduardo Muñíz Rebolledo</div><div class="firma-rol">Coordinador de CEDIS</div></div>' +
    '</div></body></html>';
}

function generarCorreoPedido(p) {
  const esPostreria = p.sucursal === 'Postreria' || p.sucursal === 'Postrería';
  const esAsi       = (p.(folio.indexOf('-ASI') >= 0 || folio.indexOf('ASI-') === 0) || p.folio.indexOf('ASI-') === 0);
  const esExt       = p.folio.indexOf('-EXT') >= 0;
  const etiqueta    = esPostreria ? 'ENCARGADO' : 'GERENTE';
  const tipoSDP     = esAsi ? 'Asignacion CEDIS' : (esExt ? 'Extraordinaria' : 'Ordinaria');
  const colorTipo   = esAsi ? '#3C8EC8' : (esExt ? '#E07B2A' : '#2E9E6B');
  const titulo      = esAsi ? 'Nueva Asignacion CEDIS' : 'Nueva Solicitud de Pedido';

  const lineas = (p.productos || '').split('\n').filter(function(l){ return l.trim(); });
  let filasCorreo = '';
  lineas.forEach(function(linea, idx) {
    const partes = linea.split('|');
    filasCorreo += '<tr style="background:' + (idx%2===0?'white':'#f5f8fc') + '">' +
      '<td style="padding:7px 12px;font-size:12px;font-family:monospace;">' + (partes[0]||'') + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;">' + (partes[1]||'') + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;">' + (partes[2]||'') + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;text-align:right;">' + (partes[3]||'') + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;text-align:right;font-weight:700;">' + (partes[4]||'').replace('x','') + '</td>' +
      '<td style="padding:7px 12px;font-size:12px;text-align:right;font-weight:700;color:#1D4F73;">' + (partes[5]||'') + '</td>' +
      '</tr>';
  });

  return '<div style="font-family:sans-serif;max-width:680px;margin:0 auto;background:#F5F8FC;">' +
    '<div style="background:linear-gradient(135deg,#1D4F73,#3C8EC8);padding:24px 28px;border-radius:12px 12px 0 0;">' +
    '<div style="color:white;font-size:22px;font-weight:700;">' + titulo + '</div>' +
    '<div style="display:inline-block;background:' + colorTipo + ';color:white;padding:2px 12px;border-radius:10px;font-size:11px;font-weight:700;margin-top:6px;">' + tipoSDP + '</div>' +
    '<div style="color:rgba(255,255,255,.8);font-size:13px;margin-top:4px;">' + p.timestamp + '</div></div>' +
    '<div style="background:white;padding:20px 28px;">' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:20px;">' +
    '<tr><td style="padding:8px 0;color:#6B7E93;font-weight:700;width:40%">FOLIO</td><td style="padding:8px 0;font-size:18px;font-weight:700;color:#3C8EC8;">' + p.folio + '</td></tr>' +
    '<tr style="background:#f5f8fc"><td style="padding:8px;color:#6B7E93;font-weight:700">' + (esAsi ? 'SUCURSAL DESTINO' : 'SUCURSAL') + '</td><td style="padding:8px;font-weight:600">' + p.sucursal + '</td></tr>' +
    '<tr><td style="padding:8px;color:#6B7E93;font-weight:700">' + etiqueta + '</td><td style="padding:8px">' + p.gerente + '</td></tr>' +
    '<tr style="background:#f5f8fc"><td style="padding:8px;color:#6B7E93;font-weight:700">ENTREGA REQUERIDA</td><td style="padding:8px;font-weight:600">' + p.fechaEntrega + '</td></tr>' +
    '<tr><td style="padding:8px;color:#6B7E93;font-weight:700">TOTAL ESTIMADO</td><td style="padding:8px;font-size:16px;font-weight:700;color:#1D4F73;">' + p.totalEstimado + '</td></tr>' +
    '</table>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<thead><tr style="background:#1D4F73;">' +
    '<th style="padding:8px 12px;color:white;font-size:11px;text-align:left;">Codigo</th>' +
    '<th style="padding:8px 12px;color:white;font-size:11px;text-align:left;">Producto</th>' +
    '<th style="padding:8px 12px;color:white;font-size:11px;text-align:left;">Presentacion</th>' +
    '<th style="padding:8px 12px;color:white;font-size:11px;text-align:right;">Precio</th>' +
    '<th style="padding:8px 12px;color:white;font-size:11px;text-align:right;">Cant.</th>' +
    '<th style="padding:8px 12px;color:white;font-size:11px;text-align:right;">Subtotal</th>' +
    '</tr></thead><tbody>' + filasCorreo + '</tbody>' +
    '<tfoot><tr style="background:#EBF4FB;">' +
    '<td colspan="5" style="padding:10px 12px;font-weight:700;text-align:right;color:#1D4F73;">TOTAL ESTIMADO</td>' +
    '<td style="padding:10px 12px;font-weight:700;color:#1D4F73;font-size:15px;">' + fmtMontoGS(p.totalEstimado) + '</td>' +
    '</tr></tfoot></table>' +
    (p.obsGenerales ? '<div style="background:#F5F8FC;border-left:4px solid #3C8EC8;padding:12px 16px;margin-top:16px;font-size:13px;color:#333;"><strong>Observaciones:</strong> ' + p.obsGenerales + '</div>' : '') +
    '</div></div>';
}

function generarCorreoEvaluacion(p) {
  if (!p) p = {};
  const esPostreria = (p.sucursal||'') === 'Postreria' || (p.sucursal||'') === 'Postrería';
  const etiqueta    = esPostreria ? 'ENCARGADO' : 'GERENTE';
  return '<div style="font-family:sans-serif;max-width:600px;margin:0 auto;">' +
    '<div style="background:#1D4F73;padding:20px;border-radius:8px 8px 0 0;">' +
    '<h2 style="color:white;margin:0;">Evaluacion de Entrega Registrada</h2></div>' +
    '<div style="background:#f5f8fc;padding:20px;border-radius:0 0 8px 8px;">' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px;">' +
    '<tr><td style="padding:8px;color:#6B7E93;font-weight:700;width:40%">FOLIO SDP</td><td style="padding:8px;font-weight:700;color:#3C8EC8">' + (p.folio||'') + '</td></tr>' +
    '<tr style="background:white"><td style="padding:8px;color:#6B7E93;font-weight:700">SUCURSAL</td><td style="padding:8px">' + (p.sucursal||'') + '</td></tr>' +
    '<tr><td style="padding:8px;color:#6B7E93;font-weight:700">CALIFICACION</td><td style="padding:8px">' + (p.calificacion||'') + '/5</td></tr>' +
    '<tr style="background:white"><td style="padding:8px;color:#6B7E93;font-weight:700">TOTAL ORIGINAL</td><td style="padding:8px">' + (p.totalOriginal||'') + '</td></tr>' +
    '<tr><td style="padding:8px;color:#6B7E93;font-weight:700">TOTAL REAL</td><td style="padding:8px;font-weight:700;color:#2E9E6B">' + (p.totalReal||'') + '</td></tr>' +
    '<tr style="background:white"><td style="padding:8px;color:#6B7E93;font-weight:700">FALTANTES</td><td style="padding:8px;color:#D4504A">' + (p.faltantes || 'Ninguno') + '</td></tr>' +
    '<tr><td style="padding:8px;color:#6B7E93;font-weight:700">COMENTARIOS</td><td style="padding:8px;font-style:italic">' + (p.comentarios || 'Sin comentarios') + '</td></tr>' +
    '</table></div></div>';
}


// ═══════════════════════════════════════════════════════════════
// SISTEMA DE ALERTAS Y SEMAFORO
// ═══════════════════════════════════════════════════════════════

const CDP_EMAIL   = 'cdplunamediacafe@gmail.com';
const TODOS_EMAILS = [ADMIN_EMAIL, CEO_EMAIL, CDP_EMAIL];

// Semaforo de colores
const SEMAFORO = {
  CRITICO: { color: '#D4504A', bg: '#FEF5F5', label: '🔴 CRÍTICO',   border: '#D4504A' },
  ALERTA:  { color: '#E07B2A', bg: '#FEF3E8', label: '🟠 ALERTA',    border: '#E07B2A' },
  AVISO:   { color: '#C9A800', bg: '#FFFBE6', label: '🟡 AVISO',     border: '#C9A800' },
  VERDE:   { color: '#2E9E6B', bg: '#E8F7F1', label: '🟢 EXCELENTE', border: '#2E9E6B' }
};

function alertaHTML(nivel, titulo, sucursal, fecha, detalle, extra) {
  const s = SEMAFORO[nivel];
  return '<div style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto;">' +
    '<div style="background:' + s.color + ';padding:20px 28px;border-radius:12px 12px 0 0;">' +
    '<div style="color:white;font-size:11px;letter-spacing:2px;text-transform:uppercase;opacity:.8;margin-bottom:4px;">Luna Media Cafe — Sistema de Alertas</div>' +
    '<div style="color:white;font-size:22px;font-weight:700;">' + s.label + ' — ' + titulo + '</div>' +
    '<div style="color:rgba(255,255,255,.8);font-size:13px;margin-top:4px;">' + fecha + '</div></div>' +
    '<div style="background:white;padding:20px 28px;">' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px;">' +
    '<tr style="background:' + s.bg + '"><td style="padding:8px;color:' + s.color + ';font-weight:700;width:40%">NIVEL</td><td style="padding:8px;font-weight:700;color:' + s.color + '">' + s.label + '</td></tr>' +
    '<tr><td style="padding:8px;color:#6B7E93;font-weight:700">SUCURSAL / AREA</td><td style="padding:8px;font-weight:600">' + sucursal + '</td></tr>' +
    '</table>' +
    '<div style="background:' + s.bg + ';border-left:4px solid ' + s.border + ';border-radius:0 8px 8px 0;padding:14px 16px;margin-bottom:16px;">' +
    '<div style="font-size:13px;color:#182533;">' + detalle + '</div>' +
    '</div>' +
    (extra || '') +
    '</div>' +
    '<div style="padding:12px 28px;text-align:center;font-size:10px;color:#bbb;border-top:1px solid #EEF2F7;">Luna Media Cafe · Sistema de Alertas Automaticas</div>' +
    '</div>';
}

function enviarAlerta(destinatarios, asunto, htmlBody) {
  destinatarios.forEach(function(to) {
    if (!to) return;
    try { MailApp.sendEmail({ to: to, subject: asunto, htmlBody: htmlBody, name: 'LunaMedia Ops — Alertas' }); }
    catch(e) { Logger.log('Error alerta a ' + to + ': ' + e); }
  });
}

// ── ALERTA 1: Entrega incompleta o rechazada ──
function alertaEntrega(p) {
  const estadoEval = (p.estado || p.estadoEntrega || '').toLowerCase();
  const esRechazado = estadoEval === 'rechazado' || estadoEval === 'rechazada';
  const nivel = esRechazado ? 'CRITICO' : 'ALERTA';
  const titulo = esRechazado ? 'Entrega Rechazada' : 'Entrega Incompleta';
  const fecha = p.timestamp || new Date().toLocaleString('es-MX');

  const detalle = esRechazado
    ? 'La entrega del folio <strong>' + p.folio + '</strong> fue <strong>RECHAZADA</strong> en su totalidad por ' + p.sucursal + '. Se requiere accion inmediata de CEDIS.'
    : 'La entrega del folio <strong>' + p.folio + '</strong> fue recibida de forma <strong>INCOMPLETA</strong> en ' + p.sucursal + '. Total original: ' + (p.totalOriginal||'—') + ' | Total real recibido: ' + (p.totalReal||'—') + '.';

  const extra = '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
    '<tr><td style="padding:8px;color:#6B7E93;font-weight:700;width:40%">FOLIO</td><td style="padding:8px;font-weight:700;color:#3C8EC8">' + p.folio + '</td></tr>' +
    '<tr style="background:#f5f8fc"><td style="padding:8px;color:#6B7E93;font-weight:700">CALIFICACION</td><td style="padding:8px">' + (p.calificacion||'—') + '/5</td></tr>' +
    '<tr><td style="padding:8px;color:#6B7E93;font-weight:700">FALTANTES</td><td style="padding:8px;color:#D4504A">' + (p.faltantes||'Sin detalle') + '</td></tr>' +
    '<tr style="background:#f5f8fc"><td style="padding:8px;color:#6B7E93;font-weight:700">COMENTARIOS</td><td style="padding:8px;font-style:italic">' + (p.comentarios||'Sin comentarios') + '</td></tr>' +
    '</table>';

  const asunto = '[' + SEMAFORO[nivel].label + '] ' + titulo + ' — ' + p.sucursal + ' | ' + p.folio;
  const html = alertaHTML(nivel, titulo, p.sucursal, fecha, detalle, extra);
  enviarAlerta([CDP_EMAIL, ADMIN_EMAIL, CEO_EMAIL], asunto, html);
  Logger.log('Alerta entrega enviada: ' + nivel + ' - ' + p.folio);
}

// ── ALERTA 2: Insumos con problemas de calidad (merma/rechazo) ──
function alertaCalidad(sucursal, productos, motivo, fecha) {
  const nivel = motivo.indexOf('caducidad') >= 0 || motivo.indexOf('vencid') >= 0 ? 'CRITICO' : 'ALERTA';
  const titulo = 'Problema de Calidad en Insumos';
  const detalle = 'Se han registrado insumos con problemas en <strong>' + sucursal + '</strong>.<br>' +
    'Motivo reportado: <strong>' + motivo + '</strong><br>' +
    'Productos afectados: ' + productos;
  const asunto = '[' + SEMAFORO[nivel].label + '] ' + titulo + ' — ' + sucursal;
  const html = alertaHTML(nivel, titulo, sucursal, fecha, detalle, '');
  enviarAlerta([CDP_EMAIL, ADMIN_EMAIL, CEO_EMAIL], asunto, html);
  Logger.log('Alerta calidad enviada: ' + sucursal);
}

// ── ALERTA 3: Productos en cero en inventario ──
function alertaInventarioCeros(sucursal, productosCero, fecha, folio) {
  const nivel = productosCero.length > 10 ? 'ALERTA' : 'AVISO';
  const titulo = 'Productos en Cero — ' + sucursal;
  const detalle = 'El inventario <strong>' + folio + '</strong> de <strong>' + sucursal + '</strong> registra <strong>' + productosCero.length + ' producto(s) sin existencia</strong>:';
  const lista = '<ul style="margin:10px 0;padding-left:20px;">' + productosCero.slice(0, 20).map(function(p){ return '<li style="font-size:12px;color:#182533;margin-bottom:3px;">' + p + '</li>'; }).join('') + (productosCero.length > 20 ? '<li style="color:#6B7E93;">...y ' + (productosCero.length-20) + ' mas</li>' : '') + '</ul>';
  const asunto = '[' + SEMAFORO[nivel].label + '] ' + titulo + ' | ' + folio;
  const html = alertaHTML(nivel, titulo, sucursal, fecha, detalle, lista);
  enviarAlerta([CDP_EMAIL, ADMIN_EMAIL, CEO_EMAIL], asunto, html);
  Logger.log('Alerta ceros enviada: ' + productosCero.length + ' productos - ' + sucursal);
}

// ── ALERTA 4: Reconocimiento mensual — Sucursales y CEDIS por separado ──
function evaluacionMensual() {
  const ss         = SpreadsheetApp.openById('1hqpPHPr9msdhqjN3WBMYGrIH9XxRVmdoayhcscipNJg');
  const ahora      = new Date();
  const mes        = ahora.getMonth();
  const anio       = ahora.getFullYear();
  const meses      = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  const nombreMes  = meses[mes];
  const periodoKey = anio + '-' + String(mes+1).padStart(2,'0');
  const primerDia  = new Date(anio, mes, 1);
  const ultimoDia  = new Date(anio, mes + 1, 0);

  // Solo procesar desde Mayo 2026
  if (anio < 2026 || (anio === 2026 && mes < 4)) {
    Logger.log('Evaluacion mensual: periodo anterior a Mayo 2026, omitiendo.');
    return;
  }

  function enMes(f) {
    var d = (f instanceof Date) ? f : new Date(f);
    return !isNaN(d) && d >= primerDia && d <= ultimoDia;
  }
  function fmtMoney(n) {
    return '$' + parseFloat(n||0).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g,',');
  }
  function pct(a,b) { return b>0 ? Math.round((a/b)*100) : 0; }
  function nivelScore(s) {
    return s>=90?{label:'Excelente',icon:'🟢',color:'#2E9E6B',bg:'#E8F7F1'}:
           s>=75?{label:'Bueno',icon:'🔵',color:'#185FA5',bg:'#E6F1FB'}:
           s>=60?{label:'Regular',icon:'🟠',color:'#854F0B',bg:'#FAEEDA'}:
                 {label:'Deficiente',icon:'🔴',color:'#A32D2D',bg:'#FCEBEB'};
  }

  const tabsSucs = [
    {nombre:'Avila Camacho',  tab:'ÁVILA CAMACHO',  correo:'lunamediacentro@gmail.com'},
    {nombre:'Paseo Jardines', tab:'PASEO JARDINES',  correo:'lunamediapaseojardines@gmail.com'},
    {nombre:'Postreria',      tab:'POSTRERÍA',       correo:'postrerialunamediacafe@gmail.com'}
  ];

  // ── HISTORIAL FILE ──
  const folder = DriveApp.getFolderById(HISTORIAL_FOLDER);
  function getHistFile(nombre) {
    var fname = 'historial_eval_' + nombre.replace(/ /g,'_') + '.json';
    var files = folder.getFilesByName(fname);
    if (files.hasNext()) return files.next();
    return folder.createFile(fname, '{}', 'application/json');
  }
  function loadHist(nombre) {
    try { return JSON.parse(getHistFile(nombre).getBlob().getDataAsString()); } catch(e) { return {}; }
  }
  function saveHist(nombre, data) {
    getHistFile(nombre).setContent(JSON.stringify(data));
  }

  // ══════════════════════════════════════════
  // EVALUACIÓN SUCURSALES
  // ══════════════════════════════════════════
  var resSucs = [];

  tabsSucs.forEach(function(suc) {
    var sheet = ss.getSheetByName(suc.tab);
    var hist  = loadHist(suc.nombre);

    // Datos SDP del mes
    var totalSDP=0, evalATiempo=0, evalVencidas=0, totalExt=0, totalOrd=0;
    var rechazos={}, motivosRech=[];
    var loSolicitado={};

    if (sheet && sheet.getLastRow() > 1) {
      var data = sheet.getDataRange().getValues();
      data.slice(1).forEach(function(row) {
        if (!row[0] || !enMes((row[1] instanceof Date)?row[1]:new Date(row[1]))) return;
        totalSDP++;
        var folio  = String(row[0]);
        var estado = String(row[10]||'');
        if (folio.indexOf('-EXT')>=0) totalExt++;
        else totalOrd++;
        if (estado==='EVALUADA') evalATiempo++;
        if (estado==='VENCIDA')  evalVencidas++;
        if (estado==='EVALUADA' && row[13]==='0') {
          var motivo = String(row[16]||'Sin motivo especificado');
          rechazos[motivo] = (rechazos[motivo]||0)+1;
        }
        // Productos solicitados (del resumen de productos en col 9 aprox)
        var prods = String(row[8]||'');
        prods.split(';').forEach(function(p){
          var partes = p.split('|');
          if(partes.length>=2){ var nm=partes[1]||partes[0]; if(nm) loSolicitado[nm]=(loSolicitado[nm]||0)+1; }
        });
      });
    }

    // Datos inventario del mes (desde historial_inventarios.json)
    var pctSinCeros = 0;
    try {
      var invFiles = folder.getFilesByName(HISTORIAL_FILE);
      if (invFiles.hasNext()) {
        var invData = JSON.parse(invFiles.next().getBlob().getDataAsString());
        var invMes  = invData.filter(function(i){ return i.sucursal===suc.nombre && enMes(new Date(i.timestamp||0)); });
        if (invMes.length > 0) {
          var totalProds=0, conExistencia=0;
          invMes.forEach(function(inv){
            var c = inv.conteo||{};
            Object.values(c).forEach(function(v){ totalProds++; if(parseFloat(v||0)>0) conExistencia++; });
          });
          pctSinCeros = pct(conExistencia, totalProds);
        }
      }
    } catch(eInv){ Logger.log('Error leyendo inventarios para ' + suc.nombre + ': ' + eInv); }

    // ── Score sucursal ──
    var scoreEval     = totalSDP>0 ? Math.round((evalATiempo/totalSDP)*100*0.35) : 0;
    var scoreInv      = Math.round(pctSinCeros * 0.30);
    var scoreOrd      = totalSDP>0 ? Math.round((1-(totalExt/totalSDP))*100*0.20) : 20;
    var scoreForms    = totalSDP>0 ? Math.round(((totalSDP-evalVencidas)/totalSDP)*100*0.15) : 15;
    var scoreTotal    = scoreEval + scoreInv + scoreOrd + scoreForms;

    // Top/bottom productos
    var topSol  = Object.entries(loSolicitado).sort((a,b)=>b[1]-a[1]).slice(0,5);
    var botSol  = Object.entries(loSolicitado).sort((a,b)=>a[1]-b[1]).slice(0,5);

    // ── Actualizar historial ──
    hist[periodoKey] = {
      score:scoreTotal, totalSDP:totalSDP, evalATiempo:evalATiempo,
      evalVencidas:evalVencidas, totalExt:totalExt, pctSinCeros:pctSinCeros,
      rechazos:rechazos, topSolicitados:topSol, bottomSolicitados:botSol
    };

    // Calcular máximos/mínimos históricos del score
    var scores = Object.values(hist).map(function(h){return h.score||0;}).filter(Boolean);
    hist._maxScore = scores.length>0 ? Math.max.apply(null,scores) : scoreTotal;
    hist._minScore = scores.length>0 ? Math.min.apply(null,scores) : scoreTotal;
    hist._tendencia = Object.keys(hist).filter(function(k){return k!=='_maxScore'&&k!=='_minScore'&&k!=='_tendencia';}).length >= 2 ?
      (scoreTotal > (Object.values(hist).filter(function(v){return typeof v==='object'&&v.score;}).slice(-2,-1)[0]||{score:scoreTotal}).score ? 'mejora' : 'baja') : 'neutro';
    saveHist(suc.nombre, hist);

    // Meses anteriores para tendencia
    var periodos = Object.keys(hist).filter(function(k){return k!=='_maxScore'&&k!=='_minScore'&&k!=='_tendencia';}).sort();
    var scoreAnterior = periodos.length>=2 ? (hist[periodos[periodos.length-2]]||{}).score||0 : 0;

    resSucs.push({
      nombre:suc.nombre, correo:suc.correo, score:scoreTotal,
      nivel:nivelScore(scoreTotal), scoreEval:Math.round(scoreEval/0.35),
      scoreInv:pctSinCeros, scoreOrd:Math.round(scoreOrd/0.20),
      totalSDP:totalSDP, evalATiempo:evalATiempo, evalVencidas:evalVencidas,
      totalExt:totalExt, rechazos:rechazos, topSol:topSol, botSol:botSol,
      maxScore:hist._maxScore, minScore:hist._minScore,
      tendencia:hist._tendencia, scoreAnterior:scoreAnterior
    });
  });

  // ══════════════════════════════════════════
  // EVALUACIÓN CEDIS
  // ══════════════════════════════════════════
  var cedisHist = loadHist('CEDIS');
  var cTotalEnt=0,cCompletas=0,cIncomp=0,cRech=0,cEvalVenc=0;
  var cSumTiempo=0,cCntTiempo=0,cMontoEst=0,cMontoReal=0;
  var cedFaltantes={};

  tabsSucs.forEach(function(suc) {
    var sheet = ss.getSheetByName(suc.tab);
    if (!sheet || sheet.getLastRow()<=1) return;
    var data = sheet.getDataRange().getValues();
    data.slice(1).forEach(function(row) {
      if (!row[0] || !enMes((row[1] instanceof Date)?row[1]:new Date(row[1]))) return;
      var estado = String(row[10]||'');
      if (estado==='CANCELADA') return;
      var est  = parseFloat(String(row[7]||'0').replace(/[$,]/g,''))||0;
      var real = parseFloat(String(row[13]||row[7]||'0').replace(/[$,]/g,''))||0;
      if (estado==='EVALUADA') {
        cTotalEnt++; cMontoEst+=est; cMontoReal+=real;
        if (real===0)      { cRech++; }
        else if (real<est) { cIncomp++; }
        else               { cCompletas++; }
        // Tiempo
        var hIn=String(row[17]||''),hOut=String(row[19]||'');
        if (hIn&&hOut) {
          var pIn=hIn.split(':'),pOut=hOut.split(':');
          if(pIn.length>=2&&pOut.length>=2){
            var mIn=parseInt(pIn[0])*60+parseInt(pIn[1]);
            var mOut=parseInt(pOut[0])*60+parseInt(pOut[1]);
            var diff=mOut-mIn;
            if(diff>0&&diff<480){cSumTiempo+=diff;cCntTiempo++;}
          }
        }
        // Faltantes
        var falt=String(row[14]||'');
        if(falt) falt.split(';').forEach(function(f){
          var n=f.split('-').slice(1).join('-').trim()||f.trim();
          if(n) cedFaltantes[n]=(cedFaltantes[n]||0)+1;
        });
      }
      if (estado==='VENCIDA') cEvalVenc++;
    });
  });

  var tiempoProm    = cCntTiempo>0 ? Math.round(cSumTiempo/cCntTiempo) : 0;
  var diffMonto     = cMontoReal - cMontoEst;
  var pctCompletas  = pct(cCompletas, cTotalEnt);
  var pctConc       = cMontoEst>0 ? Math.round(Math.min(cMontoReal/cMontoEst,1)*100) : 100;
  var pctSinVenc    = cTotalEnt>0 ? Math.round(((cTotalEnt-cEvalVenc)/cTotalEnt)*100) : 100;
  var scoreTiempo   = tiempoProm===0?100:tiempoProm<=45?100:tiempoProm<=60?85:tiempoProm<=75?70:tiempoProm<=90?50:20;

  var scoreCedis    = Math.round(
    (pctCompletas * 0.35) +
    (scoreTiempo  * 0.25) +
    (pctConc      * 0.25) +
    (pctSinVenc   * 0.15)
  );

  var topFalt = Object.entries(cedFaltantes).sort((a,b)=>b[1]-a[1]).slice(0,8);

  cedisHist[periodoKey] = {
    score:scoreCedis, totalEnt:cTotalEnt, completas:cCompletas,
    incompletas:cIncomp, rechazadas:cRech, evalVencidas:cEvalVenc,
    tiempoProm:tiempoProm, montoEst:cMontoEst, montoReal:cMontoReal,
    diferencia:diffMonto, topFaltantes:topFalt
  };
  var cScores = Object.values(cedisHist).map(function(h){return h.score||0;}).filter(Boolean);
  cedisHist._maxScore = cScores.length>0?Math.max.apply(null,cScores):scoreCedis;
  cedisHist._minScore = cScores.length>0?Math.min.apply(null,cScores):scoreCedis;
  var cPeriodos = Object.keys(cedisHist).filter(function(k){return k!=='_maxScore'&&k!=='_minScore'&&k!=='_tendencia';}).sort();
  var cScoreAnt = cPeriodos.length>=2?(cedisHist[cPeriodos[cPeriodos.length-2]]||{}).score||0:0;
  cedisHist._tendencia = cPeriodos.length>=2?(scoreCedis>cScoreAnt?'mejora':'baja'):'neutro';
  saveHist('CEDIS', cedisHist);

  var nivelCedis = nivelScore(scoreCedis);

  // ══════════════════════════════════════════
  // GENERAR CORREO HTML
  // ══════════════════════════════════════════
  function barra(pct, color) {
    var w = Math.min(Math.max(pct,0),100);
    return '<div style="display:inline-block;vertical-align:middle;width:80px;height:8px;background:#EEF2F7;border-radius:4px;overflow:hidden;">' +
           '<div style="width:'+w+'%;height:8px;background:'+color+';border-radius:4px;"></div></div>';
  }
  function tendIcon(t) { return t==='mejora'?'↑':t==='baja'?'↓':'→'; }

  var encabezado = '<div style="font-family:Arial,sans-serif;max-width:720px;margin:0 auto;background:#F5F8FC;">' +
    '<div style="background:#0C447C;padding:28px 32px;border-radius:12px 12px 0 0;">' +
    '<p style="color:#85B7EB;font-size:11px;letter-spacing:2px;text-transform:uppercase;margin:0 0 6px;">Luna Media Café · Evaluación Mensual Automática</p>' +
    '<p style="color:#fff;font-size:22px;font-weight:700;margin:0 0 4px;">Reporte de Desempeño — ' + nombreMes + ' ' + anio + '</p>' +
    '<p style="color:#B5D4F4;font-size:13px;margin:0;">Período: 1 al 30 de ' + nombreMes + ' · Generado automáticamente el 28/' + nombreMes + '/' + anio + '</p>' +
    '</div><div style="background:white;padding:28px 32px;">';

  // ── Sección SUCURSALES ──
  var secSucs = '<p style="font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:1.5px;color:#6B7E93;border-bottom:2px solid #EEF2F7;padding-bottom:8px;margin:0 0 16px;">Evaluación de Sucursales</p>';

  resSucs.forEach(function(r) {
    var n = r.nivel;
    var tendTxt = r.tendencia==='mejora'?'<span style="color:#2E9E6B;">'+tendIcon(r.tendencia)+' Mejoró vs mes anterior ('+r.scoreAnterior+'pts)</span>':
                  r.tendencia==='baja'?'<span style="color:#D4504A;">'+tendIcon(r.tendencia)+' Bajó vs mes anterior ('+r.scoreAnterior+'pts)</span>':
                  '<span style="color:#6B7E93;">→ Primer período registrado</span>';

    secSucs += '<div style="border:1px solid #EEF2F7;border-radius:10px;padding:16px;margin-bottom:14px;border-left:4px solid '+n.color+';">' +
      '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">' +
      '<div><span style="font-size:16px;font-weight:700;color:#182533;">'+r.nombre+'</span> &nbsp;'+
      '<span style="background:'+n.bg+';color:'+n.color+';font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;">'+n.icon+' '+n.label+'</span></div>' +
      '<div style="text-align:right;"><span style="font-size:28px;font-weight:700;color:'+n.color+';">'+r.score+'</span><span style="font-size:12px;color:#6B7E93;">/100</span><br>'+tendTxt+'</div>' +
      '</div>' +
      '<table style="width:100%;font-size:12px;border-collapse:collapse;">' +
      '<tr><td style="padding:5px 0;color:#6B7E93;width:40%;">Evaluaciones a tiempo (35pts)</td><td>'+barra(r.scoreEval,'#2E9E6B')+'</td><td style="padding:5px 8px;font-weight:700;color:#182533;">'+r.scoreEval+'%</td><td style="color:#6B7E93;">'+r.evalATiempo+'/'+r.totalSDP+' en plazo</td></tr>' +
      '<tr><td style="padding:5px 0;color:#6B7E93;">Inventario sin ceros (30pts)</td><td>'+barra(r.scoreInv,'#185FA5')+'</td><td style="padding:5px 8px;font-weight:700;color:#182533;">'+r.scoreInv+'%</td><td style="color:#6B7E93;">productos con existencia</td></tr>' +
      '<tr><td style="padding:5px 0;color:#6B7E93;">Pedidos ordinarios (20pts)</td><td>'+barra(r.scoreOrd,'#854F0B')+'</td><td style="padding:5px 8px;font-weight:700;color:#182533;">'+r.scoreOrd+'%</td><td style="color:#6B7E93;">'+r.totalExt+' EXT de '+r.totalSDP+' total</td></tr>' +
      '<tr><td style="padding:5px 0;color:#6B7E93;">Sin evaluaciones vencidas (15pts)</td><td>'+barra(r.evalVencidas===0?100:Math.round((1-r.evalVencidas/r.totalSDP)*100),'#534AB7')+'</td><td style="padding:5px 8px;font-weight:700;color:#182533;">'+(r.evalVencidas===0?'100':''+Math.round((1-r.evalVencidas/r.totalSDP)*100))+'%</td><td style="color:#6B7E93;">'+r.evalVencidas+' vencidas</td></tr>' +
      '</table>';

    // Histórico
    secSucs += '<div style="margin-top:10px;padding-top:10px;border-top:1px dashed #EEF2F7;font-size:11px;color:#6B7E93;">'+
      'Histórico: máx <strong style="color:#2E9E6B;">'+r.maxScore+'pts</strong> &nbsp;·&nbsp; mín <strong style="color:#D4504A;">'+r.minScore+'pts</strong></div>';

    // Top solicitados
    if (r.topSol && r.topSol.length>0) {
      secSucs += '<div style="margin-top:10px;font-size:11px;"><span style="color:#6B7E93;">Más solicitados: </span>' +
        r.topSol.map(function(p){return '<strong>'+p[0]+'</strong> ('+p[1]+'x)';}).join(' · ') + '</div>';
    }
    // Rechazos
    var rechArr = Object.entries(r.rechazos||{});
    if (rechArr.length>0) {
      secSucs += '<div style="margin-top:6px;font-size:11px;color:#D4504A;"><span style="color:#6B7E93;">Rechazos: </span>' +
        rechArr.map(function(e){return e[0]+' ('+e[1]+'x)';}).join(' · ') + '</div>';
    }
    secSucs += '</div>';
  });

  // ── Sección CEDIS ──
  var secCedis = '<p style="font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:1.5px;color:#6B7E93;border-bottom:2px solid #EEF2F7;padding-bottom:8px;margin:24px 0 16px;">Evaluación CEDIS</p>' +
    '<div style="border:1px solid #EEF2F7;border-radius:10px;padding:16px;margin-bottom:14px;border-left:4px solid '+nivelCedis.color+';">' +
    '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">' +
    '<div><span style="font-size:16px;font-weight:700;color:#182533;">Centro de Distribución CEDIS</span> &nbsp;'+
    '<span style="background:'+nivelCedis.bg+';color:'+nivelCedis.color+';font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;">'+nivelCedis.icon+' '+nivelCedis.label+'</span></div>'+
    '<div style="text-align:right;"><span style="font-size:28px;font-weight:700;color:'+nivelCedis.color+';">'+scoreCedis+'</span><span style="font-size:12px;color:#6B7E93;">/100</span></div>'+
    '</div>'+
    '<table style="width:100%;font-size:12px;border-collapse:collapse;">'+
    '<tr><td style="padding:5px 0;color:#6B7E93;width:40%;">Entregas completas (35pts)</td><td>'+barra(pctCompletas,'#2E9E6B')+'</td><td style="padding:5px 8px;font-weight:700;color:#182533;">'+pctCompletas+'%</td><td style="color:#6B7E93;">'+cCompletas+' de '+cTotalEnt+' completas</td></tr>'+
    '<tr><td style="padding:5px 0;color:#6B7E93;">Tiempo de entrega (25pts)</td><td>'+barra(scoreTiempo,'#185FA5')+'</td><td style="padding:5px 8px;font-weight:700;color:#182533;">'+(tiempoProm>0?tiempoProm+' min':'—')+'</td><td style="color:#6B7E93;">meta: ≤45 min</td></tr>'+
    '<tr><td style="padding:5px 0;color:#6B7E93;">Conciliación de montos (25pts)</td><td>'+barra(pctConc,'#854F0B')+'</td><td style="padding:5px 8px;font-weight:700;color:#182533;">'+pctConc+'%</td><td style="color:#6B7E93;">diferencia: '+fmtMoney(diffMonto)+'</td></tr>'+
    '<tr><td style="padding:5px 0;color:#6B7E93;">Evaluaciones vigentes (15pts)</td><td>'+barra(pctSinVenc,'#534AB7')+'</td><td style="padding:5px 8px;font-weight:700;color:#182533;">'+pctSinVenc+'%</td><td style="color:#6B7E93;">'+cEvalVenc+' vencidas</td></tr>'+
    '</table>';

  // Top faltantes CEDIS
  if (topFalt.length>0) {
    var maxF = topFalt[0][1];
    secCedis += '<div style="margin-top:12px;padding-top:10px;border-top:1px dashed #EEF2F7;">'+
      '<p style="font-size:11px;font-weight:700;color:#6B7E93;margin:0 0 8px;text-transform:uppercase;letter-spacing:1px;">Productos con mayor frecuencia de faltantes</p>';
    topFalt.slice(0,5).forEach(function(p){
      var w = Math.round((p[1]/maxF)*100);
      secCedis += '<div style="display:flex;align-items:center;gap:8px;margin-bottom:5px;font-size:11px;">'+
        '<span style="width:140px;color:#182533;">'+p[0]+'</span>'+
        '<div style="flex:1;height:6px;background:#EEF2F7;border-radius:3px;overflow:hidden;"><div style="width:'+w+'%;height:6px;background:#D4504A;border-radius:3px;"></div></div>'+
        '<span style="color:#D4504A;font-weight:700;width:20px;">'+p[1]+'x</span></div>';
    });
    secCedis += '</div>';
  }

  // Histórico CEDIS
  secCedis += '<div style="margin-top:10px;padding-top:10px;border-top:1px dashed #EEF2F7;font-size:11px;color:#6B7E93;">'+
    'Histórico: máx <strong style="color:#2E9E6B;">'+cedisHist._maxScore+'pts</strong> &nbsp;·&nbsp; mín <strong style="color:#D4504A;">'+cedisHist._minScore+'pts</strong></div>';
  secCedis += '</div>';

  // Incompletas/rechazadas detalle
  if (cIncomp>0||cRech>0) {
    secCedis += '<div style="background:#FDECEA;border-left:4px solid #D4504A;padding:12px 16px;border-radius:0 8px 8px 0;margin-bottom:14px;">'+
      '<p style="font-size:12px;font-weight:700;color:#A32D2D;margin:0 0 4px;">Incidencias CEDIS</p>'+
      '<p style="font-size:12px;color:#791F1F;margin:0;">'+cIncomp+' entrega(s) incompleta(s) · '+cRech+' rechazada(s). Revisar proceso de preparación y verificación previa al despacho.</p>'+
      '</div>';
  }

  // Firma
  var firma = '<div style="margin-top:24px;padding-top:16px;border-top:1px solid #EEF2F7;">'+
    '<p style="font-size:12px;color:#182533;font-style:italic;margin:0 0 10px;">"La mejora continua no es accidental — es el resultado de medirla, comunicarla y actuar sobre ella."</p>'+
    '<p style="font-size:13px;font-weight:700;color:#1D4F73;margin:0;">Uriel Rodríguez</p>'+
    '<p style="font-size:11px;color:#6B7E93;margin:0;">Director de Operaciones · Luna Media Café</p>'+
    '</div></div>'+
    '<div style="padding:10px 32px;text-align:center;font-size:10px;color:#bbb;border-top:1px solid #EEF2F7;background:#F5F8FC;">'+
    'Evaluación automática mensual · Luna Media Café · '+nombreMes+' '+anio+'</div></div>';

  var cuerpo = encabezado + secSucs + secCedis + firma;
  var asunto = '📊 Evaluación Mensual ' + nombreMes + ' ' + anio + ' — Sucursales & CEDIS';

  var todosCorreos = [ADMIN_EMAIL, CEO_EMAIL, CDP_EMAIL, CEDIS_EMAIL,
    'lunamediacentro@gmail.com','lunamediapaseojardines@gmail.com','postrerialunamediacafe@gmail.com','rhlunamediacafe@gmail.com'];
  todosCorreos.forEach(function(to){
    try { MailApp.sendEmail({to:to, subject:asunto, htmlBody:cuerpo, name:'LunaMedia Ops'}); }
    catch(e){ Logger.log('Error correo eval a '+to+': '+e); }
  });
  Logger.log('Evaluacion mensual '+nombreMes+' '+anio+' completada. CEDIS: '+scoreCedis+'pts');
}






// ═══════════════════════════════════════════════════════════════
// SESIONES Y ESPEJO EN TIEMPO REAL (CacheService)
// ═══════════════════════════════════════════════════════════════

// Registra o actualiza sesión activa (heartbeat cada 45s)
function sesionPing(p) {
  try {
    var email = p.email || '';
    var tabId = p.tabId || '';
    if (!email || !tabId) return error('datos incompletos');
    var cache = CacheService.getScriptCache();
    var key   = 'ses_' + email;
    // TTL de 90 segundos — si no hay ping en ese tiempo expira solo
    cache.put(key, JSON.stringify({ email: email, tabId: tabId, ts: new Date().getTime() }), 90);
    return ok('ping ok');
  } catch(e) { return error(e.toString()); }
}

// Verifica si hay sesión activa para este email con diferente tabId
function sesionCheck(p) {
  try {
    var email = p.email || '';
    var tabId = p.tabId || '';
    if (!email) return jsonR({ activa: false });
    var cache   = CacheService.getScriptCache();
    var cached  = cache.get('ses_' + email);
    if (!cached) return jsonR({ activa: false });
    var sesion  = JSON.parse(cached);
    // Si el tabId es el mismo, es la misma pestaña — no es espejo
    if (sesion.tabId === tabId) return jsonR({ activa: false });
    return jsonR({ activa: true, tabId: sesion.tabId });
  } catch(e) { return jsonR({ activa: false }); }
}

// Cierra sesión
function sesionCerrar(p) {
  try {
    var email = p.email || '';
    if (!email) return ok('noop');
    CacheService.getScriptCache().remove('ses_' + email);
    CacheService.getScriptCache().remove('espejo_' + email);
    return ok('sesion cerrada');
  } catch(e) { return error(e.toString()); }
}

// Publica estado del espejo (qué hace la sesión principal)
function espejoPub(p) {
  try {
    var email = p.email || '';
    if (!email) return error('sin email');
    var estado = {
      modulo:   p.modulo   || '',
      paso:     p.paso     || '',
      datos:    p.datos    || '{}',
      ts:       new Date().getTime()
    };
    // TTL 90 segundos
    CacheService.getScriptCache().put('espejo_' + email, JSON.stringify(estado), 90);
    return ok('estado publicado');
  } catch(e) { return error(e.toString()); }
}

// Obtiene estado del espejo para mostrar al dispositivo 2
function espejoGet(p) {
  try {
    var email  = p.email || '';
    var cached = CacheService.getScriptCache().get('espejo_' + email);
    if (!cached) return jsonR({ modulo: '', paso: '', datos: '{}' });
    return jsonR(JSON.parse(cached));
  } catch(e) { return jsonR({ modulo: '', paso: '', datos: '{}' }); }
}

// ═══════════════════════════════════════════════════════════════
// BORRADOR INVENTARIO (autoguardado)
// ═══════════════════════════════════════════════════════════════
const BORRADOR_FILE = 'borradores_inventario.json';

function getBorradorFile() {
  const folder = DriveApp.getFolderById(HISTORIAL_FOLDER);
  const files = folder.getFilesByName(BORRADOR_FILE);
  if (files.hasNext()) return files.next();
  return folder.createFile(BORRADOR_FILE, '{}', 'application/json');
}

function guardarBorrador(p) {
  try {
    const file = getBorradorFile();
    const data = JSON.parse(file.getBlob().getDataAsString());
    const key = (p.sucursal || '').replace(/[^a-zA-Z]/g, '_');
    data[key] = JSON.parse(p.datos || '{}');
    file.setContent(JSON.stringify(data));
    return ok('borrador guardado');
  } catch(e) { return error(e.toString()); }
}

function getBorrador(sucursal) {
  try {
    const file = getBorradorFile();
    const data = JSON.parse(file.getBlob().getDataAsString());
    const key = (sucursal || '').replace(/[^a-zA-Z]/g, '_');
    const borrador = data[key];
    if (!borrador) return jsonR({});
    // Auto-delete if older than 24hrs
    const hace24hrs = new Date().getTime() - (24 * 60 * 60 * 1000);
    if ((borrador.timestamp || 0) < hace24hrs) {
      delete data[key];
      file.setContent(JSON.stringify(data));
      return jsonR({});
    }
    return jsonR(borrador);
  } catch(e) { return jsonR({}); }
}

function eliminarBorrador(sucursal) {
  try {
    const file = getBorradorFile();
    const data = JSON.parse(file.getBlob().getDataAsString());
    const key = (sucursal || '').replace(/[^a-zA-Z]/g, '_');
    delete data[key];
    file.setContent(JSON.stringify(data));
    return ok('borrador eliminado');
  } catch(e) { return error(e.toString()); }
}

// ═══════════════════════════════════════════════════════════════
// SESION ESPEJO (monitor inventarios)
// ═══════════════════════════════════════════════════════════════
const ESPEJO_FILE = 'espejo_inventarios.json';

function getEspejoFile() {
  const folder = DriveApp.getFolderById(HISTORIAL_FOLDER);
  const files = folder.getFilesByName(ESPEJO_FILE);
  if (files.hasNext()) return files.next();
  return folder.createFile(ESPEJO_FILE, '[]', 'application/json');
}

function actualizarEspejo(p) {
  try {
    const file = getEspejoFile();
    const data = JSON.parse(file.getBlob().getDataAsString());
    const suc = p.sucursal || '';
    const idx = data.findIndex(function(e){ return e.sucursal === suc; });
    const entry = { sucursal: suc, activo: p.activo === '1', usuario: p.usuario || '', hora: p.hora || '', ts: new Date().getTime() };
    if (idx >= 0) data[idx] = entry;
    else data.push(entry);
    file.setContent(JSON.stringify(data));
    return ok('espejo actualizado');
  } catch(e) { return error(e.toString()); }
}

function listarEspejo() {
  try {
    const file = getEspejoFile();
    const data = JSON.parse(file.getBlob().getDataAsString());
    // Auto-clear stale entries (older than 2hrs)
    const hace2hrs = new Date().getTime() - (2 * 60 * 60 * 1000);
    const activos = data.filter(function(e){ return e.ts > hace2hrs; });
    return jsonR(activos);
  } catch(e) { return jsonR([]); }
}

// ═══════════════════════════════════════════════════════════════
// COMPRAS CONCILIADAS
// ═══════════════════════════════════════════════════════════════

function formatearFecha(valor) {
  try {
    var meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
    var d;
    if (valor instanceof Date) {
      d = valor;
    } else {
      d = new Date(String(valor)); // handles ISO: 2026-04-27T22:06:35.000Z
    }
    if (isNaN(d.getTime())) return String(valor);
    // Use Sheets script timezone (America/Mexico_City)
    var tz  = Session.getScriptTimeZone();
    var fmt = Utilities.formatDate(d, tz, 'dd/MM/yyyy-HH:mm:ss');
    var parts = fmt.split('/');
    var mesIdx = parseInt(parts[1]) - 1;
    return parts[0] + '/' + meses[mesIdx] + '/' + parts[2].replace('-', '-').substring(0, parts[2].length) + 'Hrs';
  } catch(e) {
    return String(valor);
  }
}

function getConciliacion(params) {
  try {
    const ss = SpreadsheetApp.openById('1hqpPHPr9msdhqjN3WBMYGrIH9XxRVmdoayhcscipNJg');
    const tabs = ['ÁVILA CAMACHO', 'PASEO JARDINES', 'POSTRERÍA'];
    const filtroSuc = params.sucursal || '';
    const desde = params.desde ? new Date(params.desde + 'T00:00:00') : null;
    const hasta = params.hasta ? new Date(params.hasta + 'T23:59:59') : null;
    const resultado = [];

    tabs.forEach(function(tab) {
      if (filtroSuc && filtroSuc !== tab) return;
      const sheet = ss.getSheetByName(tab);
      if (!sheet || sheet.getLastRow() <= 1) return;
      const data = sheet.getDataRange().getValues();
      data.slice(1).forEach(function(row) {
        if (!row[0]) return;
        const fechaStr = String(row[1]);
        const fechaRow = new Date(fechaStr);
        if (desde && fechaRow < desde) return;
        if (hasta && fechaRow > hasta) return;
        resultado.push({
          folio:         String(row[0]),
          fecha:         formatearFecha(row[1]),
          sucursal:      String(row[2]),
          fechaEntrega:  formatearFecha(row[5]),
          totalEstimado: (String(row[7]).indexOf('$') >= 0 ? String(row[7]) : '$' + String(row[7])),
          totalReal:     String(row[13] || row[7]),
          estado:        String(row[10] || 'PENDIENTE'),
          observaciones: String(row[8] || ''),
          calificacion:  String(row[15] || ''),
          comentarios:   String(row[16] || '')
        });
      });
    });

    resultado.sort(function(a, b){ return new Date(b.fecha) - new Date(a.fecha); });
    return jsonR(resultado);
  } catch(e) {
    Logger.log('Error getConciliacion: ' + e);
    return jsonR([]);
  }
}

// ═══════════════════════════════════════════════════════════════
// CANCELAR SDP
// ═══════════════════════════════════════════════════════════════

// TRANSFERENCIAS INTERNAS
function handleTransferencia(p) {
  try {
    Logger.log('TRANSFERENCIA recibida - keys: ' + Object.keys(p).join(','));
    Logger.log('TRANSFERENCIA: folio=' + p.folio + ' emisora=' + p.emisora + ' receptora=' + p.receptora + ' total=' + p.totalEstimado);
    const ss = SpreadsheetApp.openById('1hqpPHPr9msdhqjN3WBMYGrIH9XxRVmdoayhcscipNJg');

    var montoNum = parseFloat(String(p.totalEstimado||'0').replace(/[^0-9.]/g,'')) || 0;
    var montoPos = fmtMoneyAS(montoNum);
    var montoNeg = '-' + fmtMoneyAS(montoNum);

    // Guardar en Sheets de AMBAS sucursales
    ['emisora','receptora'].forEach(function(tipo, idx) {
      var sucursal  = idx===0 ? p.emisora : p.receptora;
      var tab       = normalizarTab(sucursal);
      var sheet     = ss.getSheetByName(tab);
      if (!sheet) sheet = ss.insertSheet(tab);
      var info      = SUCURSALES_INFO[sucursal] || {};
      var montoFila = idx===0 ? montoNeg : montoPos;
      sheet.appendRow([
        p.folio, p.timestamp, sucursal, info.correo||'', info.gerente||'',
        p.timestamp, p.numProductos||0, montoFila, p.observaciones||'',
        p.productos||'', 'EVALUADA', 'TRANSFERENCIA-'+(idx===0?'EMISORA':'RECEPTORA'),
        p.timestamp, montoFila, '','',''
      ]);
    });

    // Generar HTML caratula PDF
    var filasProd = '';
    var n = 1;
    (p.productos||'').split(';').forEach(function(item) {
      var pts = item.split('|');
      if (!pts[0]) return;
      var subtotal = parseFloat(String(pts[5]||'0').replace(/[^0-9.]/g,'')) || 0;
      filasProd += '<tr style="background:' + (n%2===0?'#f5f8fc':'white') + '">' +
        '<td style="padding:5px 8px;font-size:10px;">' + n + '</td>' +
        '<td style="padding:5px 8px;font-size:10px;font-family:monospace;">' + (pts[0]||'') + '</td>' +
        '<td style="padding:5px 8px;font-size:10px;">' + (pts[1]||'') + '</td>' +
        '<td style="padding:5px 8px;font-size:10px;text-align:right;">' + fmtMoneyAS(pts[3]||0) + '</td>' +
        '<td style="padding:5px 8px;font-size:10px;text-align:center;">' + (pts[4]||'') + '</td>' +
        '<td style="padding:5px 8px;font-size:10px;text-align:right;font-weight:700;">' + fmtMoneyAS(subtotal) + '</td>' +
        '</tr>';
      n++;
    });

    var htmlPDF = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
      '<style>body{font-family:Arial,sans-serif;margin:16px;font-size:12px;}' +
      'table{width:100%;border-collapse:collapse;}' +
      'th{background:#1D4F73;color:white;padding:5px 8px;font-size:10px;text-align:left;}' +
      'td{padding:4px 8px;border-bottom:1px solid #EEF2F7;}</style></head><body>' +
      '<div style="display:flex;justify-content:space-between;border-bottom:3px solid #3C8EC8;padding-bottom:8px;margin-bottom:12px;">' +
      '<div><div style="font-size:18px;font-weight:700;color:#1D4F73;">Luna Media Café</div>' +
      '<div style="font-size:9px;color:#6B7E93;text-transform:uppercase;letter-spacing:1px;">Transferencia Interna entre Sucursales</div></div>' +
      '<div style="text-align:right;"><div style="font-size:18px;font-weight:700;color:#3C8EC8;">' + p.folio + '</div>' +
      '<div style="background:#3C8EC8;color:white;padding:2px 8px;border-radius:8px;font-size:9px;font-weight:700;display:inline-block;margin-top:2px;">TRANSFERENCIA</div>' +
      '<div style="font-size:9px;color:#999;margin-top:2px;">' + p.timestamp + '</div></div></div>' +
      '<table style="margin-bottom:12px;"><tr>' +
      '<td style="padding:5px;color:#6B7E93;font-weight:700;font-size:11px;width:25%;border:none;">EMISORA (EGRESO)</td>' +
      '<td style="padding:5px;font-weight:700;color:#D4504A;font-size:11px;border:none;">📤 ' + p.emisora + ' &nbsp; ' + montoNeg + '</td>' +
      '<td style="padding:5px;color:#6B7E93;font-weight:700;font-size:11px;width:25%;border:none;">RECEPTORA (INGRESO)</td>' +
      '<td style="padding:5px;font-weight:700;color:#2E9E6B;font-size:11px;border:none;">📥 ' + p.receptora + ' &nbsp; ' + montoPos + '</td>' +
      '</tr>' +
      (p.observaciones ? '<tr><td style="padding:5px;color:#6B7E93;font-weight:700;font-size:11px;border:none;">OBSERVACIONES</td>' +
      '<td colspan="3" style="padding:5px;font-size:11px;border:none;">' + p.observaciones + '</td></tr>' : '') +
      '</table>' +
      '<table><thead><tr><th>#</th><th>Código</th><th>Producto</th><th style="text-align:right;">Precio</th><th style="text-align:center;">Cantidad</th><th style="text-align:right;">Subtotal</th></tr></thead>' +
      '<tbody>' + filasProd + '</tbody>' +
      '<tfoot><tr><td colspan="5" style="text-align:right;font-weight:700;padding:6px 8px;background:#EBF4FB;color:#1D4F73;border:none;">TOTAL TRANSFERENCIA</td>' +
      '<td style="font-weight:700;padding:6px 8px;background:#EBF4FB;color:#1D4F73;border:none;">' + montoPos + '</td></tr></tfoot></table>' +
      '<div style="display:grid;grid-template-columns:1fr 1fr;gap:40px;margin-top:30px;">' +
      '<div style="text-align:center;"><div style="border-top:2px solid #182533;margin-bottom:6px;margin-top:36px;"></div>' +
      '<div style="font-size:11px;font-weight:600;">' + p.emisora + '</div>' +
      '<div style="font-size:9px;color:#6B7E93;">Sucursal Emisora</div></div>' +
      '<div style="text-align:center;"><div style="border-top:2px solid #182533;margin-bottom:6px;margin-top:36px;"></div>' +
      '<div style="font-size:11px;font-weight:600;">' + p.receptora + '</div>' +
      '<div style="font-size:9px;color:#6B7E93;">Sucursal Receptora</div></div>' +
      '</div></body></html>';

    // Generar PDF y guardar en Drive carpeta Pedidos/emisora
    var pdfBlob = null;
    try {
      pdfBlob = generarPDF(htmlPDF, p.folio, 'Pedidos', p.emisora);
      Logger.log('PDF transferencia generado: ' + p.folio);
    } catch(ePdf) { Logger.log('Error PDF transferencia: ' + ePdf); }

    // Correo detallado
    const SUCS_C = {
      'Avila Camacho':'lunamediacentro@gmail.com','Ávila Camacho':'lunamediacentro@gmail.com',
      'Paseo Jardines':'lunamediapaseojardines@gmail.com',
      'Postreria':'postrerialunamediacafe@gmail.com','Postrería':'postrerialunamediacafe@gmail.com'
    };
    var asunto = '🔄 Transferencia ' + p.folio + ' | ' + p.emisora + ' → ' + p.receptora;
    var cuerpo = '<div style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto;background:#F5F8FC;">' +
      '<div style="background:linear-gradient(135deg,#1D4F73,#3C8EC8);padding:24px 28px;border-radius:12px 12px 0 0;">' +
      '<p style="color:rgba(255,255,255,.7);font-size:11px;text-transform:uppercase;letter-spacing:2px;margin:0 0 4px;">Luna Media Café — Transferencia Interna</p>' +
      '<p style="color:#fff;font-size:20px;font-weight:700;margin:0;">🔄 ' + p.folio + '</p>' +
      '<p style="color:rgba(255,255,255,.75);font-size:13px;margin:0;">' + p.timestamp + '</p></div>' +
      '<div style="background:white;padding:20px 28px;">' +
      '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px;">' +
      '<tr><td style="padding:8px;color:#6B7E93;font-weight:700;width:40%">EMISORA</td>' +
      '<td style="padding:8px;font-weight:700;color:#D4504A;">📤 ' + p.emisora + ' (' + montoNeg + ')</td></tr>' +
      '<tr style="background:#f5f8fc"><td style="padding:8px;color:#6B7E93;font-weight:700">RECEPTORA</td>' +
      '<td style="padding:8px;font-weight:700;color:#2E9E6B;">📥 ' + p.receptora + ' (' + montoPos + ')</td></tr>' +
      '<tr><td style="padding:8px;color:#6B7E93;font-weight:700">TOTAL</td>' +
      '<td style="padding:8px;font-size:16px;font-weight:700;color:#1D4F73;">' + montoPos + '</td></tr>' +
      (p.observaciones?'<tr style="background:#f5f8fc"><td style="padding:8px;color:#6B7E93;font-weight:700">OBSERVACIONES</td><td style="padding:8px;">' + p.observaciones + '</td></tr>':'') +
      '</table>' +
      '<p style="font-size:11px;color:#6B7E93;font-weight:700;text-transform:uppercase;letter-spacing:1px;margin:0 0 8px;">Productos transferidos</p>' +
      '<table style="width:100%;border-collapse:collapse;">' +
      '<thead><tr style="background:#1D4F73;">' +
      '<th style="padding:7px 10px;color:white;font-size:11px;text-align:left;">Código</th>' +
      '<th style="padding:7px 10px;color:white;font-size:11px;text-align:left;">Producto</th>' +
      '<th style="padding:7px 10px;color:white;font-size:11px;text-align:center;">Cant.</th>' +
      '<th style="padding:7px 10px;color:white;font-size:11px;text-align:right;">Precio</th>' +
      '<th style="padding:7px 10px;color:white;font-size:11px;text-align:right;">Subtotal</th>' +
      '</tr></thead><tbody>' +
      (p.productos||'').split(';').map(function(item, i) {
        var pts = item.split('|');
        if (!pts[0]) return '';
        return '<tr style="background:' + (i%2===0?'white':'#f5f8fc') + '">' +
          '<td style="padding:7px 10px;font-size:12px;font-family:monospace;">' + (pts[0]||'') + '</td>' +
          '<td style="padding:7px 10px;font-size:12px;">' + (pts[1]||'') + '</td>' +
          '<td style="padding:7px 10px;font-size:12px;text-align:center;">' + (pts[4]||'') + '</td>' +
          '<td style="padding:7px 10px;font-size:12px;text-align:right;">' + fmtMoneyAS(pts[3]||0) + '</td>' +
          '<td style="padding:7px 10px;font-size:12px;text-align:right;font-weight:700;color:#1D4F73;">' + fmtMoneyAS(pts[5]||0) + '</td></tr>';
      }).join('') +
      '</tbody><tfoot><tr style="background:#EBF4FB;">' +
      '<td colspan="4" style="padding:8px 10px;font-weight:700;text-align:right;color:#1D4F73;">TOTAL</td>' +
      '<td style="padding:8px 10px;font-weight:700;color:#1D4F73;font-size:15px;">' + montoPos + '</td>' +
      '</tr></tfoot></table>' +
      '<div style="background:#E6F1FB;border-left:4px solid #1D4F73;padding:12px 16px;border-radius:0 8px 8px 0;margin-top:16px;font-size:12px;color:#0C447C;">' +
      'Registrado en Compras Conciliadas. Emisora: egreso ' + montoNeg + ' | Receptora: ingreso ' + montoPos + '</div>' +
      '</div></div>';

    // Enviar a emisora, receptora, Admin, CEO
    [SUCS_C[p.emisora], SUCS_C[p.receptora], ADMIN_EMAIL, CEO_EMAIL].filter(Boolean).forEach(function(to){
      try {
        var opts = {to:to, subject:asunto, htmlBody:cuerpo, name:'LunaMedia Ops'};
        if (pdfBlob) opts.attachments = [pdfBlob];
        MailApp.sendEmail(opts);
        Logger.log('Correo transferencia enviado a: ' + to);
      } catch(eM){ Logger.log('Error correo trans a ' + to + ': ' + eM); }
    });

    Logger.log('Transferencia registrada OK: ' + p.folio + ' | ' + montoNeg + ' / ' + montoPos);
    return ok('transferencia registrada: ' + p.folio);
  } catch(e) {
    Logger.log('Error handleTransferencia: ' + e);
    return error(e.toString());
  }
}


function cancelarSDP(p) {
  try {
    const ss    = SpreadsheetApp.openById('1hqpPHPr9msdhqjN3WBMYGrIH9XxRVmdoayhcscipNJg');
    const tab   = normalizarTab(p.sucursal);
    const sheet = ss.getSheetByName(tab);
    if (!sheet) return error('Hoja no encontrada: ' + p.sucursal);

    const data = sheet.getDataRange().getValues();
    let encontrado = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.folio)) {
        const row = i + 1;
        sheet.getRange(row, 11).setValue('CANCELADA');
        sheet.getRange(row, 13).setValue(p.timestamp);
        sheet.getRange(row, 17).setValue('CANCELADO — ' + p.motivo + (p.observaciones ? '. ' + p.observaciones : '') + ' | Por: ' + (p.canceladoPor || 'Admin'));
        encontrado = true;
        break;
      }
    }
    if (!encontrado) return error('Folio no encontrado: ' + p.folio);

    // Notify all
    const asunto = '🚫 SDP Cancelada — ' + p.folio + ' | ' + p.sucursal;
    const cuerpo = '<div style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto;background:#F5F8FC;">' +
      '<div style="background:#D4504A;padding:24px 28px;border-radius:12px 12px 0 0;">' +
      '<div style="color:white;font-size:11px;letter-spacing:2px;text-transform:uppercase;opacity:.8;margin-bottom:4px;">Luna Media Cafe — Notificacion de Cancelacion</div>' +
      '<div style="color:white;font-size:22px;font-weight:700;">SDP Cancelada</div>' +
      '<div style="color:rgba(255,255,255,.8);font-size:13px;margin-top:4px;">' + p.timestamp + '</div></div>' +
      '<div style="background:white;padding:20px 28px;">' +
      '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px;">' +
      '<tr><td style="padding:8px;color:#6B7E93;font-weight:700;width:40%">FOLIO CANCELADO</td><td style="padding:8px;font-size:18px;font-weight:700;color:#D4504A;">' + p.folio + '</td></tr>' +
      '<tr style="background:#f5f8fc"><td style="padding:8px;color:#6B7E93;font-weight:700">SUCURSAL</td><td style="padding:8px;font-weight:600">' + p.sucursal + '</td></tr>' +
      '<tr><td style="padding:8px;color:#6B7E93;font-weight:700">MOTIVO</td><td style="padding:8px;font-weight:700;color:#D4504A">' + p.motivo + '</td></tr>' +
      '<tr style="background:#f5f8fc"><td style="padding:8px;color:#6B7E93;font-weight:700">CANCELADO POR</td><td style="padding:8px">' + (p.canceladoPor || 'Admin') + '</td></tr>' +
      (p.observaciones ? '<tr><td style="padding:8px;color:#6B7E93;font-weight:700">OBSERVACIONES</td><td style="padding:8px;font-style:italic">' + p.observaciones + '</td></tr>' : '') +
      '</table>' +
      '<div style="background:#FDECEA;border-left:4px solid #D4504A;padding:12px 16px;border-radius:0 8px 8px 0;font-size:13px;color:#8B1A1A;">Esta SDP ha sido marcada como CANCELADA en el sistema. El monto estimado no se sumara a los totales del periodo.</div>' +
      '</div></div>';

    const todosCorreos = [ADMIN_EMAIL, CEO_EMAIL, CEDIS_EMAIL,
      'lunamediacentro@gmail.com', 'lunamediapaseojardines@gmail.com', 'postrerialunamediacafe@gmail.com'];
    todosCorreos.forEach(function(to) {
      try { MailApp.sendEmail({ to: to, subject: asunto, htmlBody: cuerpo, name: 'LunaMedia Ops' }); }
      catch(e) { Logger.log('Error correo cancelacion a ' + to + ': ' + e); }
    });

    Logger.log('SDP cancelada OK: ' + p.folio);
    return ok('sdp cancelada');
  } catch(e) {
    Logger.log('Error cancelarSDP: ' + e);
    return error(e.toString());
  }
}


// ═══════════════════════════════════════════════════════════════
// HISTORIAL SDP CON PDF URL
// ═══════════════════════════════════════════════════════════════
function getHistorialSDP(params) {
  try {
    if (!params) params = {};
    const ss          = SpreadsheetApp.openById('1hqpPHPr9msdhqjN3WBMYGrIH9XxRVmdoayhcscipNJg');
    const tabs        = ['ÁVILA CAMACHO', 'PASEO JARDINES', 'POSTRERÍA'];
    const filtroSuc   = (params.sucursal     || '').trim();
    const filtroEst   = (params.estadoFiltro || '').trim();
    const desdeStr    = (params.desde        || '').trim();
    const hastaStr    = (params.hasta        || '').trim();
    const resultado   = [];

    Logger.log('getHistorialSDP inicio: suc=[' + filtroSuc + '] desde=[' + desdeStr + '] hasta=[' + hastaStr + ']');

    const desdeD = desdeStr ? new Date(desdeStr + 'T00:00:00') : null;
    const hastaD = hastaStr ? new Date(hastaStr + 'T23:59:59') : null;

    tabs.forEach(function(tab) {
      if (filtroSuc && filtroSuc !== tab) return;
      var sheet;
      try { sheet = ss.getSheetByName(tab); } catch(es) { return; }
      if (!sheet || sheet.getLastRow() <= 1) { Logger.log('Sin datos: ' + tab); return; }

      const data = sheet.getDataRange().getValues();
      Logger.log(tab + ': ' + (data.length - 1) + ' filas');

      data.slice(1).forEach(function(row) {
        if (!row[0]) return;

        // Date filter — parse ISO or DD/MM/YYYY format
        if (desdeD || hastaD) {
          try {
            var fd;
            if (row[1] instanceof Date) {
              fd = row[1];
            } else {
              var s2 = String(row[1]);
              // ISO format: 2026-04-27T22:06:35.000Z
              // DD/MM/YYYY format: 27/04/2026
              if (s2.indexOf('T') >= 0 || s2.indexOf('-') >= 0) {
                fd = new Date(s2);
              } else {
                var m2 = s2.match(/(\d{2})\/(\d{2})\/(\d{4})/);
                fd = m2 ? new Date(m2[3], parseInt(m2[2])-1, parseInt(m2[1])) : new Date(s2);
              }
            }
            if (!isNaN(fd.getTime())) {
              if (desdeD && fd < desdeD) return;
              if (hastaD && fd > hastaD) return;
            }
          } catch(efd) { /* include if parse fails */ }
        }

        const estado = String(row[10] || 'PENDIENTE');
        if (filtroEst && estado !== filtroEst) return;

        resultado.push({
          folio:         String(row[0]),
          fecha:         formatearFecha(row[1]),
          sucursal:      String(row[2]  || ''),
          fechaEntrega:  formatearFecha(row[5]),
          numProductos:  String(row[6]  || ''),
          totalEstimado: String(row[7]  || '$0.00').indexOf('$') >= 0 ? String(row[7]) : '$' + String(row[7]),
          estado:        estado,
          calificacion:  String(row[15] || ''),
          pdfUrl:        '' // loaded on demand via get_pdf_url endpoint
        });
      });
    });

    Logger.log('getHistorialSDP fin: ' + resultado.length + ' registros');
    return jsonR(resultado);
  } catch(e) {
    Logger.log('Error getHistorialSDP: ' + e.toString());
    return jsonR([]);
  }
}




function getPdfUrl(params) {
  try {
    const folio    = params.folio     || '';
    const tipoPdf  = params.tipo_doc  || 'Inventarios';
    const sucursal = params.sucursal  || '';
    if (!folio) return jsonR({ url: '' });
    const sucNorm = sucursal
      .replace('ÁVILA CAMACHO','Avila Camacho')
      .replace('Ávila Camacho','Avila Camacho')
      .replace('PASEO JARDINES','Paseo Jardines')
      .replace('POSTRERÍA','Postreria')
      .replace('Postrería','Postreria') || 'General';
    const folder = getPDFSubFolder(tipoPdf, sucNorm);
    const files  = folder.getFilesByName(folio + '.pdf');
    if (files.hasNext()) {
      const f = files.next();
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return jsonR({ url: 'https://drive.google.com/file/d/' + f.getId() + '/view' });
    }
    // Search all sucursal subfolders if not found
    const root   = DriveApp.getFolderById(PDF_PARENT_FOLDER);
    const mainD  = getOrCreateFolder(root, PDF_FOLDER_NAME);
    const tipoD  = getOrCreateFolder(mainD, tipoPdf);
    const subs   = tipoD.getFolders();
    while (subs.hasNext()) {
      const sub = subs.next();
      const ff  = sub.getFilesByName(folio + '.pdf');
      if (ff.hasNext()) {
        const f = ff.next();
        f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return jsonR({ url: 'https://drive.google.com/file/d/' + f.getId() + '/view' });
      }
    }
    return jsonR({ url: '' });
  } catch(e) {
    Logger.log('getPdfUrl error: ' + e);
    return jsonR({ url: '' });
  }
}


// ═══════════════════════════════════════════════════════════════
// FORMATO MONEDA Y ELIMINAR REGISTROS
// ═══════════════════════════════════════════════════════════════
function fmtMontoGS(val) {
  var n = parseFloat(String(val).replace(/[$,]/g,'')) || 0;
  return '$' + n.toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2});
}

function eliminarSDP(p) {
  try {
    const ss    = SpreadsheetApp.openById('1hqpPHPr9msdhqjN3WBMYGrIH9XxRVmdoayhcscipNJg');
    const tab   = normalizarTab(p.sucursal);
    const sheet = ss.getSheetByName(tab);
    if (!sheet) return error('Hoja no encontrada: ' + p.sucursal);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.folio)) {
        sheet.deleteRow(i + 1);
        Logger.log('SDP eliminada: ' + p.folio);
        return ok('sdp eliminada');
      }
    }
    return error('Folio no encontrado: ' + p.folio);
  } catch(e) {
    Logger.log('Error eliminarSDP: ' + e);
    return error(e.toString());
  }
}

function eliminarInv(p) {
  try {
    const file = getHistorialFile();
    const data = JSON.parse(file.getBlob().getDataAsString());
    const filtered = data.filter(function(d){ return d.folio !== p.folio; });
    file.setContent(JSON.stringify(filtered));
    Logger.log('Inventario eliminado del historial: ' + p.folio);
    return ok('inventario eliminado');
  } catch(e) {
    Logger.log('Error eliminarInv: ' + e);
    return error(e.toString());
  }
}




// ═══════════════════════════════════════════════════════════════
// CHECK QR ENTREGA — REPARTIDOR
// ═══════════════════════════════════════════════════════════════
function registrarCheckEntrega(p) {
  try {
    const sucursal  = p.sucursal  || '';
    const check     = p.check     || ''; // 'entrada' o 'salida'
    const hora      = p.hora      || '';
    const fechaHora = p.fechaHora || '';
    const folio     = p.folio     || '';

    if (!sucursal || !check) return error('Datos incompletos');

    // Update the SDP row with the check time
    if (folio) {
      const ss    = SpreadsheetApp.openById('1hqpPHPr9msdhqjN3WBMYGrIH9XxRVmdoayhcscipNJg');
      const tab   = normalizarTab(sucursal);
      const sheet = ss.getSheetByName(tab);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][0]) === String(folio)) {
            // Col 18 = hora entrada, Col 19 = hora salida (0-indexed: 17, 18)
            // Make sure sheet has enough columns
            if (check === 'entrada') {
              sheet.getRange(i + 1, 18).setValue(hora);
              sheet.getRange(i + 1, 19).setValue(fechaHora);
            } else {
              sheet.getRange(i + 1, 20).setValue(hora);
              sheet.getRange(i + 1, 21).setValue(fechaHora);
              // Marcar EN_RUTA para que no aparezca en pedidos pendientes del QR
              if (String(data[i][10]) === 'PENDIENTE') {
                sheet.getRange(i + 1, 11).setValue('EN_RUTA');
              }
            }
            Logger.log('Check ' + check + ' registrado: ' + folio + ' | ' + hora);
            break;
          }
        }
      }
    }

    // Also save to a dedicated check log
    const folder    = DriveApp.getFolderById(HISTORIAL_FOLDER);
    const checkFile = (() => {
      const f = folder.getFilesByName('check_entregas.json');
      return f.hasNext() ? f.next() : folder.createFile('check_entregas.json', '[]', 'application/json');
    })();
    const logs = JSON.parse(checkFile.getBlob().getDataAsString());
    logs.unshift({
      sucursal:  sucursal,
      folio:     folio,
      check:     check,
      hora:      hora,
      fechaHora: fechaHora,
      ts:        new Date().getTime()
    });
    // Keep last 200 entries
    checkFile.setContent(JSON.stringify(logs.slice(0, 200)));

    return ok('check ' + check + ' registrado: ' + hora);
  } catch(e) {
    Logger.log('Error registrarCheckEntrega: ' + e);
    return error(e.toString());
  }
}


function getCheckEntrega(p) {
  try {
    const folio    = p.folio    || '';
    const sucursal = p.sucursal || '';
    if (!folio) return jsonR({});

    const folder = DriveApp.getFolderById(HISTORIAL_FOLDER);
    const files  = folder.getFilesByName('check_entregas.json');
    if (!files.hasNext()) return jsonR({entrada:'', salida:'', folio:folio});

    const logs = JSON.parse(files.next().getBlob().getDataAsString());
    Logger.log('getCheckEntrega: folio=' + folio + ' logs=' + logs.length);

    // Filter by folio (most reliable) — no date filter to avoid format issues
    const byFolio = logs.filter(function(l){ return l.folio === folio; });
    const entradas = byFolio.filter(function(l){ return l.check === 'entrada'; });
    const salidas  = byFolio.filter(function(l){ return l.check === 'salida';  });

    // Fallback: search by sucursal (last entry today)
    var entSuc = entradas, salSuc = salidas;
    if (!entradas.length || !salidas.length) {
      var bySuc = logs.filter(function(l){ return l.sucursal === sucursal; });
      if (!entradas.length) entSuc = bySuc.filter(function(l){ return l.check === 'entrada'; }).slice(0,1);
      if (!salidas.length)  salSuc = bySuc.filter(function(l){ return l.check === 'salida';  }).slice(0,1);
    }

    var resultado = {
      entrada: entSuc.length ? entSuc[0].hora : '',
      salida:  salSuc.length ? salSuc[0].hora : '',
      folio:   folio
    };
    Logger.log('getCheckEntrega resultado: ' + JSON.stringify(resultado));
    return jsonR(resultado);
  } catch(e) {
    Logger.log('getCheckEntrega error: ' + e);
    return jsonR({entrada:'', salida:'', folio:p.folio||''});
  }
}

// ═══════════════════════════════════════════════════════════════
// UTILIDADES
// ═══════════════════════════════════════════════════════════════
function jsonR(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function ok(msg) { return jsonR({ result: 'ok', msg: msg || '' }); }
function error(msg) { return jsonR({ result: 'error', msg: msg || '' }); }
function normalizarTab(sucursal) {
  const mapa = {
    'Avila Camacho':  'ÁVILA CAMACHO',
    'Ávila Camacho':  'ÁVILA CAMACHO',
    'Paseo Jardines': 'PASEO JARDINES',
    'Postreria':      'POSTRERÍA',
    'Postrería':      'POSTRERÍA'
  };
  return mapa[sucursal] || (sucursal || '').toUpperCase();
}
function fmtMoneyAS(val) {
  var n = parseFloat(String(val).replace(/[$,]/g,'')) || 0;
  return '$' + n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function testHistorialSDP() {
  var result = getHistorialSDP({});
  Logger.log('TEST RESULT: ' + result.getContent().substring(0, 500));
}

function diagnosticoSheet() {
  try {
    var ss    = SpreadsheetApp.openById('1hqpPHPr9msdhqjN3WBMYGrIH9XxRVmdoayhcscipNJg');
    var tabs  = ss.getSheets();
    var info  = [];
    tabs.forEach(function(sheet) {
      var rows = sheet.getLastRow();
      var cols = sheet.getLastColumn();
      var muestra = '';
      if (rows > 1) {
        var data = sheet.getRange(2, 1, Math.min(3, rows-1), Math.min(5, cols)).getValues();
        muestra = JSON.stringify(data);
      }
      info.push('HOJA: [' + sheet.getName() + '] filas=' + rows + ' cols=' + cols + ' muestra=' + muestra);
    });
    Logger.log(info.join('\n'));
    return info.join('\n');
  } catch(e) {
    Logger.log('ERROR diagnostico: ' + e.toString());
    return 'ERROR: ' + e.toString();
  }
}

function testConciliacion() {
  var result = getConciliacion({});
  var data = JSON.parse(result.getContent());
  Logger.log('Total registros: ' + data.length);
  data.forEach(function(r) {
    Logger.log(r.folio + ' | ' + r.sucursal + ' | ' + r.estado + ' | ' + r.totalEstimado);
  });
}

function testTransferencia() {
  var p = {
    folio: 'TRA-TEST01',
    timestamp: new Date().toLocaleString('es-MX'),
    emisora: 'Ávila Camacho',
    receptora: 'Paseo Jardines',
    totalEstimado: '$500.00',
    numProductos: 2,
    observaciones: 'Prueba',
    productos: 'COD001|Producto Test||$250.00|2x|$500.00'
  };
  var result = handleTransferencia(p);
  Logger.log('Resultado: ' + result.getContent());
}

function testTransferenciaCompleta() {
  // Simula exactamente lo que manda el portal via iframePost
  var params = {
    tipo: 'transferencia',
    folio: 'TRA-' + Date.now().toString().slice(-6),
    timestamp: new Date().toLocaleString('es-MX', {hour12:false}),
    emisora: 'Ávila Camacho',
    receptora: 'Paseo Jardines',
    totalEstimado: '$500.00',
    numProductos: '1',
    productos: 'COD001|Producto Prueba||$500.00|1x|$500.00',
    observaciones: 'Prueba desde test',
    enviadoPor: 'Admin'
  };
  
  Logger.log('Llamando handleTransferencia con: ' + JSON.stringify(params));
  var result = handleTransferencia(params);
  Logger.log('Resultado: ' + result.getContent());
}
