/**
 * Carteras en vivo — Google Apps Script
 *
 * - actualizarPrecios(): cada 15 min, lunes a viernes 10:30–18:15 (hora Argentina), trae precios de Rava.
 * - revisarAlertas(): todos los días a las 9, registra en la pestaña Alertas los cheques que se acreditan ese día.
 * - generarPremarket() cada 30 min entre 8:00 y 11:30, generarPostmarket() cada 30 min entre 18:00 y 20:00 (lunes a viernes;
 *   fuera de esas franjas no hace nada) y generarSemanal() lunes 8:00:
 *   arman los informes con datos públicos (Rava, TradingView, DolarApi, ArgentinaDatos), diarios y portales (Ámbito, Infobae,
 *   El Cronista, La Nación, CNBC, MarketWatch, Yahoo Finance, Google Noticias) y, si están configurados, X y la portada con Claude.
 * - doGet(): sirve la página web. Administración: Portfolio, Alertas, Informes y Noticias.
 *   Clientes: su cartera y Noticias. Acceso con link personal + clave (pestaña Accesos).
 *
 * No envía mails. Instalación: menú Carteras → "Instalar automatizaciones" (una sola vez).
 */

const TZ = 'America/Argentina/Buenos_Aires';
const RAVA_URL = 'https://mercado.rava.com/api/prices/arg';
const CLASES_PRECIO = ['Acciones', 'CEDEARs', 'Bonos', 'Letras', 'ONs'];
const CLASES_BONO = ['Bonos', 'Letras', 'ONs'];
const RAVA_USD = { BPOC7: 'BPC7D', BPOD7: 'BPD7D' };   // BOPREAL en USD
const DIAS = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
const CFG = { mepVivo: 'B2', mepUsado: 'B4', actualizado: 'B5', tenencias: 'B7', alertasActivas: 'B8', diasProx: 'B9', estado: 'B10' };
const HANDLERS = ['actualizarPrecios', 'revisarAlertas', 'avisoCheques', 'generarPremarket', 'generarPostmarket', 'generarSemanal'];

// ================================================================ menú e instalación
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Carteras')
    .addItem('Actualizar precios ahora', 'actualizarPreciosAhora')
    .addItem('Revisar alertas ahora', 'revisarAlertasAhora')
    .addSeparator()
    .addItem('Generar informe premarket', 'premarketAhora')
    .addItem('Generar informe postmarket', 'postmarketAhora')
    .addItem('Generar informe semanal', 'semanalAhora')
    .addItem('Actualizar noticias', 'actualizarNoticiasAhora')
    .addSeparator()
    .addItem('Accesos de clientes', 'mostrarAccesos')
    .addSeparator()
    .addItem('Instalar automatizaciones', 'instalar')
    .addToUi();
}

// Versión de los disparadores: al cambiar horarios en el código, subir este número. El disparador de precios
// (que corre cada 15 minutos) detecta la versión nueva y reinstala todo solo, sin pasar por el menú.
const VERSION_DISPARADORES = 3;

function instalar() {
  const ui = SpreadsheetApp.getUi();   // primero: si se llama desde la página, falla antes de tocar nada
  crearDisparadores_();
  accesos_(SpreadsheetApp.getActive());
  const r = [actualizarPrecios_(true), revisarAlertas_(), actualizarNoticias_(true), generarInforme_('premarket', true), generarInforme_('postmarket', true), generarInforme_('semanal', true)];
  ui.alert('Automatizaciones instaladas:\n· Precios cada 15 minutos en horario de mercado\n· Alertas de cheques 9:00 h\n· Informe premarket cada 30 min de 8:00 a 11:30 h y postmarket cada 30 min de 18:00 a 20:00 h (lunes a viernes)\n· Informe semanal los lunes 8:00 h\n· Noticias cada hora\n· Links y claves de clientes: pestaña Accesos\n\n' + r.join('\n'));
}
function crearDisparadores_() {
  ScriptApp.getProjectTriggers().forEach(t => { if (HANDLERS.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('actualizarPrecios').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('revisarAlertas').timeBased().atHour(9).nearMinute(0).everyDays(1).inTimezone(TZ).create();
  ScriptApp.newTrigger('generarPremarket').timeBased().everyMinutes(30).create();    // solo actúa de 8:00 a 11:30
  ScriptApp.newTrigger('generarPostmarket').timeBased().everyMinutes(30).create();   // solo actúa de 18:00 a 20:00
  ScriptApp.newTrigger('generarSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).nearMinute(0).inTimezone(TZ).create();
  hoja_(SpreadsheetApp.getActive(), 'Agenda', ['Fecha', 'Hora', 'Evento', 'Importancia (Alta/Media)']);   // eventos locales del calendario
  PropertiesService.getScriptProperties().setProperty('VERSION_DISPARADORES', String(VERSION_DISPARADORES));
}
// La llama el disparador de precios: si el código trae disparadores nuevos, los reinstala.
function sincronizarDisparadores_() {
  const P = PropertiesService.getScriptProperties();
  if (+P.getProperty('VERSION_DISPARADORES') === VERSION_DISPARADORES) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    if (+P.getProperty('VERSION_DISPARADORES') !== VERSION_DISPARADORES) {
      crearDisparadores_();
      // Con la versión nueva publicada, regenera en el momento el informe que esté en su franja horaria.
      ['premarket', 'postmarket'].forEach(t => { if (enVentana_(t)) generarInforme_(t, false); });
    }
  }
  catch (e) { console.warn('No se pudieron actualizar los disparadores: ' + e.message); }
  finally { lock.releaseLock(); }
}

function actualizarPreciosAhora() { SpreadsheetApp.getUi().alert(actualizarPrecios_(true)); }
function revisarAlertasAhora() { SpreadsheetApp.getUi().alert(revisarAlertas_()); }
function premarketAhora() { SpreadsheetApp.getUi().alert(generarInforme_('premarket', true)); }
function postmarketAhora() { SpreadsheetApp.getUi().alert(generarInforme_('postmarket', true)); }
function semanalAhora() { SpreadsheetApp.getUi().alert(generarInforme_('semanal', true)); }

function cfg_(ss) { return ss.getSheetByName('Config'); }
function estado_(ss, msg) { cfg_(ss).getRange(CFG.estado).setValue(Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm') + ' · ' + msg); }

// ================================================================ precios
function esDisparador_(e) { return !!(e && e.triggerUid); }   // las llamadas desde la página no traen triggerUid
function actualizarPrecios(e) {   // la llama el disparador
  if (!esDisparador_(e)) return;
  sincronizarDisparadores_();
  try { accesos_(SpreadsheetApp.getActive()); } catch (x) {}
  try { actualizarNoticias_(false); } catch (x) {}
  actualizarPrecios_(false);
}

function enHorario_() {
  const now = new Date();
  const dow = +Utilities.formatDate(now, TZ, 'u');            // 1 = lunes … 7 = domingo
  const min = +Utilities.formatDate(now, TZ, 'H') * 60 + +Utilities.formatDate(now, TZ, 'm');
  return dow <= 5 && min >= 10 * 60 + 25 && min <= 18 * 60 + 15;
}

function actualizarPrecios_(forzar) {
  if (!forzar && !enHorario_()) return 'Fuera de horario de mercado.';
  const ss = SpreadsheetApp.getActive();
  try {
    agregarPosicionesNuevas_(ss);
    const want = especiesAPedir_(ss);
    const res = UrlFetchApp.fetch(RAVA_URL, { muteHttpExceptions: true, headers: { Accept: 'application/json' } });
    if (res.getResponseCode() !== 200) throw new Error('Rava respondió HTTP ' + res.getResponseCode());
    const datos = JSON.parse(res.getContentText()).datos || [];

    const best = {};
    datos.forEach(x => {
      if (x.plazo !== 1 && x.plazo !== 2) return;
      const ult = parseFloat(x.ultimo) || 0, ant = parseFloat(x.anterior) || 0;
      const k = x.simbolo + '|' + x.moneda;
      const score = (ult > 0 ? 2 : 0) + (x.plazo === 2 ? 1 : 0);   // 24 hs con precio primero
      if (!best[k] || score > best[k].score) best[k] = { score, ult, ant, hora: x.hora, plazo: x.plazo, sym: x.simbolo };
    });

    const filas = [], faltan = [];
    Object.keys(want).sort().forEach(t => {
      const w = want[t], b = best[w.sym + '|' + w.mon];
      if (!b || b.ult <= 0) { faltan.push(t); return; }
      if (w.ref && !(b.ult / w.ref > 0.4 && b.ult / w.ref < 2.5)) { faltan.push(t); return; }   // precio absurdo
      let chg = b.ant > 0 ? b.ult / b.ant - 1 : 0;
      if (Math.abs(chg) >= 0.5) chg = 0;
      filas.push([t, b.ult, b.ant, chg, b.hora, b.plazo, b.sym]);
    });
    if (filas.length < 10) throw new Error('Rava devolvió muy pocos precios (' + filas.length + '). No se tocó nada.');

    const P = ss.getSheetByName('Precios');
    const last = P.getLastRow();
    if (last > 1) P.getRange(2, 1, last - 1, 7).clearContent();
    P.getRange(2, 1, filas.length, 7).setValues(filas);

    const a = best['AL30|ARS'], d = best['AL30D|USD'];
    if (a && d && a.ult > 0 && d.ult > 0) cfg_(ss).getRange(CFG.mepVivo).setValue(Math.round(a.ult / d.ult * 100) / 100);
    cfg_(ss).getRange(CFG.actualizado).setValue(new Date());
    const msg = 'Precios OK: ' + filas.length + ' especies' + (faltan.length ? '. Sin precio: ' + faltan.join(', ') : '') + '.';
    estado_(ss, msg);
    return msg;
  } catch (e) {
    const msg = 'Error al actualizar precios: ' + e.message;
    estado_(ss, msg);
    return msg;
  }
}

function especiesAPedir_(ss) {
  const X = ss.getSheetByName('Posiciones');
  const v = X.getRange(2, 1, Math.max(X.getLastRow() - 1, 1), 12).getValues();
  const want = {};
  v.forEach(r => {
    const t = String(r[2]).trim(), clase = r[4], mon = r[5];
    if (!t || !CLASES_PRECIO.includes(clase)) return;
    const usd = mon === 'USD' && CLASES_BONO.includes(clase);
    if (!want[t]) want[t] = { sym: usd ? (RAVA_USD[t] || t + 'D') : t, mon: usd ? 'USD' : 'ARS', ref: Number(r[11]) || null };
  });
  return want;
}

// Si en Movimientos hay una compra de una especie que la cuenta no tenía, le agrega la fila en Posiciones.
function agregarPosicionesNuevas_(ss) {
  const X = ss.getSheetByName('Posiciones'), M = ss.getSheetByName('Movimientos');
  const pos = X.getRange(2, 1, Math.max(X.getLastRow() - 1, 1), 12).getValues();
  const tiene = new Set(pos.filter(r => r[2]).map(r => r[0] + '|' + r[2]));
  const modelo = {};
  pos.forEach(r => { if (r[2] && !modelo[r[2]]) modelo[r[2]] = r; });
  const movs = M.getRange(2, 1, Math.max(M.getLastRow() - 1, 1), 6).getValues();
  let fila = pos.findIndex(r => !r[2]); fila = fila < 0 ? pos.length + 2 : fila + 2;
  movs.forEach(m => {
    const cta = String(m[1]).trim(), t = String(m[2]).trim().toUpperCase();
    if (!cta || !t || m[3] !== 'Compra' || tiene.has(cta + '|' + t)) return;
    const mod = modelo[t];
    const cliente = (pos.find(r => String(r[0]) === cta) || [])[1] || '';
    const f9 = X.getRange(2, 9).getFormulaR1C1(), f10 = X.getRange(2, 10).getFormulaR1C1();
    X.getRange(fila, 1, 1, 8).setValues([[cta, cliente, t, mod ? mod[3] : '(completar)', mod ? mod[4] : 'Acciones', mod ? mod[5] : 'ARS', mod ? mod[6] : 1, 0]]);
    X.getRange(fila, 9, 1, 2).setFormulasR1C1([[f9, f10]]);
    X.getRange(fila, 12).setValue(mod ? mod[11] : m[5]);
    tiene.add(cta + '|' + t); fila++;
  });
}

// ================================================================ alertas
function revisarAlertas(e) { if (esDisparador_(e)) revisarAlertas_(); }

function revisarAlertas_() {
  const ss = SpreadsheetApp.getActive();
  if (String(cfg_(ss).getRange(CFG.alertasActivas).getValue()).trim().toUpperCase() === 'NO') return 'Alertas pausadas en Config.';
  const A = ss.getSheetByName('Alertas');
  const claves = new Set(A.getLastRow() > 1 ? A.getRange(2, 7, A.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : []);
  const hoyK = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const clave = 'CHQ-' + hoyK;
  if (claves.has(clave)) return 'La alerta de cheques de hoy ya estaba registrada.';

  const ch = chequesRaw_(ss).filter(c => c.acred && Utilities.formatDate(c.acred, TZ, 'yyyy-MM-dd') === hoyK);
  if (!ch.length) return 'Hoy no se acredita ningún cheque.';
  const total = ch.reduce((s, c) => s + c.nominal, 0);
  const dmy = d => Utilities.formatDate(d, TZ, 'dd/MM/yyyy');
  let det = ch.map(c => '- ' + c.cliente + ' (Cta. ' + c.cuenta + ') - Titular: ' + c.librador + (c.numero ? ' - N° de cheque: ' + c.numero : '') +
    (c.vto ? ' - Vencimiento: ' + dmy(c.vto) : '') + ' - Monto nominal: ' + pesos_(c.nominal)).join('\n');
  det += '\n\nTotal a acreditar hoy (' + ch.length + (ch.length === 1 ? ' cheque' : ' cheques') + '): ' + pesos_(total) +
    '\nLos fondos quedan disponibles en la cuenta comitente para su posterior aplicación. Fechas según la liquidación informada por Nasini; montos nominales.';
  const ctas = [...new Set(ch.map(c => c.cuenta))].join(', ');
  A.insertRowAfter(1);
  A.getRange(2, 1, 1, 7).setValues([[new Date(), 'Cheques', 'Acreditación de cheques – ' + dmy(new Date()), det, ctas, total, clave]]);
  return 'Alerta registrada: ' + ch.length + ' cheques por ' + pesos_(total) + '.';
}

function chequesRaw_(ss) {
  const Q = ss.getSheetByName('Cheques');
  const n = Q.getLastRow() - 1;
  if (n < 1) return [];
  return Q.getRange(2, 1, n, 11).getValues().filter(r => r[2]).map(r => ({
    cuenta: String(r[0]), cliente: r[1], codigo: r[2], librador: r[3], numero: String(r[4] || ''),
    vto: r[5] instanceof Date ? r[5] : null, acred: r[6] instanceof Date ? r[6] : null,
    nominal: Number(r[7]) || 0, estado: r[9], valor: Number(r[10]) || 0,
  }));
}

function pesos_(x) {
  const s = (Math.round((Number(x) || 0) * 100) / 100).toFixed(2).split('.');
  return '$' + s[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + s[1];
}

// ================================================================ informes (premarket, postmarket, semanal)
// Se arman con datos públicos. La portada (titular, bajada, claves) puede redactarla Claude si existe la propiedad
// ANTHROPIC_API_KEY; los posteos de X requieren X_BEARER_TOKEN. Sin esas propiedades todo sigue funcionando. Se guardan como JSON en la pestaña "Informes"
// (premarket y postmarket se sobrescriben; los semanales se acumulan) y la página los muestra.

const TV_SCAN = 'https://scanner.tradingview.com/global/scan';
const MACRO = [
  ['SP:SPX', 'S&P 500', 'eeuu'], ['NASDAQ:IXIC', 'Nasdaq Composite', 'eeuu'], ['DJ:DJI', 'Dow Jones', 'eeuu'],
  ['CME_MINI:ES1!', 'Futuro S&P 500', 'futuros'], ['CME_MINI:NQ1!', 'Futuro Nasdaq 100', 'futuros'], ['CBOT_MINI:YM1!', 'Futuro Dow Jones', 'futuros'],
  ['CBOE:VIX', 'VIX (volatilidad)', 'riesgo'], ['TVC:US10Y', 'Bono del Tesoro EE.UU. 10 años (%)', 'riesgo'], ['TVC:US02Y', 'Bono del Tesoro EE.UU. 2 años (%)', 'riesgo'],
  ['TVC:DXY', 'Índice dólar (DXY)', 'riesgo'], ['NYMEX:CL1!', 'Petróleo WTI', 'commodities'], ['ICEEUR:BRN1!', 'Petróleo Brent', 'commodities'],
  ['COMEX:GC1!', 'Oro', 'commodities'], ['BITSTAMP:BTCUSD', 'Bitcoin', 'commodities'],
  ['XETR:DAX', 'DAX (Alemania)', 'global'], ['TVC:SX5E', 'Euro Stoxx 50', 'global'], ['TVC:UKX', 'FTSE 100 (Reino Unido)', 'global'],
  ['TVC:NI225', 'Nikkei 225 (Japón)', 'global'], ['TVC:HSI', 'Hang Seng (Hong Kong)', 'global'],
];
const ADRS = ['GGAL', 'YPF', 'PAM', 'BMA', 'SUPV', 'TGS', 'CEPU', 'IRS', 'LOMA', 'BBAR', 'TEO', 'EDN', 'CRESY', 'VIST', 'MELI', 'GLOB'];
// Diarios y portales: [fuente, url, sección]. Una fuente caída no frena el informe.
const RSS = [
  ['Ámbito', 'https://www.ambito.com/rss/pages/finanzas.xml', 'argentina'],
  ['Ámbito', 'https://www.ambito.com/rss/pages/economia.xml', 'argentina'],
  ['Ámbito', 'https://www.ambito.com/rss/pages/negocios.xml', 'argentina'],
  ['Infobae', 'https://www.infobae.com/arc/outboundfeeds/rss/category/economia/', 'argentina'],
  ['El Cronista', 'https://www.cronista.com/arc/outboundfeeds/rss/?outputType=xml', 'argentina'],
  ['La Nación', 'https://www.lanacion.com.ar/arc/outboundfeeds/rss/category/economia/?outputType=xml', 'argentina'],
  ['CNBC', 'https://search.cnbc.com/rs/search/combinedcms/view.xml?partnerId=wrss01&id=20910258', 'mundo'],
  ['MarketWatch', 'https://feeds.content.dowjones.io/public/rss/mw_topstories', 'mundo'],
  ['Yahoo Finance', 'https://finance.yahoo.com/news/rssindex', 'mundo'],
];
const PALABRAS = ['merval', 'riesgo país', 'riesgo pais', 'bonos', 'dólar', 'dolar', 'bcra', 'wall street', 'reserva federal', 'fed ', 'acciones',
  'adr', 'inflación', 'licitación', 'tesoro', 'cedear', 'tasas', 'caución', 'mercados', 'byma', 'cnv', 'fmi', 'deuda',
  'stocks', 'treasury', 'yields', 's&p', 'nasdaq', 'dow ', 'inflation', 'earnings', 'oil', 'dollar', 'rates', 'powell', 'tariff', 'jobs report', 'argentina'];
// Cuentas de X que se siguen si no se configura la propiedad X_CUENTAS (ver README).
const X_CUENTAS = ['FinanzasArgy', 'BCRAoficial', 'LuisCaputoAR', 'Ambitocom', 'cronistacom', 'infobae', 'WSJmarkets', 'federalreserve', 'DeItaone'];
// Canales públicos de Telegram que se leen sin clave (vista web t.me/s/...). Se cambian con la propiedad TELEGRAM_CANALES.
const TELEGRAM_CANALES = ['finanzasargy'];
const IA_MODELO = 'claude-opus-5-5';
const EMPRESAS = { GGAL: 'galicia', YPFD: 'ypf', PAMP: 'pampa', BMA: 'macro', SUPV: 'supervielle', TGSU2: 'transportadora de gas del sur',
  TGNO4: 'gas del norte', METR: 'metrogas', ECOG: 'ecogas', LOMA: 'loma negra', IRSA: 'irsa', ALUA: 'aluar', TXAR: 'ternium', BBAR: 'bbva',
  BYMA: 'byma', COME: 'comercial del plata', CEPU: 'central puerto', AGRO: 'agrometal', INVJ: 'juramento', MELI: 'mercadolibre', NVDA: 'nvidia',
  MU: 'micron', TSM: 'tsmc', GOOGL: 'google', META: 'meta', NFLX: 'netflix', NKE: 'nike', MCD: "mcdonald", UBER: 'uber', IBM: 'ibm', AAL: 'american airlines',
  JD: 'jd.com', NU: 'nubank', SNDK: 'sandisk', WMT: 'walmart', MSFT: 'microsoft' };

// Franjas de actualización automática (hora Argentina, lunes a viernes). Fuera de ellas el informe queda congelado
// hasta el día siguiente; desde el menú de la planilla se puede regenerar a mano en cualquier momento.
const VENTANA = { premarket: [8 * 60, 11 * 60 + 30], postmarket: [18 * 60, 20 * 60] };
function enVentana_(tipo) {
  const now = new Date(), min = +Utilities.formatDate(now, TZ, 'H') * 60 + +Utilities.formatDate(now, TZ, 'm');
  return diaHabil_() && min >= VENTANA[tipo][0] && min <= VENTANA[tipo][1];
}
function generarPremarket(e) { if (esDisparador_(e) && enVentana_('premarket')) generarInforme_('premarket', false); }
function generarPostmarket(e) { if (esDisparador_(e) && enVentana_('postmarket')) generarInforme_('postmarket', false); }
function generarSemanal(e) { if (esDisparador_(e)) generarInforme_('semanal', false); }
function diaHabil_() { return +Utilities.formatDate(new Date(), TZ, 'u') <= 5; }

function generarInforme_(tipo, manual) {
  const ss = SpreadsheetApp.getActive();
  try {
    if (tipo === 'postmarket') { actualizarPrecios_(true); SpreadsheetApp.flush(); }
    const inf = tipo === 'premarket' ? armarPremarket_(ss) : tipo === 'postmarket' ? armarPostmarket_(ss) : armarSemanal_(ss);
    if (tipo !== 'semanal') inf.editorial = editorialIA_(tipo, datosEditorial_(inf));
    inf.fuentes = fuentes_(inf);
    guardarInforme_(ss, tipo, recortar_(inf));
    if (tipo === 'postmarket') guardarHistorial_(ss, inf);
    const msg = 'Informe ' + tipo + ' generado (' + Utilities.formatDate(new Date(), TZ, 'dd/MM HH:mm') + ').';
    estado_(ss, msg);
    return msg;
  } catch (e) {
    const msg = 'Error al generar el informe ' + tipo + ': ' + e.message;
    estado_(ss, msg);
    return msg;
  }
}

// ---------------------------------------------------------------- almacenamiento
function hoja_(ss, nombre, encabezado) {
  let sh = ss.getSheetByName(nombre);
  if (!sh) { sh = ss.insertSheet(nombre); sh.getRange(1, 1, 1, encabezado.length).setValues([encabezado]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function guardarInforme_(ss, tipo, inf) {
  const sh = hoja_(ss, 'Informes', ['Tipo', 'Fecha', 'Generado', 'Contenido (JSON)']);
  const json = JSON.stringify(inf);
  if (json.length > 49000) throw new Error('el informe supera el tamaño de una celda');
  const fila = [tipo, inf.fecha, new Date(), json];
  const n = sh.getLastRow() - 1;
  if (tipo !== 'semanal' && n > 0) {
    const tipos = sh.getRange(2, 1, n, 1).getValues().map(r => r[0]);
    const i = tipos.indexOf(tipo);
    if (i >= 0) { sh.getRange(i + 2, 1, 1, 4).setValues([fila]); return; }   // se sobrescribe
  }
  if (tipo === 'semanal' && n > 0) {   // si se regenera la misma semana, reemplaza
    const v = sh.getRange(2, 1, n, 2).getValues();
    const i = v.findIndex(r => r[0] === 'semanal' && r[1] === inf.fecha);
    if (i >= 0) { sh.getRange(i + 2, 1, 1, 4).setValues([fila]); return; }
  }
  sh.insertRowAfter(1);
  sh.getRange(2, 1, 1, 4).setValues([fila]);
  const sem = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map((r, i) => r[0] === 'semanal' ? i + 2 : 0).filter(Boolean);
  sem.slice(26).reverse().forEach(r => sh.deleteRow(r));   // conserva 26 semanas
}
function guardarHistorial_(ss, inf) {
  const sh = hoja_(ss, 'Historial', ['Fecha', 'Contenido (JSON)']);
  const X = ss.getSheetByName('Posiciones');
  const v = X.getRange(2, 1, Math.max(X.getLastRow() - 1, 1), 14).getValues();
  const precios = {};
  v.forEach(r => { if (r[2] && typeof r[13] === 'number') precios[r[2]] = r[13]; });
  const snap = { fecha: inf.fecha, precios, mep: cfg_(ss).getRange(CFG.mepUsado).getValue(),
    merval: inf.rueda && inf.rueda.merval, riesgo: inf.rueda && inf.rueda.riesgo, dolares: inf.rueda && inf.rueda.dolares,
    spx: inf.rueda && inf.rueda.spx, cuentas: inf.cartera.cuentas.map(c => ({ cuenta: c.cuenta, cartera: c.cartera, capital: c.capital })) };
  const n = sh.getLastRow() - 1;
  const fechas = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(r => r[0]) : [];
  const i = fechas.indexOf(inf.fecha);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(JSON.stringify(snap));
  else { sh.insertRowAfter(1); sh.getRange(2, 1, 1, 2).setValues([[inf.fecha, JSON.stringify(snap)]]); }
  if (sh.getLastRow() > 400) sh.deleteRows(401, sh.getLastRow() - 400);
}
function historial_(ss) {
  const sh = ss.getSheetByName('Historial');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().map(r => { try { return JSON.parse(r[1]); } catch (e) { return null; } }).filter(Boolean);
}

// ---------------------------------------------------------------- fuentes
function fetchJson_(url, opts) {
  const r = UrlFetchApp.fetch(url, Object.assign({ muteHttpExceptions: true, headers: { Accept: 'application/json' } }, opts || {}));
  if (r.getResponseCode() !== 200) throw new Error(url.split('/')[2] + ' respondió HTTP ' + r.getResponseCode());
  return JSON.parse(r.getContentText());
}
function tvScan_(tickers, cols) {
  try {
    const d = fetchJson_(TV_SCAN, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ symbols: { tickers }, columns: cols }) });
    const o = {};
    (d.data || []).forEach(x => { const v = {}; cols.forEach((c, i) => v[c] = x.d[i]); o[x.s] = v; });
    return o;
  } catch (e) { return {}; }
}
// Acciones de EE.UU. por ticker (prueba NASDAQ, NYSE y AMEX y se queda con la que exista).
function tvEEUU_(nombres) {
  const cols = ['close', 'change', 'premarket_change', 'premarket_close', 'postmarket_change', 'postmarket_close', 'earnings_release_next_date', 'description'];
  const raw = tvScan_([].concat(...nombres.map(n => ['NASDAQ:' + n, 'NYSE:' + n, 'AMEX:' + n])), cols);
  const o = {};
  Object.keys(raw).forEach(s => { const n = s.split(':')[1]; if (!o[n]) o[n] = Object.assign({ simbolo: s }, raw[s]); });
  return o;
}
function macro_() {
  let m = tvScan_(MACRO.map(x => x[0]), ['close', 'change', 'Perf.W', 'Perf.1M']);
  if (!Object.keys(m).length) m = tvScan_(MACRO.map(x => x[0]), ['close', 'change']);   // por si TradingView rechaza las columnas de rendimiento
  const n_ = v => typeof v === 'number' && isFinite(v) ? v : null;
  return MACRO.map(([s, n, g]) => ({ nombre: n, grupo: g, valor: m[s] ? m[s].close : null, var: m[s] ? m[s].change : null,
    sem: m[s] ? n_(m[s]['Perf.W']) : null, mes: m[s] ? n_(m[s]['Perf.1M']) : null }));
}
function dolares_() {
  try {
    const d = fetchJson_('https://dolarapi.com/v1/dolares');
    const nom = { oficial: 'Oficial', blue: 'Blue', bolsa: 'MEP', contadoconliqui: 'Contado con liquidación', mayorista: 'Mayorista', cripto: 'Cripto' };
    return d.filter(x => nom[x.casa]).map(x => ({ nombre: nom[x.casa], compra: x.compra, venta: x.venta, hora: x.fechaActualizacion }));
  } catch (e) { return []; }
}
function riesgoPais_() {
  try { const d = fetchJson_('https://api.argentinadatos.com/v1/finanzas/indices/riesgo-pais/ultimo'); return { valor: d.valor, fecha: d.fecha }; }
  catch (e) { return null; }
}
function ravaArg_() {
  const datos = fetchJson_(RAVA_URL).datos || [];
  const best = {};
  datos.forEach(x => {
    if (x.plazo !== 1 && x.plazo !== 2) return;
    const ult = parseFloat(x.ultimo) || 0, ant = parseFloat(x.anterior) || 0;
    const k = x.simbolo + '|' + x.moneda, score = (ult > 0 ? 2 : 0) + (x.plazo === 2 ? 1 : 0);
    if (!best[k] || score > best[k].score) best[k] = { score, ult, ant, var: ant > 0 && ult > 0 ? (ult / ant - 1) * 100 : null };
  });
  const mv = datos.find(x => x.simbolo === 'SPMERVAL' || x.especie === 'SPMERVAL');
  const merval = mv ? { valor: parseFloat(mv.ultimo) || null, anterior: parseFloat(mv.anterior) || null } : null;
  if (merval && merval.valor && merval.anterior && merval.valor !== merval.anterior) merval.var = (merval.valor / merval.anterior - 1) * 100;
  const mervalOk = merval && merval.var != null;
  return { best, merval: mervalOk ? merval : null };
}
function ravaBonos_() {
  try {
    const o = {};
    (fetchJson_('https://mercado.rava.com/api/prices/bonos').datos || []).forEach(x => {
      o[x.especie] = { tir: parseFloat(x.tir) * 100, dm: parseFloat(x.dm), paridad: parseFloat(x.paridad) * 100, vto: x.vencimiento ? x.vencimiento.slice(0, 10) : '' };
    });
    return o;
  } catch (e) { return {}; }
}
// Texto plano a partir del HTML/entidades que traen los RSS y X.
function texto_(s) {
  return String(s || '').replace(/<!\[CDATA\[|\]\]>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&#0?39;|&apos;/g, "'").replace(/&nbsp;/g, ' ').replace(/&#(\d+);/g, (m, n) => String.fromCharCode(+n)).replace(/&amp;/g, '&');
}
function plano_(html) { return texto_(texto_(html).replace(/<[^>]+>/g, ' ')).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim(); }
function corto_(s, n) { if (s.length <= n) return s; const c = s.slice(0, n); return c.slice(0, Math.max(c.lastIndexOf(' '), n - 30)).replace(/[,;:.\s]+$/, '') + '…'; }

// Ítems de un RSS/Atom: título, link, fecha, resumen e imagen.
function rssItems_(xml) {
  return (xml.match(/<(item|entry)[\s>][\s\S]*?<\/(item|entry)>/g) || []).map(it => {
    const g = tag => { const m = it.match(new RegExp('<' + tag + '(?:\\s[^>]*)?>([\\s\\S]*?)</' + tag + '>')); return m ? m[1].trim() : ''; };
    const at = re => { const m = it.match(re); return m ? texto_(m[1]) : ''; };
    const cuerpo = texto_(g('description') || g('content:encoded') || g('summary') || g('content'));
    const im = cuerpo.match(/<img[^>]+src=["']([^"']+)["']/i);
    let link = plano_(g('link')) || at(/<link[^>]*href="([^"]+)"/);
    const imagen = at(/<media:content[^>]*url="([^"]+)"[^>]*(?:medium="image"|type="image)/) || at(/<media:content[^>]*url="([^"]+\.(?:jpe?g|png|webp)[^"]*)"/i)
      || at(/<media:thumbnail[^>]*url="([^"]+)"/) || at(/<enclosure[^>]*url="([^"]+)"[^>]*type="image/) || (im ? im[1] : '');
    return { titulo: plano_(g('title')), link, fecha: Date.parse(plano_(g('pubDate') || g('published') || g('updated') || g('dc:date'))),
      resumen: plano_(cuerpo), imagen: /^https:\/\//.test(imagen) ? imagen : '', fuente: plano_(g('source')) };
  });
}

// Titulares de diarios argentinos, portales internacionales y Google Noticias (empresas en cartera),
// puntuados por palabras clave y tenencias. Devuelve como máximo `max`, repartidos entre secciones.
function noticias_(horas, max, claves) {
  const desde = Date.now() - horas * 3600000, vistos = new Set(), out = [];
  const fuentes = RSS.slice();
  if (claves.length) {
    const q = '(' + claves.slice(0, 10).map(c => '"' + c + '"').join(' OR ') + ') when:' + Math.max(1, Math.ceil(horas / 24)) + 'd';
    fuentes.push(['Google Noticias', 'https://news.google.com/rss/search?q=' + encodeURIComponent(q) + '&hl=es-419&gl=AR&ceid=AR:es-419', 'cartera']);
  }
  let resp = [];
  try { resp = UrlFetchApp.fetchAll(fuentes.map(f => ({ url: f[1], muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } }))); } catch (e) { resp = []; }
  fuentes.forEach(([fuente, url, seccion], i) => {
    try {
      const r = resp[i];
      if (!r || r.getResponseCode() !== 200) return;
      rssItems_(r.getContentText()).forEach(x => {
        let titulo = x.titulo, medio = fuente;
        if (seccion === 'cartera') { const m = titulo.match(/^(.*)\s+-\s+([^-]{2,60})$/); if (m) { titulo = m[1]; medio = x.fuente || m[2]; } }
        if (!titulo || !/^https:\/\//.test(x.link) || !(x.fecha >= desde)) return;
        const t = titulo.toLowerCase(), key = t.replace(/[^a-z0-9áéíóúñ]/g, '').slice(0, 60);
        if (vistos.has(key)) return;
        let score = seccion === 'cartera' ? 3 : 0;
        claves.forEach(c => { if (t.includes(c)) score += 3; });
        PALABRAS.forEach(p => { if (t.includes(p)) score += 1; });
        if (!score) return;
        vistos.add(key);
        let resumen = x.resumen;
        if (resumen.toLowerCase().startsWith(t.slice(0, 40)) || resumen.length < 25) resumen = '';   // Google Noticias repite el título
        out.push({ titulo, link: x.link, fuente: medio, seccion, fecha: x.fecha, resumen: corto_(resumen, 260), imagen: x.imagen, score });
      });
    } catch (e) { /* una fuente caída no frena el informe */ }
  });
  out.sort((a, b) => b.score - a.score || b.fecha - a.fecha);
  const tope = Math.ceil(max / 2), cuenta = {}, sel = [];
  out.forEach(x => { if (sel.length < max && (cuenta[x.seccion] = (cuenta[x.seccion] || 0) + 1) <= tope) sel.push(x); });
  return sel.map(({ score, ...x }) => x);
}

// Posteos recientes de X de las cuentas seguidas. Requiere la propiedad del script X_BEARER_TOKEN
// (API de X). Sin token devuelve null y la página muestra cómo activarlo.
function posteosX_(claves) {
  const P = PropertiesService.getScriptProperties(), token = P.getProperty('X_BEARER_TOKEN');
  if (!token) return null;
  const cuentas = (P.getProperty('X_CUENTAS') || X_CUENTAS.join(',')).split(',').map(c => c.trim().replace(/^@/, '')).filter(Boolean).slice(0, 20);
  const q = '(' + cuentas.map(c => 'from:' + c).join(' OR ') + ') -is:retweet -is:reply';
  try {
    const d = fetchJson_('https://api.x.com/2/tweets/search/recent?max_results=50&query=' + encodeURIComponent(q) +
      '&tweet.fields=created_at,public_metrics,author_id&expansions=author_id,attachments.media_keys&user.fields=name,username,profile_image_url&media.fields=url,preview_image_url,type',
      { headers: { Authorization: 'Bearer ' + token } });
    const U = {}, M = {};
    ((d.includes && d.includes.users) || []).forEach(u => U[u.id] = u);
    ((d.includes && d.includes.media) || []).forEach(m => M[m.media_key] = m.url || m.preview_image_url || '');
    const hace = Date.now() - 24 * 3600000;
    return (d.data || []).map(t => {
      const u = U[t.author_id] || {}, pm = t.public_metrics || {}, txt = texto_(t.text).replace(/\s*https:\/\/t\.co\/\w+$/, '');
      const low = txt.toLowerCase(), fecha = Date.parse(t.created_at);
      let score = (pm.like_count || 0) + 2 * (pm.retweet_count || 0);
      claves.forEach(c => { if (low.includes(c)) score *= 1.5; });
      const media = ((t.attachments && t.attachments.media_keys) || []).map(k => M[k]).find(x => /^https:\/\//.test(x || '')) || '';
      return { texto: corto_(txt, 400), autor: u.name || '', usuario: u.username || '', avatar: /^https:\/\//.test(u.profile_image_url || '') ? u.profile_image_url : '',
        fecha, likes: pm.like_count || 0, rts: pm.retweet_count || 0, imagen: media, link: 'https://x.com/' + (u.username || 'i') + '/status/' + t.id, score: fecha >= hace ? score : score / 4 };
    }).sort((a, b) => b.score - a.score).slice(0, 8).map(({ score, ...x }) => x);
  } catch (e) { return []; }
}

// Últimos mensajes de canales públicos de Telegram (vista web, sin clave).
function telegram_(horas) {
  const P = PropertiesService.getScriptProperties();
  const canales = (P.getProperty('TELEGRAM_CANALES') || TELEGRAM_CANALES.join(',')).split(',').map(c => c.trim().replace(/^@/, '')).filter(Boolean).slice(0, 6);
  const desde = Date.now() - horas * 3600000, out = [];
  let resp = [];
  try { resp = UrlFetchApp.fetchAll(canales.map(c => ({ url: 'https://t.me/s/' + encodeURIComponent(c), muteHttpExceptions: true, followRedirects: false, headers: { 'User-Agent': 'Mozilla/5.0' } }))); } catch (e) { resp = []; }
  canales.forEach((c, i) => {
    try {
      const r = resp[i];
      if (!r || r.getResponseCode() !== 200) return;
      const html = r.getContentText();
      const nombre = plano_((html.match(/<meta property="og:title" content="([^"]*)"/) || [])[1] || c);
      html.split('tgme_widget_message_wrap').slice(1).forEach(b => {
        const post = (b.match(/data-post="([\w]+\/\d+)"/) || [])[1];
        const txt = (b.match(/<div class="tgme_widget_message_text[^"]*"[^>]*>([\s\S]*?)<\/div>/) || [])[1];
        const fecha = Date.parse((b.match(/<time[^>]*datetime="([^"]+)"/) || [])[1] || '');
        if (!post || !txt || !(fecha >= desde)) return;
        const imagen = (b.match(/tgme_widget_message_photo_wrap[^>]*background-image:url\('([^']+)'\)/) || [])[1] || '';
        out.push({ texto: corto_(plano_(txt.replace(/<br\s*\/?>/gi, '\n')), 400), autor: nombre, usuario: c, avatar: '', fecha, likes: null, rts: null,
          imagen: /^https:\/\//.test(imagen) ? imagen : '', link: 'https://t.me/' + post, red: 'Telegram' });
      });
    } catch (e) { /* un canal caído no frena el informe */ }
  });
  return out.sort((a, b) => b.fecha - a.fecha).slice(0, 6);
}
// Redes: X (si hay token) + Telegram. Devuelve null solo si no hay nada configurado ni publicado.
function redes_(claves, horas) {
  const x = posteosX_(claves), tg = telegram_(horas);
  if (x == null && !tg.length) return null;
  return (x || []).map(p => Object.assign({ red: 'X' }, p)).concat(tg).slice(0, 12);
}

// Titular, bajada, claves y lectura del día redactados por Claude a partir de los datos del informe.
// Opcional: requiere la propiedad del script ANTHROPIC_API_KEY. Ante cualquier error devuelve null
// y la página arma la portada con reglas fijas.
function editorialIA_(tipo, datos) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) return null;
  const momento = tipo === 'premarket' ? 'antes de la apertura de la rueda (qué mirar hoy)' : 'al cierre de la rueda (qué dejó el día y qué queda para mañana)';
  const sistema = 'Sos el editor de un informe de mercado diario, estilo diario financiero, para un asesor que administra carteras en Argentina. ' +
    'Escribís ' + momento + '. Usá exclusivamente los datos y titulares que se te pasan en JSON: no inventes cifras, hechos ni citas, ' +
    'y si un dato falta no lo menciones. No des recomendaciones de compra o venta. Español rioplatense formal, frases cortas, cifras con coma decimal. ' +
    'titular: hasta 90 caracteres, informativo, sin signos de exclamación. bajada: una o dos oraciones que amplíen el titular con cifras. ' +
    'claves: entre 4 y 6 puntos breves (hasta 160 caracteres cada uno) que crucen mercados, noticias y la cartera. ' +
    'lectura: un párrafo de 70 a 120 palabras con la interpretación del día.';
  const schema = { type: 'object', additionalProperties: false, required: ['titular', 'bajada', 'claves', 'lectura'],
    properties: { titular: { type: 'string' }, bajada: { type: 'string' }, claves: { type: 'array', items: { type: 'string' } }, lectura: { type: 'string' } } };
  try {
    const r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { 'x-api-key': key, 'anthropic-version': '2026-10-21', 'anthropic-beta': 'server-side-fallback-2026-07-01' },
      payload: JSON.stringify({ model: IA_MODELO, max_tokens: 4000, fallbacks: 'default', system: sistema,
        output_config: { effort: 'low', format: { type: 'json_schema', schema } },
        messages: [{ role: 'user', content: JSON.stringify(datos) }] }),
    });
    if (r.getResponseCode() !== 200) { console.warn('Claude respondió HTTP ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 300)); return null; }
    const d = JSON.parse(r.getContentText());
    if (d.stop_reason !== 'end_turn') { console.warn('Editorial no generada: stop_reason ' + d.stop_reason); return null; }
    const e = JSON.parse((d.content || []).filter(c => c.type === 'text').map(c => c.text).join(''));
    return { titular: String(e.titular).slice(0, 140), bajada: String(e.bajada).slice(0, 400), claves: (e.claves || []).slice(0, 6).map(x => String(x).slice(0, 220)),
      lectura: String(e.lectura).slice(0, 1200), ia: true };
  } catch (e) { console.warn('Editorial no generada: ' + e.message); return null; }
}
// Lo que se le pasa a Claude: cifras redondeadas y titulares, sin datos personales de los clientes.
function datosEditorial_(inf) {
  const r2 = v => typeof v === 'number' && isFinite(v) ? Math.round(v * 100) / 100 : null;
  const pct = inf.cartera && inf.cartera.total ? r2(inf.cartera.total.varPct) : null;
  return {
    informe: inf.tipo, fecha: inf.fecha,
    mercados: (inf.macro || []).map(m => [m.nombre, r2(m.valor), r2(m.var)]),
    argentina: inf.tipo === 'premarket' ? { riesgoPais: inf.argentina.riesgo && inf.argentina.riesgo.valor, merval: inf.argentina.merval, dolares: (inf.argentina.dolares || []).map(d => [d.nombre, d.venta]) }
      : { riesgoPais: inf.rueda.riesgo && inf.rueda.riesgo.valor, riesgoPaisAnterior: inf.rueda.riesgoAnterior, merval: inf.rueda.merval, dolares: (inf.rueda.dolares || []).map(d => [d.nombre, d.venta]) },
    adrs: (inf.adrs || []).map(a => [a.ticker, r2(a.varCierre), r2(a.pre)]),
    cartera: inf.tipo === 'premarket' ? { impactoCedearsPct: inf.impacto && r2(inf.impacto.pct), mayoresCedears: (inf.cedears || []).slice(0, 5).map(c => [c.ticker, r2(c.pre)]) }
      : { variacionDiaPct: pct, suben: (inf.suben || []).slice(0, 3).map(x => [x.ticker, r2(x.var)]), bajan: (inf.bajan || []).slice(0, 3).map(x => [x.ticker, r2(x.var)]),
        afterHours: (inf.afterHours || []).slice(0, 5).map(a => [a.ticker, r2(a.post)]) },
    balancesProximos: ((inf.agenda && inf.agenda.balances) || []).map(b => [b.ticker, b.fecha]),
    calendarioEconomico: (inf.calendario || []).map(c => [c.dia, c.hora, c.pais, c.evento, c.impacto, c.estimado, c.anterior]),
    titulares: (inf.noticias || []).map(n => ({ seccion: n.seccion, fuente: n.fuente, titulo: n.titulo, resumen: n.resumen })),
    x: (inf.x || []).map(p => ({ cuenta: p.usuario, texto: p.texto })),
  };
}
// Riesgo país, Merval y MEP de las últimas ruedas registradas (más antigua primero), para los gráficos de tendencia.
function serie_(hist, hoy) {
  const s = hist.filter(h => !hoy || h.fecha < hoy.fecha).slice(0, 20).reverse()
    .map(h => ({ fecha: h.fecha, riesgo: h.riesgo ? h.riesgo.valor : null, merval: h.merval ? h.merval.valor : null, mep: h.mep || null }));
  if (hoy) s.push(hoy);
  return s;
}
// Si el informe no entra en una celda, recorta lo menos importante.
function recortar_(inf) {
  const largo = () => JSON.stringify(inf).length;
  while (largo() > 48000 && inf.noticias && inf.noticias.length > 6) inf.noticias.pop();
  if (largo() > 48000 && inf.noticias) inf.noticias.forEach(n => { n.resumen = corto_(n.resumen || '', 120); n.imagen = ''; });
  while (largo() > 48000 && inf.x && inf.x.length > 3) inf.x.pop();
  return inf;
}

// Calendario económico: Forex Factory (semana en curso; EE.UU., zona euro, China, Reino Unido y Japón, impacto alto y medio)
// más los eventos locales que se carguen en la pestaña "Agenda" (licitaciones, INDEC, BCRA, etc.).
const FF_URL = 'https://nfs.faireconomy.media/ff_calendar_thisweek.json';
const PAISES = { USD: 'EE.UU.', EUR: 'Zona euro', CNY: 'China', GBP: 'Reino Unido', JPY: 'Japón' };
function calendario_(ss, desde, hasta) {
  const cache = CacheService.getScriptCache();
  let ev = null;
  try { ev = JSON.parse(cache.get('ff_semana') || 'null'); } catch (e) { ev = null; }
  if (!ev) {
    try {
      const r = UrlFetchApp.fetch(FF_URL, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
      if (r.getResponseCode() === 200) {
        ev = JSON.parse(r.getContentText()).filter(e => PAISES[e.country] && ['High', 'Medium', 'Holiday'].includes(e.impact))
          .map(e => ({ evento: String(e.title || ''), pais: e.country, f: Date.parse(e.date), impacto: e.impact, estimado: e.forecast || '', anterior: e.previous || '' }))
          .filter(e => e.evento && isFinite(e.f));
        try { cache.put('ff_semana', JSON.stringify(ev), 3600); } catch (e) { /* sin cache */ }   // el sitio pide no consultarlo seguido
      }
    } catch (e) { ev = null; }
  }
  const k = d => Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  const out = (ev || []).filter(e => k(new Date(e.f)) >= desde && k(new Date(e.f)) <= hasta && (e.impacto !== 'Holiday' || e.pais === 'USD'))
    .map(e => ({ dia: k(new Date(e.f)), hora: e.impacto === 'Holiday' ? '' : Utilities.formatDate(new Date(e.f), TZ, 'HH:mm'), pais: PAISES[e.pais],
      evento: e.impacto === 'Holiday' ? 'Feriado: ' + e.evento : e.evento, impacto: e.impacto === 'High' ? 'alta' : e.impacto === 'Medium' ? 'media' : 'feriado',
      estimado: e.estimado, anterior: e.anterior }));
  const A = ss.getSheetByName('Agenda');
  if (A && A.getLastRow() > 1) {
    A.getRange(2, 1, A.getLastRow() - 1, 4).getValues().forEach(r => {
      if (!(r[0] instanceof Date) || !r[2]) return;
      const dia = k(r[0]); if (dia < desde || dia > hasta) return;
      const hora = r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'HH:mm') : String(r[1] || '').trim();
      out.push({ dia, hora, pais: 'Argentina', evento: String(r[2]), impacto: /alta/i.test(r[3]) ? 'alta' : 'media', estimado: '', anterior: '' });
    });
  }
  return out.sort((a, b) => (a.dia + (a.hora || '00:00')) < (b.dia + (b.hora || '00:00')) ? -1 : 1).slice(0, 40);
}
// Próximo día hábil (lunes a viernes), como clave yyyy-MM-dd.
function proximoHabil_() { const d = +Utilities.formatDate(new Date(), TZ, 'u'); return hoyK_(d === 5 ? 3 : d === 6 ? 2 : 1); }

// ---------------------------------------------------------------- datos propios
function cartera_(ss) {
  const R = ss.getSheetByName('Resumen').getRange('A5:I11').getValues();
  const num = v => typeof v === 'number' && isFinite(v) ? v : 0;
  const cuentas = R.slice(0, 6).map(r => { const m = String(r[0]).match(/^(.*?)\s*\((\d+)\)$/);
    return { cuenta: m ? m[2] : '', nombre: m ? m[1] : r[0], cartera: num(r[1]), varHoy: num(r[2]), varPct: num(r[3]) * 100, pnl: num(r[4]), pnlPct: num(r[5]) * 100, efectivo: num(r[6]), capital: num(r[7]) }; });
  const t = R[6];
  return { cuentas, total: { cartera: num(t[1]), varHoy: num(t[2]), varPct: num(t[3]) * 100, pnl: num(t[4]), pnlPct: num(t[5]) * 100, efectivo: num(t[6]), capital: num(t[7]) } };
}
function posiciones_(ss) {
  const X = ss.getSheetByName('Posiciones');
  return X.getRange(2, 1, Math.max(X.getLastRow() - 1, 1), 18).getValues().filter(r => r[2] && Number(r[9])).map(r => ({
    cuenta: String(r[0]), ticker: String(r[2]), nombre: r[3], clase: r[4], moneda: r[5], div: Number(r[6]) || 1, cantidad: Number(r[9]),
    precio: Number(r[13]) || null, precioNasini: Number(r[11]) || null, fuente: r[14], var: typeof r[15] === 'number' ? r[15] * 100 : null,
    valor: Number(r[16]) || 0, varHoy: Number(r[17]) || 0 }));
}
function porEspecie_(pos) {
  const m = {};
  pos.forEach(p => { const x = m[p.ticker] = m[p.ticker] || { ticker: p.ticker, nombre: p.nombre, clase: p.clase, moneda: p.moneda, precio: p.precio, var: p.var, fuente: p.fuente, valor: 0, varHoy: 0, cuentas: [] };
    x.valor += p.valor; x.varHoy += p.varHoy; if (!x.cuentas.includes(p.cuenta)) x.cuentas.push(p.cuenta); });
  return Object.values(m);
}
function claves_(pos) {
  const c = new Set();
  pos.forEach(p => { if (EMPRESAS[p.ticker]) c.add(EMPRESAS[p.ticker]); });
  return [...c];
}
function aforos_(ss, pos) {
  const K = ss.getSheetByName('Cuentas').getRange(2, 1, 6, 11).getValues();
  return ['11287', '11587'].map(cta => {
    const k = K.find(r => String(r[0]) === cta); if (!k) return null;
    const bonos = pos.filter(p => p.cuenta === cta && CLASES_BONO.includes(p.clase)).reduce((s, p) => s + p.valor, 0);
    const cau = Number(k[4]) || 0;
    return { cuenta: cta, caucion: cau, bonos, pct: bonos ? cau / bonos * 100 : null };
  }).filter(Boolean);
}
function chequesEntre_(ss, d0, d1) {
  const k = d => Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
  return chequesRaw_(ss).filter(c => c.acred && k(c.acred) >= d0 && k(c.acred) <= d1)
    .map(c => ({ fecha: k(c.acred), cuenta: c.cuenta, librador: c.librador, numero: c.numero, nominal: c.nominal }))
    .sort((a, b) => a.fecha < b.fecha ? -1 : 1);
}
function balances_(us, dias) {
  const lim = Date.now() + dias * 86400000, out = [];
  Object.keys(us).forEach(n => { const e = us[n].earnings_release_next_date; if (e && e * 1000 >= Date.now() - 43200000 && e * 1000 <= lim) out.push({ ticker: n, empresa: us[n].description, fecha: Utilities.formatDate(new Date(e * 1000), TZ, 'yyyy-MM-dd') }); });
  return out.sort((a, b) => a.fecha < b.fecha ? -1 : 1);
}
function hoyK_(offsetDias) { return Utilities.formatDate(new Date(Date.now() + (offsetDias || 0) * 86400000), TZ, 'yyyy-MM-dd'); }
function cedearsUS_(pos) { return [...new Set(pos.filter(p => p.clase === 'CEDEARs').map(p => p.ticker))]; }
function fuentes_(inf) {
  const f = ['BYMA vía Rava Bursátil', 'TradingView', 'DolarApi', 'ArgentinaDatos'];
  const medios = [...new Set(((inf && inf.noticias) || []).map(n => n.fuente))];
  f.push(...(medios.length ? medios : ['Ámbito', 'Infobae', 'El Cronista']));
  [...new Set(((inf && inf.x) || []).map(p => p.red === 'Telegram' ? 'Telegram (' + p.autor + ')' : 'X'))].forEach(r => f.push(r));
  if (inf && inf.editorial && inf.editorial.ia) f.push('Portada redactada con Claude (Anthropic) sobre los datos del informe');
  f.push('Resumen de Posiciones Nasini');
  return f;
}

// ---------------------------------------------------------------- premarket
function armarPremarket_(ss) {
  const pos = posiciones_(ss), esp = porEspecie_(pos), ced = cedearsUS_(pos);
  const us = tvEEUU_([...new Set(ced.concat(ADRS))]);
  const rava = ravaArg_();
  const cedears = esp.filter(e => e.clase === 'CEDEARs').map(e => { const u = us[e.ticker] || {};
    const pre = typeof u.premarket_change === 'number' ? u.premarket_change : null;
    return { ticker: e.ticker, empresa: u.description || e.nombre, valor: e.valor, cierre: u.close, varCierre: u.change, pre, impacto: pre != null ? e.valor * pre / 100 : null, cuentas: e.cuentas }; })
    .sort((a, b) => Math.abs(b.impacto || 0) - Math.abs(a.impacto || 0));
  const impactoTotal = cedears.reduce((s, c) => s + (c.impacto || 0), 0);
  const expCed = cedears.reduce((s, c) => s + c.valor, 0);
  const adrs = ADRS.filter(t => us[t]).map(t => ({ ticker: t, empresa: us[t].description, cierre: us[t].close, varCierre: us[t].change,
    pre: typeof us[t].premarket_change === 'number' ? us[t].premarket_change : null, local: (rava.best[(t === 'YPF' ? 'YPFD' : t === 'PAM' ? 'PAMP' : t === 'TGS' ? 'TGSU2' : t === 'IRS' ? 'IRSA' : t === 'TEO' ? 'TECO2' : t === 'CRESY' ? 'CRES' : t) + '|ARS'] || {}).var }));
  const car = cartera_(ss), claves = claves_(pos);
  return {
    tipo: 'premarket', fecha: hoyK_(), generado: Date.now(),
    macro: macro_(), adrs, cedears: cedears.slice(0, 25), impacto: { total: impactoTotal, exposicion: expCed, pct: expCed ? impactoTotal / expCed * 100 : null, cartera: car.total.cartera },
    argentina: { merval: rava.merval, riesgo: riesgoPais_(), dolares: dolares_(), mep: cfg_(ss).getRange(CFG.mepUsado).getValue() },
    calendario: calendario_(ss, hoyK_(), hoyK_()),
    agenda: { cheques: chequesEntre_(ss, hoyK_(), hoyK_()), balances: balances_(us, 7) },
    noticias: noticias_(20, 16, claves), x: redes_(claves, 20), serie: serie_(historial_(ss)),
  };
}

// ---------------------------------------------------------------- postmarket
function armarPostmarket_(ss) {
  const pos = posiciones_(ss), esp = porEspecie_(pos), ced = cedearsUS_(pos);
  const us = tvEEUU_([...new Set(ced.concat(ADRS))]);
  const rava = ravaArg_(), bonosRef = ravaBonos_(), mac = macro_(), car = cartera_(ss);
  const m = n => mac.find(x => x.nombre === n) || {};
  const conVar = esp.filter(e => e.fuente === 'Rava' && e.var != null && CLASES_PRECIO.includes(e.clase));
  const contrib = conVar.map(e => ({ ticker: e.ticker, nombre: e.nombre, clase: e.clase, var: e.var, aporte: e.varHoy, valor: e.valor, cuentas: e.cuentas }));
  const suben = contrib.filter(x => x.aporte > 0).sort((a, b) => b.aporte - a.aporte).slice(0, 6);
  const bajan = contrib.filter(x => x.aporte < 0).sort((a, b) => a.aporte - b.aporte).slice(0, 6);
  const rf = esp.filter(e => CLASES_BONO.includes(e.clase)).map(e => { const sym = e.moneda === 'USD' ? (RAVA_USD[e.ticker] || e.ticker + 'D') : e.ticker;
    const b = bonosRef[sym] || bonosRef[e.ticker] || {};
    return { ticker: e.ticker, nombre: e.nombre, moneda: e.moneda, precio: e.precio, var: e.var, tir: isFinite(b.tir) ? b.tir : null, dm: isFinite(b.dm) ? b.dm : null,
      paridad: isFinite(b.paridad) ? b.paridad : null, vto: b.vto || '', valor: e.valor, cuentas: e.cuentas }; })
    .sort((a, b) => b.valor - a.valor);
  const after = [...new Set(ced.concat(ADRS))].filter(t => us[t] && typeof us[t].postmarket_change === 'number' && Math.abs(us[t].postmarket_change) >= 0.01)
    .map(t => ({ ticker: t, empresa: us[t].description, cierre: us[t].close, varCierre: us[t].change, post: us[t].postmarket_change, cartera: ced.includes(t) }))
    .sort((a, b) => Math.abs(b.post) - Math.abs(a.post)).slice(0, 10);
  const dol = dolares_(), rp = riesgoPais_();
  const hist = historial_(ss), prev = hist.find(h => h.fecha < hoyK_()), claves = claves_(pos);
  const mepHoy = (dol.find(d => d.nombre === 'MEP') || {}).venta || cfg_(ss).getRange(CFG.mepUsado).getValue();
  const adrs = ADRS.filter(t => us[t]).map(t => ({ ticker: t, empresa: us[t].description, cierre: us[t].close, varCierre: us[t].change,
    post: typeof us[t].postmarket_change === 'number' ? us[t].postmarket_change : null }));
  return {
    tipo: 'postmarket', fecha: hoyK_(), generado: Date.now(),
    rueda: { merval: rava.merval, riesgo: rp, riesgoAnterior: prev && prev.riesgo ? prev.riesgo.valor : null, dolares: dol,
      spx: { valor: m('S&P 500').valor, var: m('S&P 500').var }, ndx: { valor: m('Nasdaq Composite').valor, var: m('Nasdaq Composite').var },
      dji: { valor: m('Dow Jones').valor, var: m('Dow Jones').var }, vix: { valor: m('VIX (volatilidad)').valor, var: m('VIX (volatilidad)').var },
      us10: { valor: m('Bono del Tesoro EE.UU. 10 años (%)').valor, var: m('Bono del Tesoro EE.UU. 10 años (%)').var } },
    macro: mac, cartera: car, suben, bajan, adrs, rentaFija: rf, aforo: aforos_(ss, pos), afterHours: after,
    serie: serie_(hist, { fecha: hoyK_(), riesgo: rp ? rp.valor : null, merval: rava.merval ? rava.merval.valor : null, mep: mepHoy }),
    calendario: calendario_(ss, proximoHabil_(), proximoHabil_()), calendarioDia: proximoHabil_(),
    agenda: { cheques: chequesEntre_(ss, hoyK_(1), hoyK_(4)), balances: balances_(us, 7) },
    noticias: noticias_(12, 16, claves), x: redes_(claves, 12),
  };
}

// ---------------------------------------------------------------- semanal
function armarSemanal_(ss) {
  const pos = posiciones_(ss), hist = historial_(ss);   // historial: más reciente primero
  const mep = Number(cfg_(ss).getRange(CFG.mepUsado).getValue()) || 1;
  const fin = hist[0] || null;
  const corte = fin ? fin.fecha : hoyK_();
  const limite = Utilities.formatDate(new Date(Date.parse(corte + 'T12:00:00-03:00') - 6 * 86400000), TZ, 'yyyy-MM-dd');
  const ini = hist.find(h => h.fecha <= limite) || hist[hist.length - 1] || null;
  const usaNasini = !ini || ini === fin;
  const base = t => usaNasini ? null : ini.precios[t];
  const ult = t => fin && fin.precios[t] != null ? fin.precios[t] : null;
  const filas = pos.filter(p => CLASES_PRECIO.includes(p.clase)).map(p => {
    const p1 = ult(p.ticker) != null ? ult(p.ticker) : p.precio, p0 = base(p.ticker) != null ? base(p.ticker) : p.precioNasini;
    const k = p.cantidad / p.div * (p.moneda === 'USD' ? mep : 1);
    return Object.assign({}, p, { p0, p1, aporte: p0 && p1 ? (p1 - p0) * k : 0, v0: p0 ? p0 * k : 0, v1: p1 ? p1 * k : 0, varSem: p0 && p1 ? (p1 / p0 - 1) * 100 : null });
  });
  const cuentas = ['11200', '11710', '11750', '11054', '11287', '11587'].map(cta => {
    const f = filas.filter(x => x.cuenta === cta), v0 = f.reduce((s, x) => s + x.v0, 0), ap = f.reduce((s, x) => s + x.aporte, 0);
    const c = cartera_(ss).cuentas.find(x => x.cuenta === cta) || {};
    return { cuenta: cta, nombre: c.nombre, aporte: ap, pct: v0 ? ap / v0 * 100 : null, capital: c.capital, cartera: c.cartera };
  });
  const esp = {};
  filas.forEach(x => { const e = esp[x.ticker] = esp[x.ticker] || { ticker: x.ticker, nombre: x.nombre, clase: x.clase, varSem: x.varSem, aporte: 0, cuentas: [] };
    e.aporte += x.aporte; if (!e.cuentas.includes(x.cuenta)) e.cuentas.push(x.cuenta); });
  const lista = Object.values(esp);
  const totAp = cuentas.reduce((s, c) => s + c.aporte, 0), totV0 = filas.reduce((s, x) => s + x.v0, 0);
  const ced = cedearsUS_(pos), us = tvEEUU_([...new Set(ced.concat(ADRS))]);
  const serie = hist.filter(h => h.fecha >= (ini ? ini.fecha : '0') && h.fecha <= corte).reverse()
    .map(h => ({ fecha: h.fecha, riesgo: h.riesgo ? h.riesgo.valor : null, merval: h.merval ? h.merval.valor : null, mep: h.mep }));
  return {
    tipo: 'semanal', fecha: hoyK_(), generado: Date.now(),
    periodo: { desde: usaNasini ? 'Resumen Nasini ' + cfg_(ss).getRange(CFG.tenencias).getDisplayValue() : ini.fecha, hasta: corte, baseNasini: usaNasini },
    total: { aporte: totAp, pct: totV0 ? totAp / totV0 * 100 : null, aporteUSD: totAp / mep, capital: cartera_(ss).total.capital, mep },
    cuentas, suman: lista.filter(x => x.aporte > 0).sort((a, b) => b.aporte - a.aporte).slice(0, 5), restan: lista.filter(x => x.aporte < 0).sort((a, b) => a.aporte - b.aporte).slice(0, 5),
    mercado: { inicio: ini && !usaNasini ? { riesgo: ini.riesgo, merval: ini.merval, mep: ini.mep, spx: ini.spx } : null, fin: fin ? { riesgo: fin.riesgo, merval: fin.merval, mep: fin.mep, spx: fin.spx } : null, serie },
    macro: macro_(),
    agenda: { cheques: chequesEntre_(ss, hoyK_(), hoyK_(6)), balances: balances_(us, 7) },
    noticias: noticias_(72, 8, claves_(pos)),
  };
}

function informesParaWeb_(ss) {
  const sh = ss.getSheetByName('Informes');
  const out = { premarket: null, postmarket: null, semanales: [] };
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(r => {
    let j = null; try { j = JSON.parse(r[3]); } catch (e) { return; }
    if (r[0] === 'premarket' && !out.premarket) out.premarket = j;
    else if (r[0] === 'postmarket' && !out.postmarket) out.postmarket = j;
    else if (r[0] === 'semanal' && out.semanales.length < 12) out.semanales.push(j);
  });
  return out;
}

// ================================================================ accesos (portal de clientes)
// La app web corre con la cuenta del dueño y es accesible sin cuenta de Google. Por eso NINGÚN dato sale
// sin un link personal (token) y su clave, y las funciones de menú/disparadores rechazan llamadas desde la página.
// Pestaña "Accesos": una fila por cliente (sus cuentas) y una de administración (todas las cuentas).

const ACCESOS_HDR = ['Nombre', 'Rol', 'Cuentas', 'Token', 'Clave', 'Activo', 'Último ingreso', 'Link personal'];
const SESION_SEG = 21600;          // 6 h
const MAX_INTENTOS = 5;

function nuevoToken_() { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 40); }
function nuevaClave_() { return String(100000 + Math.floor(Math.random() * 900000)); }
function urlApp_() { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } }

function accesos_(ss) {
  const sh = hoja_(ss, 'Accesos', ACCESOS_HDR);
  if (sh.getLastRow() < 2) {   // primera vez: un acceso por cliente (según la pestaña Cuentas) y uno de administración
    const K = ss.getSheetByName('Cuentas');
    const ctas = K.getRange(2, 1, Math.max(K.getLastRow() - 1, 1), 2).getValues().filter(r => /^\d+$/.test(String(r[0])));
    const porCliente = {};
    ctas.forEach(r => { (porCliente[r[1]] = porCliente[r[1]] || []).push(String(r[0])); });
    const filas = [['Administración', 'admin', 'todas', nuevoToken_(), nuevaClave_(), 'SI', '', '']]
      .concat(Object.keys(porCliente).map(n => [n, 'cliente', porCliente[n].join(', '), nuevoToken_(), nuevaClave_(), 'SI', '', '']));
    sh.getRange(2, 1, filas.length, ACCESOS_HDR.length).setValues(filas);
  }
  const url = urlApp_();
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, ACCESOS_HDR.length).getValues();
  v.forEach((r, i) => {   // completa token/clave/link si alguien agregó un cliente a mano
    let cambio = false;
    if (r[0] && !r[3]) { r[3] = nuevoToken_(); cambio = true; }
    if (r[0] && !String(r[4]).trim()) { r[4] = nuevaClave_(); cambio = true; }
    if (r[0] && url && r[7] !== url + '?c=' + r[3]) { r[7] = url + '?c=' + r[3]; cambio = true; }
    if (cambio) sh.getRange(i + 2, 1, 1, ACCESOS_HDR.length).setValues([r]);
  });
  return v.map((r, i) => ({ r, fila: i + 2 })).filter(x => x.r[0]).map(({ r, fila }) => ({ fila, nombre: String(r[0]), rol: String(r[1]).trim().toLowerCase() === 'admin' ? 'admin' : 'cliente',
    cuentas: String(r[2]).split(/[,;\s]+/).filter(x => /^\d+$/.test(x)), token: String(r[3]), clave: String(r[4]).trim(),
    activo: String(r[5]).trim().toUpperCase() !== 'NO' }));
}

// Llamada desde la página: valida link + clave y devuelve una sesión temporal.
function login(token, clave) {
  token = String(token || '').trim(); clave = String(clave || '').trim();
  if (!/^[a-f0-9]{20,64}$/i.test(token)) throw new Error('Enlace de acceso inválido.');
  const cache = CacheService.getScriptCache(), kf = 'f_' + token;
  const fallos = Number(cache.get(kf) || 0);
  if (fallos >= MAX_INTENTOS) throw new Error('Demasiados intentos fallidos. Vuelva a intentar en 15 minutos.');
  const ss = SpreadsheetApp.getActive();
  const a = accesos_(ss).find(x => x.token === token && x.activo);
  if (!a || !clave || clave !== a.clave) {
    cache.put(kf, String(fallos + 1), 900);
    throw new Error(a ? 'Clave incorrecta.' : 'Enlace de acceso inválido o desactivado.');
  }
  cache.remove(kf);
  const sid = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  cache.put('s_' + sid, token, SESION_SEG);
  ss.getSheetByName('Accesos').getRange(a.fila, 7).setValue(new Date());
  return { sid, nombre: a.nombre, rol: a.rol };
}

// Llamada desde la página: datos según la sesión. "ver" = nombre de cliente (solo administración, vista previa).
function datos(sid, ver) {
  const cache = CacheService.getScriptCache();
  const token = sid && /^[a-f0-9]{40,80}$/i.test(String(sid)) ? cache.get('s_' + sid) : null;
  if (!token) throw new Error('SESION');
  const ss = SpreadsheetApp.getActive();
  const lista = accesos_(ss);
  const a = lista.find(x => x.token === token && x.activo);
  if (!a) { cache.remove('s_' + sid); throw new Error('SESION'); }
  cache.put('s_' + sid, token, SESION_SEG);   // renueva la sesión mientras se usa
  try { actualizarNoticias_(false); } catch (e) { /* sin noticias nuevas */ }
  const full = datosCompletos_(ss);
  full.noticias = noticiasGuardadas_(ss);
  full.mercado = mercadoPublico_(full.informes);
  if (a.rol === 'admin' && !ver) {
    full.historial = historialCuentas_(ss, null);
    full.acceso = { rol: 'admin', nombre: a.nombre, clientes: lista.filter(x => x.rol === 'cliente').map(x => x.nombre) };
    return full;
  }
  const destino = a.rol === 'admin' ? lista.find(x => x.rol === 'cliente' && x.nombre === ver) : a;
  if (!destino) throw new Error('Cliente inexistente.');
  return filtrarCliente_(ss, full, destino, a.rol === 'admin');
}

function filtrarCliente_(ss, full, c, preview) {
  const ok = new Set(c.cuentas), tiene = x => ok.has(String(x.cuenta));
  return {
    meta: { actualizado: full.meta.actualizado, mep: full.meta.mep, mepNasini: full.meta.mepNasini, tenenciasAl: full.meta.tenenciasAl, diasProx: full.meta.diasProx, ahora: Date.now() },
    lines: full.lines.filter(tiene), cuentas: full.cuentas.filter(tiene), cheques: full.cheques.filter(tiene),
    alertas: full.alertas.filter(x => String(x.cuentas).split(/[,\s]+/).some(k => ok.has(k))),
    movimientos: full.movimientos.filter(tiene),
    informes: null, noticias: full.noticias, mercado: full.mercado, historial: historialCuentas_(ss, ok),
    acceso: { rol: 'cliente', nombre: c.nombre, preview: !!preview },
  };
}

function historialCuentas_(ss, ok) {
  return historial_(ss).slice(0, 250).reverse().map(h => {
    const cs = (h.cuentas || []).filter(x => !ok || ok.has(String(x.cuenta)));
    return { fecha: h.fecha, cuentas: cs, cartera: cs.reduce((s, x) => s + (Number(x.cartera) || 0), 0), capital: cs.reduce((s, x) => s + (Number(x.capital) || 0), 0) };
  }).filter(h => h.cuentas.length);
}

function mercadoPublico_(inf) {
  const src = inf && (inf.postmarket && (!inf.premarket || inf.postmarket.generado >= inf.premarket.generado) ? inf.postmarket : inf.premarket);
  if (!src) return null;
  const dol = src.rueda ? src.rueda.dolares : src.argentina ? src.argentina.dolares : [];
  const rp = src.rueda ? src.rueda.riesgo : src.argentina ? src.argentina.riesgo : null;
  return { generado: src.generado, dolares: dol || [], riesgo: rp, macro: (src.macro || []).filter(x => ['eeuu', 'riesgo', 'commodities'].includes(x.grupo)) };
}

// Muestra los accesos (menú de la planilla).
function mostrarAccesos() {
  const ui = SpreadsheetApp.getUi();
  const a = accesos_(SpreadsheetApp.getActive());
  ui.alert('Accesos', 'Los links personales y las claves están en la pestaña "Accesos".\n\n' +
    a.map(x => '· ' + x.nombre + ' (' + (x.rol === 'admin' ? 'todas las cuentas' : x.cuentas.join(', ')) + ')' + (x.activo ? '' : ' — DESACTIVADO')).join('\n') +
    '\n\nPara cambiar una clave, borrala y volvé a abrir este menú: se genera una nueva. Para revocar un link, borrá el token. Para desactivar, poné NO en Activo.', ui.ButtonSet.OK);
}

// ================================================================ noticias generales
const NOTICIAS_FEEDS = [
  ['Ámbito', 'https://www.ambito.com/rss/pages/finanzas.xml'],
  ['Ámbito', 'https://www.ambito.com/rss/pages/economia.xml'],
  ['Infobae', 'https://www.infobae.com/arc/outboundfeeds/rss/category/economia/'],
  ['El Cronista', 'https://www.cronista.com/arc/outboundfeeds/rss/?outputType=xml'],
];
const CATEGORIAS = [
  ['Dólar', ['dólar', 'dolar', 'tipo de cambio', 'bcra', 'reservas', 'cepo', 'blue', 'mep', 'ccl', 'devaluación']],
  ['Mercados', ['bonos', 'acciones', 'merval', 'riesgo país', 'riesgo pais', 'adr', 'wall street', 'cedear', 'bolsa', 'byma', 'caución', 'licitación', 'letras', 'obligaciones negociables', 'inversores', 's&p']],
  ['Internacional', ['ee.uu', 'estados unidos', 'china', 'fed ', 'reserva federal', 'trump', 'europa', 'brasil', 'mundial', 'global']],
  ['Empresas', ['empresa', 'compañía', 'balance', 'ganancias', 'inversión', 'ypf', 'galicia', 'pampa', 'mercadolibre', 'vaca muerta', 'industria']],
  ['Economía', ['inflación', 'actividad', 'impuesto', 'gobierno', 'fmi', 'salario', 'consumo', 'déficit', 'recaudación', 'economía', 'caputo', 'milei']],
];
function categoria_(t) {
  t = t.toLowerCase();
  for (const [c, ps] of CATEGORIAS) if (ps.some(p => t.includes(p))) return c;
  return 'Economía';
}
function actualizarNoticias_(forzar) {
  const props = PropertiesService.getScriptProperties();
  const ult = Number(props.getProperty('noticiasTs') || 0);
  if (!forzar && Date.now() - ult < 55 * 60000) return 'Noticias al día.';
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return 'Actualización de noticias en curso.';
  try {
    const desde = Date.now() - 4 * 86400000, items = [], vistos = new Set();
    NOTICIAS_FEEDS.forEach(([fuente, url]) => {
      try {
        const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
        if (r.getResponseCode() !== 200) return;
        (r.getContentText().match(/<item>[\s\S]*?<\/item>/g) || []).forEach(it => {
          const g = tag => { const m = it.match(new RegExp('<' + tag + '>(?:<!\\[CDATA\\[)?([\\s\\S]*?)(?:\\]\\]>)?</' + tag + '>')); return m ? m[1].trim() : ''; };
          const limpia = s => s.replace(/<[^>]+>/g, ' ').replace(/&#x([0-9a-f]+);/gi, (m, h) => String.fromCharCode(parseInt(h, 16))).replace(/&#(\d+);/g, (m, d) => String.fromCharCode(+d))
            .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/\s+/g, ' ').trim();
          const titulo = limpia(g('title')), link = g('link'), fecha = Date.parse(g('pubDate'));
          if (!titulo || !/^https:\/\//.test(link) || !(fecha >= desde)) return;
          if (/cronista\.com\//i.test(link) && !/cronista\.com\/(economia-politica|finanzas-mercados|negocios)\//i.test(link)) return;   // solo secciones argentinas
          const key = titulo.toLowerCase().replace(/[^a-z0-9áéíóúñ]/g, '').slice(0, 60);
          if (vistos.has(key)) return; vistos.add(key);
          let resumen = limpia(g('description'));
          if (resumen.length > 240) resumen = resumen.slice(0, 237).replace(/\s+\S*$/, '') + '…';
          items.push([new Date(fecha), fuente, categoria_(titulo + ' ' + resumen), titulo, link, resumen]);
        });
      } catch (e) { /* una fuente caída no frena el resto */ }
    });
    if (!items.length) return 'Sin noticias nuevas (fuentes no disponibles).';
    items.sort((a, b) => b[0] - a[0]);
    const sh = hoja_(SpreadsheetApp.getActive(), 'Noticias', ['Fecha', 'Fuente', 'Categoría', 'Título', 'Link', 'Resumen']);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
    const top = items.slice(0, 90);
    sh.getRange(2, 1, top.length, 6).setValues(top);
    props.setProperty('noticiasTs', String(Date.now()));
    return 'Noticias actualizadas: ' + top.length + '.';
  } finally { lock.releaseLock(); }
}
function noticiasGuardadas_(ss) {
  const sh = ss.getSheetByName('Noticias');
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, Math.min(sh.getLastRow() - 1, 90), 6).getValues().filter(r => r[3]).map(r => ({
    fecha: r[0] instanceof Date ? r[0].getTime() : null, fuente: r[1], categoria: r[2], titulo: r[3], link: /^https:\/\//.test(r[4]) ? r[4] : '', resumen: r[5] }));
}
function actualizarNoticiasAhora() { const ui = SpreadsheetApp.getUi(); ui.alert(actualizarNoticias_(true)); }

// ================================================================ página web
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Carteras en vivo')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);   // permite insertarla en Google Sites
}

// Todo lo que muestra la página (uso interno: la página accede con datos(sesión)). Fechas en milisegundos.
function datosCompletos_() {
  const ss = SpreadsheetApp.getActive();
  const C = cfg_(ss);
  const ms = d => d instanceof Date ? d.getTime() : null;
  const num = v => typeof v === 'number' && isFinite(v) ? v : null;

  const X = ss.getSheetByName('Posiciones');
  const lines = X.getRange(2, 1, Math.max(X.getLastRow() - 1, 1), 21).getValues().filter(r => r[2] && Number(r[9])).map(r => ({
    cuenta: String(r[0]), cliente: r[1], ticker: String(r[2]), instrumento: r[3], clase: r[4], moneda: r[5], div: Number(r[6]) || 1,
    cantNasini: num(r[7]), cantidad: num(r[9]), ppc: num(r[12]), precioNasini: num(r[11]), precio: num(r[13]), fuente: r[14],
    varPct: num(r[15]), valorArs: num(r[16]), varHoy: num(r[17]), costoArs: num(r[18]) }));

  const K = ss.getSheetByName('Cuentas');
  const cuentas = K.getRange(2, 1, Math.max(K.getLastRow() - 1, 1), 11).getValues().filter(r => /^\d+$/.test(String(r[0]))).map(r => ({
    cuenta: String(r[0]), cliente: r[1],
    dineroARS: (num(r[2]) || 0) + (num(r[4]) || 0), dineroUSD: (num(r[3]) || 0) + (num(r[5]) || 0),   // efectivo neto + caución = dinero
    caucARS: num(r[4]) || 0, caucUSD: num(r[5]) || 0, movARS: num(r[6]) || 0, movUSD: num(r[7]) || 0, chqAcred: num(r[8]) || 0,
    efARS: num(r[9]) || 0, efUSD: num(r[10]) || 0 }));

  const cheques = chequesRaw_(ss).map(c => ({ cuenta: c.cuenta, cliente: c.cliente, codigo: c.codigo, librador: c.librador, numero: c.numero,
    vto: ms(c.vto), acred: ms(c.acred), nominal: c.nominal, estado: c.estado, valor: c.valor,
    factor: c.nominal ? c.valor / c.nominal : null }));

  const A = ss.getSheetByName('Alertas');
  const alertas = A.getLastRow() > 1 ? A.getRange(2, 1, Math.min(A.getLastRow() - 1, 100), 6).getValues().filter(r => r[2]).map(r => ({
    fecha: ms(r[0]), tipo: r[1], titulo: r[2], detalle: r[3], cuentas: String(r[4]), monto: num(r[5]) })) : [];

  const M = ss.getSheetByName('Movimientos');
  const movimientos = M.getLastRow() > 1 ? M.getRange(2, 1, M.getLastRow() - 1, 9).getValues().filter(r => r[1] && r[3]).map(r => ({
    fecha: ms(r[0]), cuenta: String(r[1]), ticker: String(r[2]), tipo: r[3], cantidad: num(r[4]), precio: num(r[5]), moneda: r[6],
    importe: num(r[7]), nota: r[8] })).slice(-50).reverse() : [];

  return {
    meta: { actualizado: ms(C.getRange(CFG.actualizado).getValue()), mep: num(C.getRange(CFG.mepUsado).getValue()),
      mepNasini: num(C.getRange('B3').getValue()), tenenciasAl: String(C.getRange(CFG.tenencias).getDisplayValue()),
      estado: String(C.getRange(CFG.estado).getValue()), diasProx: Number(C.getRange(CFG.diasProx).getValue()) || 30, ahora: Date.now() },
    lines, cuentas, cheques, alertas, movimientos, informes: informesParaWeb_(ss),
  };
}
