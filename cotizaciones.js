/* ===== Módulo Cotizaciones =====
   Colección Firestore: cotizaciones
   {
     cliente, descripcion, cantidad, link, imagenes: [dataURL JPEG comprimido],
     estado: 'pendiente' | 'cotizada' | 'finalizada',
     costos: { totalUnidades, valorUnidad, aumentoRmb (por unidad), unidadesPorCaja, cbmCaja, fleteId, fleteValor, pesoCaja, broker,
               comisionPct, usdCop, usdRmb  <- comisión y tasas vigentes al guardar, para que la cotización no cambie después },
     creadoPor, creadoEn
   }
   Las imágenes se comprimen y se guardan como base64 dentro del documento
   (igual que el logo). Por eso hay un máximo de imágenes y de peso total. */

const COT_MAX_IMAGENES = 4;
const COT_MAX_LADO_PX = 900;          // lado mayor de cada imagen tras comprimir
const COT_MAX_CHARS_TOTAL = 800000;   // Firestore limita cada documento a ~1 MB

let cotDesuscribir = null;
let cotLista = [];                    // cotizaciones cargadas, para buscarlas por id
let cotImagenesForm = [];             // imágenes (dataURL) del formulario abierto
let cotImagenesReales = [];           // imágenes reales cotizadas (formulario de costos)
const cotAbiertas = new Set();
const cotSeleccion = new Set();       // ids de tarjetas marcadas para el Excel múltiple        // ids de tarjetas expandidas
let cotEditandoId = null;             // id de la cotización a la que se le llenan costos

const cotEsc = s => String(s ?? '').replace(/[&<>"']/g, c =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const cotLinkSeguro = u => /^https?:\/\//i.test(u || '') ? u : '';
const cotRef = () => db.collection('cotizaciones');
const cotNum = (n, dec = 2) => Number(n).toLocaleString('es-CO', { minimumFractionDigits: 0, maximumFractionDigits: dec });

/* ---------- Estructura (sección + botón + modales) ---------- */
function cotConstruirUI() {
  const sec = document.getElementById('cotizaciones');
  if (!sec) return false;
  sec.innerHTML = `
    <div class="cot-barra"><h2 class="cot-titulo">Cotizaciones</h2>
      <div class="cot-filtros">
        <input type="search" id="cotBuscar" placeholder="Buscar por cliente, producto, comercial, estado…" aria-label="Buscar cotizaciones" autocomplete="off">
        <select id="cotFiltroEstado" aria-label="Filtrar por estado"><option value="">Todos los estados</option><option value="pendiente">Pendiente</option><option value="cotizada">Cotizada</option><option value="finalizada">Finalizada</option></select>
        <select id="cotFiltroOrigen" aria-label="Filtrar por origen"><option value="">Todos los orígenes</option><option>Adma Company</option><option>Groupack</option></select>
      </div>
      <p id="cotConteo" class="cot-conteo" aria-live="polite"></p>
      <div class="cot-acciones">
        <button type="button" class="cot-enlace" id="cotBtnSelVisibles">Seleccionar pendientes visibles</button>
        <span id="cotSelTexto" class="cot-conteo" style="flex-basis:auto"></span>
        <button type="button" class="btn btn-chico" id="cotBtnExcelSel" hidden></button>
        <button type="button" class="cot-enlace" id="cotBtnLimpiarSel" hidden>Quitar selección</button>
      </div></div>
    <div id="cotGrid" class="cot-grid" aria-live="polite"><p class="lista-vacia">Cargando cotizaciones…</p></div>
    <button type="button" id="cotBtnNueva" class="cot-fab" aria-label="Nueva cotización" title="Nueva cotización">+</button>`;

  const modales = document.createElement('div');
  modales.innerHTML = `
  <div id="cotModalNueva" class="modal" hidden>
    <div class="modal-fondo" data-cot-cerrar></div>
    <form id="cotFormNueva" class="modal-tarjeta cot-modal" novalidate>
      <div class="modal-cabecera"><h2>Nueva cotización</h2>
        <button type="button" class="modal-cerrar" data-cot-cerrar aria-label="Cerrar">✕</button></div>
      <div id="cotErrorNueva" class="login-error" role="alert" hidden></div>
      <label class="login-campo"><span>Nombre del cliente</span><input type="text" id="cotCliente" required></label>
      <label class="login-campo"><span>Cliente proveniente de</span><select id="cotOrigen" required><option value="">Selecciona…</option><option value="Adma Company">Adma Company</option><option value="Groupack">Groupack</option></select></label>
      <label class="login-campo" id="cotComercialWrap" hidden><span>Nombre del comercial que proporcionó el cliente</span><input type="text" id="cotComercial" autocomplete="off"></label>
      <label class="login-campo"><span>Descripción del producto</span><textarea id="cotDescripcion" rows="3" required></textarea></label>
      <label class="login-campo"><span>Cantidad</span><input type="number" id="cotCantidad" min="1" step="1" required></label>
      <label class="login-campo"><span>Link</span><input type="url" id="cotLink" placeholder="https://" autocapitalize="none"></label>
      <div class="login-campo"><span>Imágenes de referencia (máx. ${COT_MAX_IMAGENES})</span>
        <input type="file" id="cotArchivos" accept="image/*" multiple>
        <div id="cotPrevias" class="cot-previas"></div></div>
      <div class="modal-acciones">
        <button type="button" class="btn secundario" data-cot-cerrar>Cancelar</button>
        <button type="submit" class="btn" id="cotBtnGuardar">Crear cotización</button>
      </div>
    </form>
  </div>

  <div id="cotModalCostos" class="modal" hidden>
    <div class="modal-fondo" data-cot-cerrar></div>
    <form id="cotFormCostos" class="modal-tarjeta cot-modal" novalidate>
      <div class="modal-cabecera"><h2 id="cotCostosTitulo">Costos</h2>
        <button type="button" class="modal-cerrar" data-cot-cerrar aria-label="Cerrar">✕</button></div>
      <div id="cotErrorCostos" class="login-error" role="alert" hidden></div>
      <p class="cot-dato cot-calc" id="cotResumenCajas" aria-live="polite"></p>
      <label class="login-campo"><span>Valor por unidad del producto (RMB)</span><input type="number" id="cotValorUnidad" min="0" step="0.01" required></label>
      <div class="login-campo"><span>Aumento de RMB</span>
        <div class="cot-aumento"><select id="cotAumentoTipo" aria-label="Tipo de aumento"><option value="unidad">Por unidad</option><option value="porcentaje">Por porcentaje</option></select>
        <input type="number" id="cotAumento" min="0" step="0.01" required aria-label="Valor del aumento"></div></div>
      <label class="login-campo"><span>Unidades por caja</span><input type="number" id="cotUnidadesCaja" min="1" step="1" required></label>
      <label class="login-campo"><span>CBM por caja</span><input type="number" id="cotCbm" min="0" step="0.0001" required></label>
      <label class="login-campo"><span>Vía de transporte</span><select id="cotVia" required><option value="maritima">Marítima</option><option value="aerea">Aérea</option></select></label>
      <label class="login-campo" id="cotFleteWrap"><span>Precio del flete marítimo (COP por m³)</span><select id="cotFlete"></select></label>
      <label class="login-campo" id="cotFleteAereoWrap" hidden><span>Valor total del flete aéreo (USD)</span><input type="number" id="cotFleteAereo" min="0" step="0.01"></label>
      <label class="login-campo"><span>Peso por caja (kg) — opcional</span><input type="number" id="cotPesoCaja" min="0" step="0.01"></label>
      <label class="login-campo"><span>Broker</span><input type="text" id="cotBroker" autocomplete="off"></label>
      <div class="login-campo"><span>Imágenes reales cotizadas (máx. ${COT_MAX_IMAGENES}) — se usarán en el PDF de cotización</span>
        <input type="file" id="cotArchivosReales" accept="image/*" multiple>
        <div id="cotPreviasReales" class="cot-previas"></div></div>
      <div class="modal-acciones">
        <button type="button" class="btn secundario" data-cot-cerrar>Cancelar</button>
        <button type="submit" class="btn" id="cotBtnGuardarCostos">Guardar</button>
      </div>
    </form>
  </div>

  <div id="cotModalPago" class="modal" hidden>
    <div class="modal-fondo" data-cot-cerrar></div>
    <form id="cotFormPago" class="modal-tarjeta cot-modal" novalidate>
      <div class="modal-cabecera"><h2 id="cotPagoTitulo">Registrar pago</h2>
        <button type="button" class="modal-cerrar" data-cot-cerrar aria-label="Cerrar">✕</button></div>
      <div id="cotErrorPago" class="login-error" role="alert" hidden></div>
      <p class="cot-dato cot-calc" id="cotPagoResumen" aria-live="polite"></p>
      <label class="login-campo"><span>¿Qué estás pagando?</span><select id="cotPagoConcepto" required>
        <option value="producto">Precio del producto</option><option value="flete">Precio del flete</option><option value="todo">Todo (producto + flete)</option></select></label>
      <label class="login-campo"><span>Valor pagado (COP)</span><input type="number" id="cotPagoMonto" min="0" step="any" required></label>
      <button type="button" class="cot-enlace" id="cotPagoSaldo" style="align-self:flex-start">Usar todo lo que falta por pagar</button>
      <label class="login-campo"><span>Fecha del pago</span><input type="date" id="cotPagoFecha"></label>
      <label class="login-campo"><span>Nota (opcional)</span><input type="text" id="cotPagoNota" maxlength="200" autocomplete="off"></label>
      <div class="login-campo"><span>Comprobante de pago (obligatorio, imagen)</span><input type="file" id="cotPagoArchivo" accept="image/*" required></div>
      <div class="modal-acciones">
        <button type="button" class="btn secundario" data-cot-cerrar>Cancelar</button>
        <button type="submit" class="btn" id="cotBtnGuardarPago">Guardar pago</button>
      </div>
    </form>
  </div>`;
  document.getElementById('appShell').appendChild(modales);
  return true;
}

/* ---------- Abrir / cerrar modales ---------- */
const cotAbrir = id => { document.getElementById(id).hidden = false; };
function cotCerrarTodo() {
  document.querySelectorAll('#cotModalNueva, #cotModalCostos, #cotModalPago').forEach(m => { m.hidden = true; });
}
function cotError(id, msg) {
  const el = document.getElementById(id);
  el.textContent = msg || ''; el.hidden = !msg;
}

/* ---------- Imágenes: leer, comprimir, previsualizar ---------- */
function cotComprimir(archivo, maxLado = COT_MAX_LADO_PX, calidad = 0.72) {
  return new Promise((ok, fallo) => {
    const lector = new FileReader();
    lector.onerror = () => fallo(lector.error);
    lector.onload = () => {
      const img = new Image();
      img.onerror = () => fallo(new Error('Imagen no válida'));
      img.onload = () => {
        const k = Math.min(1, maxLado / Math.max(img.width, img.height));
        const c = document.createElement('canvas');
        c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
        const ctx = c.getContext('2d');
        ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height); // fondo blanco para PNG con transparencia
        ctx.drawImage(img, 0, 0, c.width, c.height);
        ok(c.toDataURL('image/jpeg', calidad));
      };
      img.src = lector.result;
    };
    lector.readAsDataURL(archivo);
  });
}

function cotPintarPrevias() {
  const cont = document.getElementById('cotPrevias');
  cont.innerHTML = cotImagenesForm.map((src, i) =>
    `<div class="cot-previa"><img src="${src}" alt="Imagen ${i + 1}">
     <button type="button" data-cot-quitar="${i}" aria-label="Quitar imagen ${i + 1}">✕</button></div>`).join('');
}

async function cotManejarArchivos(e) {
  const archivos = [...e.target.files];
  e.target.value = '';
  cotError('cotErrorNueva', '');
  for (const f of archivos) {
    if (cotImagenesForm.length >= COT_MAX_IMAGENES) { cotError('cotErrorNueva', `Máximo ${COT_MAX_IMAGENES} imágenes.`); break; }
    if (!f.type.startsWith('image/')) { cotError('cotErrorNueva', 'Solo se permiten imágenes.'); continue; }
    try { cotImagenesForm.push(await cotComprimir(f)); }
    catch (err) { console.error(err); cotError('cotErrorNueva', 'No se pudo leer una de las imágenes.'); }
  }
  cotPintarPrevias();
}

/* ---------- Crear cotización ---------- */
function cotAbrirNueva() {
  document.getElementById('cotFormNueva').reset();
  cotImagenesForm = []; cotPintarPrevias(); cotError('cotErrorNueva', '');
  cotToggleComercial();
  cotAbrir('cotModalNueva');
  document.getElementById('cotCliente').focus();
}

async function cotGuardarNueva(e) {
  e.preventDefault();
  const cliente = document.getElementById('cotCliente').value.trim();
  const descripcion = document.getElementById('cotDescripcion').value.trim();
  const cantidad = parseInt(document.getElementById('cotCantidad').value, 10);
  const link = document.getElementById('cotLink').value.trim();

  const origen = document.getElementById('cotOrigen').value;
  const comercial = origen === 'Adma Company' ? document.getElementById('cotComercial').value.trim() : '';
  if (!cliente || !descripcion) return cotError('cotErrorNueva', 'Escribe el nombre del cliente y la descripción.');
  if (!origen) return cotError('cotErrorNueva', 'Indica de dónde proviene el cliente.');
  if (origen === 'Adma Company' && !comercial) return cotError('cotErrorNueva', 'Escribe el nombre del comercial que proporcionó el cliente.');
  if (!(cantidad > 0)) return cotError('cotErrorNueva', 'La cantidad debe ser mayor que cero.');
  if (link && !cotLinkSeguro(link)) return cotError('cotErrorNueva', 'El link debe empezar por http:// o https://');
  if (cotImagenesForm.join('').length > COT_MAX_CHARS_TOTAL) return cotError('cotErrorNueva', 'Las imágenes pesan demasiado. Quita alguna.');

  const btn = document.getElementById('cotBtnGuardar');
  btn.disabled = true;
  try {
    await cotRef().add({
      cliente, origen, comercial, descripcion, cantidad, link, imagenes: cotImagenesForm,
      estado: 'pendiente', costos: null,
      creadoPor: sesionActual ? sesionActual.nombre : '',
      creadoEn: firebase.firestore.FieldValue.serverTimestamp()
    });
    cotCerrarTodo();
  } catch (err) {
    console.error('Error al crear la cotización:', err);
    cotError('cotErrorNueva', 'No se pudo guardar. Revisa tu conexión y las reglas de Firestore.');
  } finally { btn.disabled = false; }
}

/* ---------- Costos (botón Continuar) ---------- */
/* Fletes de Configuración > Datos; si la cotización ya usaba uno que luego se borró, se conserva su copia guardada. */
function cotFletesDisponibles(c) {
  const lista = (typeof DATOS !== 'undefined' ? DATOS.fletes : []).slice();
  const k = c && c.costos;
  if (k && k.fleteId && !lista.some(f => f.id === k.fleteId))
    lista.push({ id: k.fleteId, valor: k.fleteValor });
  return lista;
}

function cotAbrirCostos(id) {
  const c = cotLista.find(x => x.id === id);
  if (!c) return;
  cotEditandoId = id;
  document.getElementById('cotFormCostos').reset();
  cotError('cotErrorCostos', '');
  document.getElementById('cotCostosTitulo').textContent = 'Costos · ' + c.cliente;

  const sel = document.getElementById('cotFlete');
  const fletes = cotFletesDisponibles(c);
  sel.innerHTML = '<option value="">Selecciona un flete…</option>' +
    fletes.map(f => `<option value="${cotEsc(f.id)}">${cotNum(f.valor)} COP</option>`).join('');
  if (!fletes.length) cotError('cotErrorCostos', 'Aún no hay fletes. Agrégalos en Configuración > Datos.');

  const k = c.costos || {};
  document.getElementById('cotValorUnidad').value = k.valorUnidad ?? '';
  document.getElementById('cotAumentoTipo').value = k.aumentoTipo === 'porcentaje' ? 'porcentaje' : 'unidad';
  document.getElementById('cotAumento').value = k.aumentoRmb ?? '';
  cotImagenesReales = (c.imagenesReales || []).slice(); cotPintarPreviasReales();
  document.getElementById('cotUnidadesCaja').value = k.unidadesPorCaja ?? '';
  document.getElementById('cotCbm').value = k.cbmCaja ?? '';
  document.getElementById('cotPesoCaja').value = k.pesoCaja ?? '';
  document.getElementById('cotBroker').value = k.broker ?? '';
  if (k.fleteId) sel.value = k.fleteId;
  document.getElementById('cotVia').value = k.viaFlete === 'aerea' ? 'aerea' : 'maritima';
  document.getElementById('cotFleteAereo').value = k.viaFlete === 'aerea' ? (k.fleteValor ?? '') : '';
  cotToggleVia();
  cotAbrir('cotModalCostos');
  cotActualizarCajas();
  document.getElementById('cotUnidadesCaja').focus();
}

async function cotGuardarCostos(e) {
  e.preventDefault();
  const c = cotLista.find(x => x.id === cotEditandoId);
  const via = document.getElementById('cotVia').value;
  const aereo = parseFloat(document.getElementById('cotFleteAereo').value);
  const flete = via === 'aerea'
    ? (aereo >= 0 ? { id: '', valor: aereo } : null)
    : cotFletesDisponibles(c).find(f => f.id === document.getElementById('cotFlete').value);
  if (typeof TASAS === 'undefined' || !(TASAS.usdCop > 0) || !(TASAS.usdRmb > 0))
    return cotError('cotErrorCostos', 'Aún no se cargan las tasas de cambio. Espera un momento e inténtalo de nuevo.');
  // cajas = cantidad de la cotización / unidades por caja (redondeado hacia arriba); total = cajas × unidades por caja
  const upc = parseInt(document.getElementById('cotUnidadesCaja').value, 10);
  const cajas = upc > 0 && c ? Math.ceil(c.cantidad / upc) : NaN;
  const costos = {
    cajas,
    totalUnidades: cajas * upc,
    valorUnidad: parseFloat(document.getElementById('cotValorUnidad').value),
    aumentoTipo: document.getElementById('cotAumentoTipo').value,
    aumentoRmb: parseFloat(document.getElementById('cotAumento').value),
    unidadesPorCaja: parseInt(document.getElementById('cotUnidadesCaja').value, 10),
    cbmCaja: parseFloat(document.getElementById('cotCbm').value),
    viaFlete: via,
    fleteId: flete ? flete.id : '',
    fleteValor: flete ? Number(flete.valor) : null,
    pesoCaja: (v => Number.isNaN(v) ? null : v)(parseFloat(document.getElementById('cotPesoCaja').value)),
    broker: document.getElementById('cotBroker').value.trim(),
    comisionPct: DATOS.comisionAdma,
    usdCop: TASAS.usdCop,
    usdRmb: TASAS.usdRmb
  };
  if (!(costos.totalUnidades > 0) || !(costos.valorUnidad >= 0) || !(costos.aumentoRmb >= 0) || !(costos.unidadesPorCaja > 0) ||
      !(costos.cbmCaja >= 0) || !flete)
    return cotError('cotErrorCostos', 'Completa todos los campos con valores válidos y indica el flete.');

  if (c && (c.imagenes || []).join('').length + cotImagenesReales.join('').length > COT_MAX_CHARS_TOTAL + 100000)
    return cotError('cotErrorCostos', 'Entre las imágenes de referencia y las reales el peso es demasiado. Quita alguna imagen real.');
  const btn = document.getElementById('cotBtnGuardarCostos');
  btn.disabled = true;
  try {
    await cotRef().doc(cotEditandoId).update({ costos, imagenesReales: cotImagenesReales, estado: 'cotizada' });
    cotCerrarTodo();
  } catch (err) {
    console.error('Error al guardar costos:', err);
    cotError('cotErrorCostos', 'No se pudo guardar. Inténtalo de nuevo.');
  } finally { btn.disabled = false; }
}

/* ---------- Eliminar ---------- */
async function cotEliminar(id) {
  const c = cotLista.find(x => x.id === id);
  if (!c || !confirm(`¿Eliminar la cotización de "${c.cliente}"? Esta acción no se puede deshacer.`)) return;
  try { await cotRef().doc(id).delete(); }
  catch (err) { console.error(err); alert('No se pudo eliminar la cotización.'); }
}

/* ---------- Estados y cálculos ---------- */
/* pendiente (rojo) -> cotizada (naranja, ya tiene costos) -> finalizada (verde, orden de compra generada) */
const COT_ESTADOS = {
  pendiente:  'Pendiente',
  cotizada:   'Cotizada',
  finalizada: 'Finalizada'
};
const cotEstadoDe = c => c.estado === 'costeada' ? 'cotizada' : (COT_ESTADOS[c.estado] ? c.estado : 'pendiente');

/* Cálculos del resumen. Usa la comisión y las tasas guardadas con la cotización; lo que no se pueda calcular (datos faltantes) se omite. */
const cotOk = v => typeof v === 'number' && !Number.isNaN(v);
function cotCalculos(c) {
  const k = c.costos;
  if (!k) return null;
  const u = cotOk(k.totalUnidades) && k.totalUnidades > 0 ? k.totalUnidades : c.cantidad;
  const r = { unidades: u };
  if (cotOk(k.unidadesPorCaja) && k.unidadesPorCaja > 0) r.cajas = cotOk(k.cajas) ? k.cajas : Math.ceil(u / k.unidadesPorCaja);
  if (cotOk(k.valorUnidad) && cotOk(k.aumentoRmb)) {
    const porc = k.aumentoTipo === 'porcentaje';
    r.precioUnidad = k.valorUnidad;
    r.totalSinComision = u * k.valorUnidad;
    r.aumentoValor = k.aumentoRmb;
    r.aumentoPorc = porc;
    r.aumentoSuf = porc ? ' %' : ' por unidad';
    if (cotOk(k.comisionPct)) {
      r.comisionPct = k.comisionPct;
      // Valor por unidad con aumento y con la comisión de Adma
      const unitAumentado = porc ? k.valorUnidad * (1 + k.aumentoRmb / 100) : k.valorUnidad + k.aumentoRmb;
      r.unitConComision = unitAumentado * (1 + k.comisionPct / 100);
      r.totalRmb = r.unitConComision * u;
      r.comisionTotal = (r.unitConComision - unitAumentado) * u;   // monto de la comisión sobre todo el pedido
    }
  }
  if (r.cajas != null && cotOk(k.cbmCaja)) r.cbmTotal = k.cbmCaja * r.cajas;
  if (cotOk(k.pesoCaja)) r.pesoCaja = k.pesoCaja;
  r.viaAerea = k.viaFlete === 'aerea';
  if (cotOk(k.fleteValor)) {
    if (r.viaAerea) {
      // Aérea: el valor ingresado es el flete total en USD (se muestra en USD; en COP solo para el resumen)
      r.logistica360Usd = k.fleteValor;
      if (cotOk(k.usdCop)) r.logistica360 = k.fleteValor * k.usdCop;
    } else if (r.cbmTotal != null) {
      // Marítima: valor de la lista de Datos (COP por m³) × CBM total
      r.logistica360 = k.fleteValor * r.cbmTotal;
    }
    if (r.logistica360 != null && u > 0) r.logisticaUnit = r.logistica360 / u;
  }
  if (r.totalRmb != null && cotOk(k.usdRmb) && k.usdRmb > 0 && cotOk(k.usdCop)) {
    r.totalCop = (r.totalRmb / k.usdRmb) * k.usdCop;
    r.precioUnitCop = r.totalCop / u;
    if (r.logisticaUnit != null) r.totalUnitCop = r.precioUnitCop + r.logisticaUnit;
    if (r.logistica360 != null) {
      r.totalConFlete = r.totalCop + r.logistica360;
      r.unitConFlete = r.totalConFlete / u;
    }
  }
  return r;
}

/* Resumen en tres grupos: [{ titulo, lineas: [[etiqueta, valor], ...] }] para la tarjeta y el PDF. */
function cotLineasCostos(c) {
  const r = cotCalculos(c);
  if (!r) return [];
  // Cada definición: [etiqueta, valor, decimales, sufijo, prefijo]. ¥ = RMB, $ = pesos (COP) o dólares (USD).
  const grupo = (titulo, defs) => ({
    titulo,
    lineas: defs.filter(([, v]) => v != null).map(([txt, v, dec, suf = '', pre = '']) => [txt, pre + cotNum(v, dec) + suf])
  });
  const pct = r.comisionPct != null ? cotNum(r.comisionPct, 2) + '% ' : '';
  const Y = '¥ ', P = '$ ';
  return [
    grupo('PRECIO PRODUCTO', [
      ['Total unidades cotizadas', r.unidades, 0],
      ['Precio real por unidad', r.precioUnidad, 2, '', Y],
      ['Total real sin comisión', r.totalSinComision, 2, '', Y],
      ['Aumento del RMB', r.aumentoValor, 2, r.aumentoSuf, r.aumentoPorc ? '' : Y],
      ['Comisión de Adma' + (pct ? ` (${pct.trim()})` : ''), r.comisionTotal, 2, ' del total', Y],
      ['Total para el cliente con comisión', r.totalRmb, 2, '', Y]
    ]),
    grupo(r.viaAerea ? 'FLETE AÉREO SUJETO A IMPUESTOS' : 'FLETE MARÍTIMO', [
      ['CBM', r.cbmTotal, 4, ' m3'],
      ['Cantidad de cajas', r.cajas, 0],
      ['Peso por caja', r.pesoCaja, 2, ' kg'],
      ['Logística unitaria en COP', r.viaAerea ? null : r.logisticaUnit, 2, ' COP', P],
      ['Logística 360 COP', r.viaAerea ? null : r.logistica360, 2, ' COP', P],
      ['Logística 360 USD', r.viaAerea ? r.logistica360Usd : null, 2, ' USD', P]
    ]),
    grupo('RESUMEN', [
      ['Valor total del producto sin flete', r.totalCop, 2, ' COP', P],
      ['Valor total del producto con flete', r.totalConFlete, 2, ' COP', P],
      ['Valor por unidad del producto con flete', r.unitConFlete, 2, ' COP', P]
    ])
  ].filter(g => g.lineas.length);
}

/* ---------- Buscador ---------- */
const cotNorm = s => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
/* Búsqueda sin tildes ni mayúsculas; cada palabra escrita debe aparecer en algún dato de la cotización. */
function cotFiltradas() {
  const q = cotNorm(document.getElementById('cotBuscar').value).split(/\s+/).filter(Boolean);
  const est = document.getElementById('cotFiltroEstado').value;
  const ori = document.getElementById('cotFiltroOrigen').value;
  return cotLista.filter(c => {
    const e = cotEstadoDe(c);
    if (est && e !== est) return false;
    if (ori && c.origen !== ori) return false;
    if (!q.length) return true;
    const k = c.costos || {};
    const pajar = cotNorm([c.cliente, c.descripcion, c.origen, c.comercial, COT_ESTADOS[e], k.broker, c.link, c.creadoPor, c.cantidad].join(' '));
    return q.every(t => pajar.includes(t));
  });
}

/* ---------- Tarjetas (compactas; al hacer clic se expanden) ---------- */
function cotPintar() {
  cotActualizarAcciones();
  const grid = document.getElementById('cotGrid');
  const conteo = document.getElementById('cotConteo');
  if (!cotLista.length) {
    conteo.textContent = '';
    grid.innerHTML = '<p class="lista-vacia">Aún no hay cotizaciones. Toca el botón + para crear la primera.</p>';
    return;
  }
  const lista = cotFiltradas();
  conteo.textContent = lista.length === cotLista.length ? `${cotLista.length} cotizaciones` : `${lista.length} de ${cotLista.length} cotizaciones`;
  if (!lista.length) {
    grid.innerHTML = '<p class="lista-vacia">Ninguna cotización coincide con la búsqueda.</p>';
    return;
  }
  grid.innerHTML = lista.map(c => {
    const est = cotEstadoDe(c);
    const abierta = cotAbiertas.has(c.id);
    const link = cotLinkSeguro(c.link);
    const foto = (c.imagenes || [])[0];
    const costos = c.costos ? cotLineasCostos(c).map(g => `<p class="cot-costos"><strong>${cotEsc(g.titulo)}</strong><br>${g.lineas.map(([a, b]) => `<strong>${cotEsc(a)}:</strong> ${cotEsc(b)}`).join('<br>')}</p>`).join('') : '';

    let botones = `<button type="button" class="btn secundario btn-chico" data-cot-excel="${c.id}">A cotizar - Generar Excel</button>`;
    if (est === 'pendiente') {
      botones += `<button type="button" class="btn btn-chico" data-cot-continuar="${c.id}">Continuar</button>`;
    } else {
      botones += `<button type="button" class="btn secundario btn-chico" data-cot-pdf="${c.id}" data-tipo="cotizacion">Cotización - Generar PDF</button>`;
      const btnModulo = `<button type="button" class="btn btn-chico" data-cot-modulo="${c.id}">Mandar a módulo</button>`;
      botones += est === 'cotizada'
        ? `${btnModulo}
           <button type="button" class="cot-enlace" data-cot-continuar="${c.id}">Editar costos</button>`
        : (c.modulo ? '' : btnModulo) +
          `<button type="button" class="btn secundario btn-chico" data-cot-orden="${c.id}">Orden de compra - Generar PDF</button>`;
    }

    return `<article class="cot-card cot-${est}${abierta ? ' abierta' : ''}${est === 'pendiente' ? ' cot-seleccionable' : ''}">
      ${est === 'pendiente' ? `<input type="checkbox" class="cot-sel" data-cot-sel="${c.id}"${cotSeleccion.has(c.id) ? ' checked' : ''} aria-label="Seleccionar la cotización de ${cotEsc(c.cliente)} para el Excel">` : ''}
      <button type="button" class="cot-resumen" data-cot-toggle="${c.id}" aria-expanded="${abierta}">
        ${foto ? `<img class="cot-foto" src="${foto}" alt="">` : '<span class="cot-foto" aria-hidden="true"></span>'}
        <span class="cot-resumen-txt"><strong>${cotEsc(c.cliente)}</strong><span class="cot-articulo">${cotEsc(c.descripcion)}</span></span>
        <span class="cot-estado">${COT_ESTADOS[est]}</span>
      </button>
      <div class="cot-detalle"${abierta ? '' : ' hidden'}>
        <p>${cotEsc(c.descripcion)}</p>
        ${c.origen ? `<p class="cot-dato"><strong>Cliente de:</strong> ${cotEsc(c.origen)}${c.comercial ? ` · <strong>Comercial:</strong> ${cotEsc(c.comercial)}` : ''}</p>` : ''}
        <p class="cot-dato"><strong>Cantidad:</strong> ${cotNum(c.cantidad, 0)}</p>
        ${c.modulo ? `<p class="cot-dato"><strong>Enviada a:</strong> ${cotEsc(cotNombreModulo(c.modulo))}</p>` : ''}
        ${link ? `<p class="cot-dato"><a href="${cotEsc(link)}" target="_blank" rel="noopener noreferrer">Ver link del producto</a></p>` : ''}
        ${(c.imagenes || []).length ? `<div class="cot-miniaturas">${c.imagenes.map((s, i) => `<img src="${s}" alt="Referencia ${i + 1} de ${cotEsc(c.cliente)}">`).join('')}</div>` : ''}
        ${costos}
        <footer>${botones}
          <button type="button" class="cot-eliminar" data-cot-eliminar="${c.id}" aria-label="Eliminar cotización de ${cotEsc(c.cliente)}">Eliminar</button>
        </footer>
      </div>
    </article>`;
  }).join('');
}

function cotAlternar(id) {
  cotAbiertas.has(id) ? cotAbiertas.delete(id) : cotAbiertas.add(id);
  cotPintar();
}

/* ---------- Ayudas de formularios ---------- */
function cotToggleComercial() {
  const esAdma = document.getElementById('cotOrigen').value === 'Adma Company';
  document.getElementById('cotComercialWrap').hidden = !esAdma;
  document.getElementById('cotComercial').required = esAdma;
}

function cotPintarPreviasReales() {
  document.getElementById('cotPreviasReales').innerHTML = cotImagenesReales.map((src, i) =>
    `<div class="cot-previa"><img src="${src}" alt="Imagen real ${i + 1}">
     <button type="button" data-cot-quitar-real="${i}" aria-label="Quitar imagen real ${i + 1}">✕</button></div>`).join('');
}

async function cotManejarArchivosReales(e) {
  const archivos = [...e.target.files];
  e.target.value = '';
  cotError('cotErrorCostos', '');
  for (const f of archivos) {
    if (cotImagenesReales.length >= COT_MAX_IMAGENES) { cotError('cotErrorCostos', `Máximo ${COT_MAX_IMAGENES} imágenes.`); break; }
    if (!f.type.startsWith('image/')) { cotError('cotErrorCostos', 'Solo se permiten imágenes.'); continue; }
    try { cotImagenesReales.push(await cotComprimir(f)); }
    catch (err) { console.error(err); cotError('cotErrorCostos', 'No se pudo leer una de las imágenes.'); }
  }
  cotPintarPreviasReales();
}

/* ---------- PDF (jsPDF) ----------
   tipo: 'cotizar'    -> solicitud "A cotizar" (datos del producto)
         'cotizacion' -> cotización con costos
         'orden'      -> orden de compra con costos */
function cotDimensionesImagen(src) {
  return new Promise(ok => { const i = new Image(); i.onload = () => ok({ w: i.width, h: i.height }); i.onerror = () => ok(null); i.src = src; });
}

const COT_TITULOS_PDF = { cotizar: 'A cotizar', cotizacion: 'Cotización', orden: 'Orden de compra' };

async function cotGenerarPDF(id, tipo) {
  const c = cotLista.find(x => x.id === id);
  if (!c) return false;
  if (!window.jspdf) { alert('No se cargó la librería de PDF. Revisa tu conexión y recarga la página.'); return false; }
  const doc = new window.jspdf.jsPDF({ unit: 'mm', format: 'a4' });
  const M = 18, ANCHO = 210 - M * 2, ALTO = 297;
  let y = M;

  const nuevaPaginaSi = h => { if (y + h > ALTO - M) { doc.addPage(); y = M; } };
  const escribir = (txt, tam, estilo = 'normal', interlineado = 1.25) => {
    doc.setFont('helvetica', estilo); doc.setFontSize(tam);
    doc.splitTextToSize(String(txt), ANCHO).forEach(l => {
      const h = tam * 0.3528 * interlineado;
      nuevaPaginaSi(h); doc.text(l, M, y + tam * 0.3528); y += h;
    });
  };

  const escribirPar = (a, b, tam) => {
    const h = tam * 0.3528 * 1.25; nuevaPaginaSi(h);
    doc.setFont('helvetica', 'bold'); doc.setFontSize(tam);
    const t = a + ': '; doc.text(t, M, y + tam * 0.3528);
    const w = doc.getTextWidth(t);
    doc.setFont('helvetica', 'normal'); doc.text(String(b), M + w, y + tam * 0.3528);
    y += h;
  };

  doc.setTextColor(30, 41, 59);
  escribir(COT_TITULOS_PDF[tipo] || 'Cotización', 22, 'bold'); y += 1;
  const fecha = new Date().toLocaleDateString('es-CO', { timeZone: 'America/Bogota', day: '2-digit', month: 'long', year: 'numeric' });
  doc.setTextColor(100, 116, 139); escribir(`ADMA Importaciones · ${fecha}`, 10); doc.setTextColor(30, 41, 59);
  y += 3; doc.setDrawColor(203, 213, 225); doc.line(M, y, M + ANCHO, y); y += 7;

  escribir('Cliente', 10, 'bold'); escribir(c.cliente, 13); y += 4;
  escribir('Descripción del producto', 10, 'bold'); escribir(c.descripcion, 12); y += 4;
  escribir('Cantidad', 10, 'bold'); escribir(cotNum(c.cantidad, 0) + ' unidades', 12); y += 4;

  const link = cotLinkSeguro(c.link);
  if (link) {
    escribir('Link', 10, 'bold');
    doc.setTextColor(37, 99, 235); doc.setFontSize(10);
    doc.splitTextToSize(link, ANCHO).forEach(l => { nuevaPaginaSi(5); doc.textWithLink(l, M, y + 3.5, { url: link }); y += 5; });
    doc.setTextColor(30, 41, 59); y += 4;
  }

  if (c.costos && (tipo === 'cotizacion' || tipo === 'orden')) {
    escribir('Costos', 10, 'bold');
    cotLineasCostos(c).forEach(g => {
      y += 2; escribir(g.titulo, 11, 'bold');
      g.lineas.forEach(([a, b]) => escribirPar(a, b, 12));
    });
    y += 4;
  }

  const usaReales = tipo !== 'cotizar' && (c.imagenesReales || []).length > 0;
  const imgs = usaReales ? c.imagenesReales : (c.imagenes || []);
  if (imgs.length) {
    escribir(usaReales ? 'Imágenes del producto cotizado' : 'Imágenes de referencia', 10, 'bold'); y += 2;
    const colW = (ANCHO - 6) / 2, maxH = 85;
    let col = 0, altoFila = 0;
    for (const src of imgs) {
      const d = await cotDimensionesImagen(src);
      if (!d) continue;
      const k = Math.min(colW / d.w, maxH / d.h);
      const w = d.w * k, h = d.h * k;
      if (col === 0) { nuevaPaginaSi(Math.max(h, 10)); altoFila = 0; }
      doc.addImage(src, 'JPEG', M + col * (colW + 6), y, w, h);
      altoFila = Math.max(altoFila, h);
      if (col === 1) { y += altoFila + 6; col = 0; } else col = 1;
    }
    if (col === 1) y += altoFila + 6;
  }

  const base = c.cliente.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'cliente';
  const prefijo = { cotizar: 'a-cotizar', cotizacion: 'cotizacion', orden: 'orden-de-compra' }[tipo] || 'cotizacion';
  doc.save(`${prefijo}-${base}.pdf`);
  return true;
}

/* Genera la orden de compra y, la primera vez, deja la tarjeta como Finalizada (verde). */
async function cotGenerarOrden(id) {
  const c = cotLista.find(x => x.id === id);
  if (!c) return;
  if (cotEstadoDe(c) === 'cotizada' &&
      !confirm(`Se generará la orden de compra de "${c.cliente}" y la cotización pasará a Finalizada. ¿Continuar?`)) return;
  const ok = await cotGenerarPDF(id, 'orden');
  if (!ok || cotEstadoDe(c) === 'finalizada') return;
  try { await cotRef().doc(id).update({ estado: 'finalizada', finalizadaEn: firebase.firestore.FieldValue.serverTimestamp() }); }
  catch (err) { console.error(err); alert('Se generó el PDF, pero no se pudo marcar la cotización como finalizada.'); }
}

/* ---------- Cajas (formulario de costos) ---------- */
function cotToggleVia() {
  const aerea = document.getElementById('cotVia').value === 'aerea';
  document.getElementById('cotFleteWrap').hidden = aerea;
  document.getElementById('cotFleteAereoWrap').hidden = !aerea;
  if (aerea) cotError('cotErrorCostos', '');
}

function cotActualizarCajas() {
  const c = cotLista.find(x => x.id === cotEditandoId);
  const el = document.getElementById('cotResumenCajas');
  const upc = parseInt(document.getElementById('cotUnidadesCaja').value, 10);
  if (!c) { el.textContent = ''; return; }
  if (!(upc > 0)) { el.textContent = `Cantidad solicitada: ${cotNum(c.cantidad, 0)}`; return; }
  const cajas = Math.ceil(c.cantidad / upc);
  el.textContent = `Cantidad solicitada: ${cotNum(c.cantidad, 0)} · Cajas: ${cotNum(cajas, 0)} · Total unidades cotizadas: ${cotNum(cajas * upc, 0)}`;
}

/* ---------- Excel "A cotizar" (una o varias cotizaciones en un solo archivo) ---------- */
function cotActualizarAcciones() {
  const n = cotSeleccion.size;
  document.getElementById('cotSelTexto').textContent = n ? `${n} seleccionada${n > 1 ? 's' : ''}` : '';
  const b = document.getElementById('cotBtnExcelSel');
  b.hidden = !n; b.textContent = `A cotizar - Generar Excel (${n})`;
  document.getElementById('cotBtnLimpiarSel').hidden = !n;
}

/* Una fila por cotización. No incluye el nombre del cliente: el archivo se envía al cotizador.
   Lo que aún no se conoce (PCS, valor, CBM, peso) queda en blanco para que lo complete. */
async function cotGenerarExcel(ids) {
  const lista = cotLista.filter(c => ids.includes(c.id));
  if (!lista.length) return;
  if (!window.ExcelJS) { alert('No se cargó la librería de Excel. Revisa tu conexión y recarga la página.'); return; }
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('A cotizar');
  ws.columns = [
    { header: 'Imagen de referencia', key: 'img', width: 26 },
    { header: 'Descripción', key: 'desc', width: 50 },
    { header: 'Cantidad', key: 'cant', width: 12 },
    { header: 'Link', key: 'link', width: 42 },
    { header: 'PCS * Caja', key: 'pcs', width: 14 },
    { header: 'Valor RMB * PCS', key: 'rmb', width: 18 },
    { header: 'CBM * Caja', key: 'cbm', width: 14 },
    { header: 'Peso * Caja', key: 'peso', width: 14 }
  ];
  const cab = ws.getRow(1);
  cab.font = { bold: true }; cab.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }; cab.height = 24;
  cab.eachCell(cell => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } }; });

  const MAX_W = 170, MAX_H = 130;
  for (const c of lista) {
    const k = c.costos || {};
    const link = cotLinkSeguro(c.link);
    const fila = ws.addRow({
      desc: c.descripcion, cant: c.cantidad, link: link ? { text: link, hyperlink: link } : null,
      pcs: k.unidadesPorCaja ?? null, rmb: k.valorUnidad ?? null, cbm: k.cbmCaja ?? null, peso: k.pesoCaja ?? null
    });
    fila.height = 105;
    fila.alignment = { vertical: 'middle', wrapText: true };
    const src = (c.imagenes || [])[0];
    const d = src ? await cotDimensionesImagen(src) : null;
    if (d) {
      const f = Math.min(MAX_W / d.w, MAX_H / d.h, 1);
      const idImg = wb.addImage({ base64: src, extension: 'jpeg' });
      ws.addImage(idImg, { tl: { col: 0.1, row: fila.number - 1 + 0.08 }, ext: { width: d.w * f, height: d.h * f } });
    }
  }
  const buf = await wb.xlsx.writeBuffer();
  const url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  const a = document.createElement('a');
  a.href = url; a.download = `a-cotizar-${typeof hoyBogota === 'function' ? hoyBogota() : 'excel'}.xlsx`;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

/* ---------- Mandar a módulo (Adma Company / Groupack) y pagos ----------
   Al mandar una cotización a su módulo se guarda: modulo ('adma-company' | 'groupack'), estado 'finalizada' y pagos: [].
   Cada pago: { id, concepto: 'producto'|'flete'|'todo', monto, montoProducto, montoFlete, fecha, nota, comprobante, por, registradoEn }
   El estado de pago (sin pagar / abonado / pagado) se calcula siempre a partir de los pagos y de los totales de la cotización. */
const COT_MODULO_DE_ORIGEN = { 'Adma Company': 'adma-company', 'Groupack': 'groupack' };
const cotNombreModulo = id => (APP_CONFIG.MODULOS.find(m => m.id === id) || {}).nombre || id;
const PG_ESTADOS = { sinpagar: 'Sin pagar', abonado: 'Abonado', pagado: 'Pagado' };
const PG_CONCEPTOS = { producto: 'Producto', flete: 'Flete', todo: 'Producto + flete' };
const pgAbiertas = new Set();
let pgPagoId = null;
const pgEsSuper = () => !!sesionActual && sesionActual.rol === APP_CONFIG.ROLES.SUPER;
const pgCop = v => '$ ' + cotNum(v, 0) + ' COP';
const pgSuma = (lista, campo) => lista.reduce((t, p) => t + (Number(p[campo]) || 0), 0);

async function cotMandarAModulo(id) {
  const c = cotLista.find(x => x.id === id);
  if (!c) return;
  const mod = COT_MODULO_DE_ORIGEN[c.origen];
  if (!mod) { alert('Esta cotización no tiene indicado si el cliente viene de Adma Company o Groupack, por eso no se puede mandar a un módulo.'); return; }
  const nombre = cotNombreModulo(mod);
  if (!confirm(`Se mandará la cotización de "${c.cliente}" al módulo ${nombre} y quedará Finalizada. ¿Continuar?`)) return;
  try {
    const datos = {
      estado: 'finalizada', modulo: mod,
      enviadaEn: firebase.firestore.FieldValue.serverTimestamp(),
      enviadaPor: sesionActual ? sesionActual.nombre : ''
    };
    if (!Array.isArray(c.pagos)) datos.pagos = [];
    if (!c.finalizadaEn) datos.finalizadaEn = firebase.firestore.FieldValue.serverTimestamp();
    await cotRef().doc(id).update(datos);
  } catch (err) { console.error(err); alert('No se pudo mandar la cotización al módulo.'); }
}

/* Totales en COP y estado de pago de una cotización */
function pgResumen(c) {
  const r = cotCalculos(c);
  if (!r) return null;
  const producto = cotOk(r.totalCop) ? r.totalCop : 0;
  const fleteTotal = cotOk(r.logistica360) ? r.logistica360 : 0;
  // Groupack: el flete a pagar es la "parte correspondiente a Adma" que se ingresa en la tarjeta (Precio total = producto + esa parte)
  const esGroupack = c.modulo === 'groupack';
  const parteAdma = esGroupack && cotOk(c.parteAdma) ? c.parteAdma : null;
  const flete = esGroupack ? (parteAdma ?? 0) : fleteTotal;
  const pagos = c.pagos || [];
  const pagadoP = pgSuma(pagos, 'montoProducto'), pagadoF = pgSuma(pagos, 'montoFlete');
  const pendP = Math.max(0, producto - pagadoP), pendF = Math.max(0, flete - pagadoF);
  const pagado = pagadoP + pagadoF, pendiente = pendP + pendF;
  const estado = pagado <= 0 ? 'sinpagar' : (pendiente <= 1 ? 'pagado' : 'abonado');
  return { unidades: r.unidades, producto, flete, fleteTotal, esGroupack, parteAdma, total: producto + flete, pagadoP, pagadoF, pendP, pendF, pagado, pendiente, estado };
}
const pgMaximo = (concepto, R) => concepto === 'producto' ? R.pendP : concepto === 'flete' ? R.pendF : R.pendP + R.pendF;

/* Reparte un pago "todo" entre producto y flete en proporción a lo que falta de cada uno */
function pgRepartir(concepto, monto, R) {
  if (concepto === 'producto') return { mp: monto, mf: 0 };
  if (concepto === 'flete') return { mp: 0, mf: monto };
  const pend = R.pendP + R.pendF;
  let mp = pend > 0 ? Math.min(R.pendP, monto * R.pendP / pend) : monto;
  let mf = monto - mp;
  if (mf > R.pendF) { mf = R.pendF; mp = monto - mf; }
  return { mp: Math.round(mp * 100) / 100, mf: Math.round(mf * 100) / 100 };
}

/* ---------- Módulos: estructura y tarjetas ---------- */
function pgConstruirModulo(mod) {
  const sec = document.getElementById(mod);
  if (!sec) return;
  sec.innerHTML = `
    <div class="cot-barra"><h2 class="cot-titulo">${cotEsc(cotNombreModulo(mod))}</h2>
      <div class="cot-filtros">
        <input type="search" data-pg-buscar placeholder="Buscar por cliente o estado de pago…" aria-label="Buscar en ${cotEsc(cotNombreModulo(mod))}" autocomplete="off">
        <select data-pg-estado aria-label="Filtrar por estado de pago"><option value="">Todos los pagos</option><option value="sinpagar">Sin pagar</option><option value="abonado">Abonado</option><option value="pagado">Pagado</option></select>
      </div>
      <p class="cot-conteo" data-pg-conteo aria-live="polite"></p></div>
    <div class="cot-grid" data-pg-grid aria-live="polite"><p class="lista-vacia">Cargando…</p></div>`;
  sec.addEventListener('input', e => { if (e.target.closest('[data-pg-buscar]')) pgPintar(mod); });
  sec.addEventListener('change', e => { if (e.target.closest('[data-pg-estado]')) pgPintar(mod); });
  sec.addEventListener('click', e => {
    const tog = e.target.closest('[data-pg-toggle]');
    const pag = e.target.closest('[data-pg-pagar]');
    const comp = e.target.closest('[data-pg-comp]');
    const bor = e.target.closest('[data-pg-borrar]');
    const parte = e.target.closest('[data-pg-parte-guardar]');
    if (parte) pgGuardarParteAdma(parte.dataset.pgParteGuardar, sec);
    else if (tog) { const id = tog.dataset.pgToggle; pgAbiertas.has(id) ? pgAbiertas.delete(id) : pgAbiertas.add(id); pgPintar(mod); }
    else if (pag) pgAbrirPago(pag.dataset.pgPagar);
    else if (comp) pgVerComprobante(comp.dataset.pgComp, comp.dataset.pid);
    else if (bor) pgQuitarPago(bor.dataset.pgBorrar, bor.dataset.pid);
  });
}

function pgTarjeta(c) {
  const R = pgResumen(c);
  if (!R) return '';
  const abierta = pgAbiertas.has(c.id);
  const fotos = c.imagenesReales || [];
  const pagos = (c.pagos || []).slice().sort((a, b) => (a.registradoEn || 0) - (b.registradoEn || 0));
  const lista = pagos.map(p => `<li class="pg-pago"><span>${cotEsc(p.fecha || '')} · ${cotEsc(PG_CONCEPTOS[p.concepto] || p.concepto)} · <strong>${pgCop(p.monto)}</strong>${p.nota ? ` · ${cotEsc(p.nota)}` : ''}${p.por ? ` · ${cotEsc(p.por)}` : ''}</span>
      ${p.comprobante ? `<button type="button" class="cot-enlace" data-pg-comp="${c.id}" data-pid="${cotEsc(p.id)}">Comprobante</button>` : ''}
      ${pgEsSuper() ? `<button type="button" class="cot-eliminar" data-pg-borrar="${c.id}" data-pid="${cotEsc(p.id)}" aria-label="Quitar este pago">Quitar</button>` : ''}</li>`).join('');
  return `<article class="cot-card cot-pg-${R.estado}${abierta ? ' abierta' : ''}">
    <button type="button" class="cot-resumen" data-pg-toggle="${c.id}" aria-expanded="${abierta}">
      ${fotos[0] ? `<img class="cot-foto" src="${fotos[0]}" alt="">` : '<span class="cot-foto" aria-hidden="true"></span>'}
      <span class="cot-resumen-txt"><strong>${cotEsc(c.cliente)}</strong><span class="cot-articulo">Total: ${pgCop(R.total)}</span></span>
      <span class="cot-estado pg-${R.estado}">${PG_ESTADOS[R.estado]}</span>
    </button>
    <div class="cot-detalle"${abierta ? '' : ' hidden'}>
      <p class="cot-dato"><strong>Cliente:</strong> ${cotEsc(c.cliente)}</p>
      ${fotos.length ? `<div class="cot-miniaturas">${fotos.map((s, i) => `<img src="${s}" alt="Foto ${i + 1} de ${cotEsc(c.cliente)}">`).join('')}</div>` : ''}
      <p class="cot-dato"><strong>Cantidad cotizada:</strong> ${cotNum(R.unidades, 0)}</p>
      <p class="cot-dato"><strong>Precio total del producto:</strong> ${pgCop(R.producto)}</p>
      <p class="cot-dato pg-flete-fila"><span><strong>Precio total del flete:</strong> ${pgCop(R.fleteTotal)}</span>
        ${R.esGroupack ? `<span class="pg-parte"><strong>Parte correspondiente a Adma:</strong>
          <input type="number" min="0" step="any" data-pg-parte-input="${c.id}" value="${R.parteAdma ?? ''}" placeholder="Valor en COP" aria-label="Parte correspondiente a Adma en COP">
          <button type="button" class="btn btn-chico" data-pg-parte-guardar="${c.id}">Guardar</button></span>` : ''}</p>
      <p class="cot-dato"><strong>Precio total:</strong> ${pgCop(R.total)}</p>
      <p class="cot-costos"><strong>PAGOS</strong><br>
        <strong>Producto:</strong> pagado ${pgCop(R.pagadoP)} · falta ${pgCop(R.pendP)}<br>
        <strong>Flete:</strong> pagado ${pgCop(R.pagadoF)} · falta ${pgCop(R.pendF)}<br>
        <strong>Total:</strong> pagado ${pgCop(R.pagado)} · falta ${pgCop(R.pendiente)}</p>
      ${lista ? `<ul class="pg-pagos">${lista}</ul>` : ''}
      <footer>${R.estado === 'pagado' ? '' : `<button type="button" class="btn btn-chico" data-pg-pagar="${c.id}">Registrar pago</button>`}</footer>
    </div>
  </article>`;
}

function pgPintar(mod) {
  const sec = document.getElementById(mod);
  const grid = sec && sec.querySelector('[data-pg-grid]');
  if (!grid) return;
  const conteo = sec.querySelector('[data-pg-conteo]');
  const todas = cotLista.filter(c => c.modulo === mod && cotEstadoDe(c) === 'finalizada' && c.costos);
  if (!todas.length) {
    conteo.textContent = '';
    grid.innerHTML = `<p class="lista-vacia">Aún no hay cotizaciones en este módulo. Se agregan desde Cotizaciones con el botón "Mandar a módulo".</p>`;
    return;
  }
  const q = cotNorm(sec.querySelector('[data-pg-buscar]').value).split(/\s+/).filter(Boolean);
  const est = sec.querySelector('[data-pg-estado]').value;
  const lista = todas.filter(c => {
    const R = pgResumen(c);
    if (!R || (est && R.estado !== est)) return false;
    const pajar = cotNorm(c.cliente + ' ' + PG_ESTADOS[R.estado]);
    return q.every(t => pajar.includes(t));
  });
  conteo.textContent = lista.length === todas.length ? `${todas.length} cotizaciones` : `${lista.length} de ${todas.length} cotizaciones`;
  grid.innerHTML = lista.length ? lista.map(pgTarjeta).join('') : '<p class="lista-vacia">Ninguna cotización coincide con la búsqueda.</p>';
}
const pgPintarTodos = () => Object.values(COT_MODULO_DE_ORIGEN).forEach(pgPintar);

/* ---------- Groupack: parte correspondiente a Adma ---------- */
async function pgGuardarParteAdma(id, sec) {
  const input = sec.querySelector(`[data-pg-parte-input="${id}"]`);
  const v = parseFloat(input && input.value);
  if (!(v >= 0)) { alert('Escribe un valor válido (en COP) para la parte correspondiente a Adma.'); return; }
  try { await cotRef().doc(id).update({ parteAdma: v }); }
  catch (err) { console.error(err); alert('No se pudo guardar el valor.'); }
}

/* ---------- Pagos ---------- */
function pgActualizarResumen() {
  const c = cotLista.find(x => x.id === pgPagoId);
  const R = c && pgResumen(c);
  document.getElementById('cotPagoResumen').textContent = R
    ? `Falta por pagar · Producto: ${pgCop(R.pendP)} · Flete: ${pgCop(R.pendF)} · Total: ${pgCop(R.pendP + R.pendF)}` : '';
}

function pgAbrirPago(id) {
  const c = cotLista.find(x => x.id === id);
  if (!c) return;
  pgPagoId = id;
  document.getElementById('cotFormPago').reset();
  cotError('cotErrorPago', '');
  document.getElementById('cotPagoTitulo').textContent = 'Registrar pago · ' + c.cliente;
  document.getElementById('cotPagoFecha').value = typeof hoyBogota === 'function' ? hoyBogota() : '';
  pgActualizarResumen();
  cotAbrir('cotModalPago');
  document.getElementById('cotPagoMonto').focus();
}

async function pgGuardarPago(e) {
  e.preventDefault();
  const c = cotLista.find(x => x.id === pgPagoId);
  const R = c && pgResumen(c);
  if (!R) return;
  const concepto = document.getElementById('cotPagoConcepto').value;
  const monto = parseFloat(document.getElementById('cotPagoMonto').value);
  const max = pgMaximo(concepto, R);
  if (!(monto > 0)) return cotError('cotErrorPago', 'Escribe un valor mayor que cero.');
  if (max <= 1) return cotError('cotErrorPago', `Ya no falta nada por pagar en "${PG_CONCEPTOS[concepto]}".`);
  if (monto > max + 1) return cotError('cotErrorPago', `El valor supera lo que falta por pagar (${pgCop(max)}).`);

  let comprobante = '';
  const archivo = document.getElementById('cotPagoArchivo').files[0];
  if (!archivo) return cotError('cotErrorPago', 'Sube el comprobante de pago para poder registrar el pago.');
  {
    if (!archivo.type.startsWith('image/')) return cotError('cotErrorPago', 'El comprobante debe ser una imagen.');
    try { comprobante = await cotComprimir(archivo, 800, 0.6); }
    catch (err) { console.error(err); return cotError('cotErrorPago', 'No se pudo leer la imagen del comprobante.'); }
    const usado = (c.imagenes || []).join('').length + (c.imagenesReales || []).join('').length + (c.pagos || []).reduce((t, p) => t + (p.comprobante || '').length, 0);
    if (usado + comprobante.length > 950000) return cotError('cotErrorPago', 'Esta cotización ya no tiene espacio para otro comprobante (límite de tamaño). Avisa al administrador.');
  }
  const { mp, mf } = pgRepartir(concepto, monto, R);
  const pago = {
    id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
    concepto, monto, montoProducto: mp, montoFlete: mf,
    fecha: document.getElementById('cotPagoFecha').value || (typeof hoyBogota === 'function' ? hoyBogota() : ''),
    nota: document.getElementById('cotPagoNota').value.trim(),
    comprobante, por: sesionActual ? sesionActual.nombre : '', registradoEn: Date.now()
  };
  const btn = document.getElementById('cotBtnGuardarPago');
  btn.disabled = true;
  try {
    await cotRef().doc(pgPagoId).update({ pagos: firebase.firestore.FieldValue.arrayUnion(pago) });
    cotCerrarTodo();
  } catch (err) {
    console.error('Error al guardar el pago:', err);
    cotError('cotErrorPago', 'No se pudo guardar el pago. Inténtalo de nuevo.');
  } finally { btn.disabled = false; }
}

async function pgQuitarPago(id, pid) {
  if (!pgEsSuper()) { alert('Solo el Super usuario puede quitar pagos registrados.'); return; }
  const c = cotLista.find(x => x.id === id);
  const p = c && (c.pagos || []).find(x => x.id === pid);
  if (!p || !confirm(`¿Quitar el pago de ${pgCop(p.monto)} del ${p.fecha || 'sin fecha'}?`)) return;
  try { await cotRef().doc(id).update({ pagos: firebase.firestore.FieldValue.arrayRemove(p) }); }
  catch (err) { console.error(err); alert('No se pudo quitar el pago.'); }
}

function pgVerComprobante(id, pid) {
  const c = cotLista.find(x => x.id === id);
  const p = c && (c.pagos || []).find(x => x.id === pid);
  if (!p || !/^data:image\//.test(p.comprobante || '')) return;
  const w = window.open('', '_blank');
  if (!w) { alert('El navegador bloqueó la ventana. Permite las ventanas emergentes e inténtalo de nuevo.'); return; }
  w.document.write(`<title>Comprobante</title><body style="margin:0;background:#222"><img src="${p.comprobante}" style="max-width:100%;display:block;margin:auto">`);
  w.document.close();
}

/* ---------- Escucha en tiempo real ---------- */
function initEscuchaCotizaciones() {
  if (!document.getElementById('cotGrid')) return;
  detenerEscuchaCotizaciones();
  cotDesuscribir = cotRef().orderBy('creadoEn', 'desc').onSnapshot(snap => {
    cotLista = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    // Solo se pueden seleccionar las pendientes: si una cambia de estado o se borra, sale de la selección
    [...cotSeleccion].forEach(id => { if (!cotLista.some(c => c.id === id && cotEstadoDe(c) === 'pendiente')) cotSeleccion.delete(id); });
    cotPintar();
    pgPintarTodos();
  }, err => {
    console.error('Error al leer cotizaciones:', err);
    document.getElementById('cotGrid').innerHTML = '<p class="lista-vacia">No se pudieron cargar las cotizaciones. Revisa las reglas de Firestore.</p>';
    pgPintarTodos();
  });
}
function detenerEscuchaCotizaciones() {
  if (cotDesuscribir) { cotDesuscribir(); cotDesuscribir = null; }
  cotLista = [];
  cotSeleccion.clear();
  pgPintarTodos();
}

/* ---------- Inicialización (se registra después de script.js) ---------- */
document.addEventListener('DOMContentLoaded', () => {
  if (!cotConstruirUI()) return;
  document.getElementById('cotBtnNueva').addEventListener('click', cotAbrirNueva);
  document.getElementById('cotFormNueva').addEventListener('submit', cotGuardarNueva);
  document.getElementById('cotFormCostos').addEventListener('submit', cotGuardarCostos);
  document.getElementById('cotArchivos').addEventListener('change', cotManejarArchivos);

  document.getElementById('cotPrevias').addEventListener('click', e => {
    const b = e.target.closest('[data-cot-quitar]');
    if (b) { cotImagenesForm.splice(Number(b.dataset.cotQuitar), 1); cotPintarPrevias(); }
  });
  document.querySelectorAll('[data-cot-cerrar]').forEach(el => el.addEventListener('click', cotCerrarTodo));
  document.addEventListener('keydown', e => { if (e.key === 'Escape') cotCerrarTodo(); });

  document.getElementById('cotUnidadesCaja').addEventListener('input', cotActualizarCajas);
  document.getElementById('cotVia').addEventListener('change', cotToggleVia);
  document.getElementById('cotOrigen').addEventListener('change', cotToggleComercial);
  document.getElementById('cotArchivosReales').addEventListener('change', cotManejarArchivosReales);
  document.getElementById('cotPreviasReales').addEventListener('click', e => {
    const b = e.target.closest('[data-cot-quitar-real]');
    if (b) { cotImagenesReales.splice(Number(b.dataset.cotQuitarReal), 1); cotPintarPreviasReales(); }
  });
  document.getElementById('cotAumentoTipo').addEventListener('change', e => {
    document.getElementById('cotAumento').placeholder = e.target.value === 'porcentaje' ? '% sobre el total' : 'RMB por unidad';
  });
  ['cotBuscar', 'cotFiltroEstado', 'cotFiltroOrigen'].forEach(id =>
    document.getElementById(id).addEventListener(id === 'cotBuscar' ? 'input' : 'change', cotPintar));

  document.getElementById('cotGrid').addEventListener('change', e => {
    const cb = e.target.closest('[data-cot-sel]');
    if (!cb) return;
    cb.checked ? cotSeleccion.add(cb.dataset.cotSel) : cotSeleccion.delete(cb.dataset.cotSel);
    cotActualizarAcciones();
  });
  document.getElementById('cotBtnSelVisibles').addEventListener('click', () => { cotFiltradas().filter(c => cotEstadoDe(c) === 'pendiente').forEach(c => cotSeleccion.add(c.id)); cotPintar(); });
  document.getElementById('cotBtnLimpiarSel').addEventListener('click', () => { cotSeleccion.clear(); cotPintar(); });
  document.getElementById('cotBtnExcelSel').addEventListener('click', () => cotGenerarExcel([...cotSeleccion]));

  document.getElementById('cotGrid').addEventListener('click', e => {
    const tog = e.target.closest('[data-cot-toggle]');
    const xls = e.target.closest('[data-cot-excel]');
    const pdf = e.target.closest('[data-cot-pdf]');
    const ord = e.target.closest('[data-cot-orden]');
    const mdl = e.target.closest('[data-cot-modulo]');
    const cont = e.target.closest('[data-cot-continuar]');
    const del = e.target.closest('[data-cot-eliminar]');
    if (tog) cotAlternar(tog.dataset.cotToggle);
    else if (xls) cotGenerarExcel([xls.dataset.cotExcel]);
    else if (pdf) cotGenerarPDF(pdf.dataset.cotPdf, pdf.dataset.tipo);
    else if (ord) cotGenerarOrden(ord.dataset.cotOrden);
    else if (mdl) cotMandarAModulo(mdl.dataset.cotModulo);
    else if (cont) cotAbrirCostos(cont.dataset.cotContinuar);
    else if (del) cotEliminar(del.dataset.cotEliminar);
  });

  /* Módulos Adma Company / Groupack y formulario de pagos */
  Object.values(COT_MODULO_DE_ORIGEN).forEach(pgConstruirModulo);
  document.getElementById('cotFormPago').addEventListener('submit', pgGuardarPago);
  document.getElementById('cotPagoConcepto').addEventListener('change', () => { document.getElementById('cotPagoMonto').value = ''; });
  document.getElementById('cotPagoSaldo').addEventListener('click', () => {
    const c = cotLista.find(x => x.id === pgPagoId), R = c && pgResumen(c);
    if (R) document.getElementById('cotPagoMonto').value = Math.round(pgMaximo(document.getElementById('cotPagoConcepto').value, R));
  });
});
