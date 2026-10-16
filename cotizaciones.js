/* ===== Módulo Cotizaciones =====
   Colección Firestore: cotizaciones
   {
     cliente, descripcion, cantidad (número; si es un rango, el máximo), cantidadMin (solo si es rango), link, imagenes: [dataURL JPEG comprimido],
     estado: 'pendiente' | 'cotizada' | 'finalizada',
     observaciones: texto con tallas/cantidades (se pide al mandar a módulo; va en NOTES de la orden de compra),
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
let cotDetalleId = null;              // id de la cotización abierta en la ventana de detalle
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
        <button type="button" class="cot-enlace" id="cotBtnSelVisibles">Seleccionar visibles</button>
        <span id="cotSelTexto" class="cot-conteo" style="flex-basis:auto"></span>
        <button type="button" class="btn btn-chico" id="cotBtnExcelSel" hidden></button>
        <button type="button" class="btn btn-chico" id="cotBtnPdfSel" hidden></button>
        <button type="button" class="btn btn-chico" id="cotBtnOrdenSel" hidden></button>
        <button type="button" class="btn btn-chico secundario" id="cotBtnOrdenPrueba" hidden title="Genera el PDF con datos ficticios, sin tocar consecutivos ni guardar nada"></button>
        <button type="button" class="cot-enlace" id="cotBtnLimpiarSel" hidden>Quitar selección</button>
      </div></div>
    <div id="cotGrid" class="cot-tabla-wrap" aria-live="polite"><p class="lista-vacia">Cargando cotizaciones…</p></div>
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
      <label class="login-campo"><span>Cantidad</span><input type="text" id="cotCantidad" inputmode="numeric" autocomplete="off" placeholder="Ej: 100 o un rango 50-100" required></label>
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

  <div id="cotModalDetalle" class="modal" hidden>
    <div class="modal-fondo" data-cot-cerrar></div>
    <div class="modal-tarjeta cot-modal cot-modal-detalle" role="dialog" aria-modal="true" aria-labelledby="cotDetalleTitulo">
      <div class="modal-cabecera"><h2 id="cotDetalleTitulo">Cotización</h2>
        <button type="button" class="modal-cerrar" data-cot-cerrar aria-label="Cerrar">✕</button></div>
      <div id="cotDetalleCuerpo" class="cot-detalle-tabla"></div>
    </div>
  </div>

  <div id="cotModalObs" class="modal" hidden>
    <div class="modal-fondo" data-cot-cerrar></div>
    <form id="cotFormObs" class="modal-tarjeta cot-modal" novalidate>
      <div class="modal-cabecera"><h2 id="cotObsTitulo">Observaciones</h2>
        <button type="button" class="modal-cerrar" data-cot-cerrar aria-label="Cerrar">✕</button></div>
      <div id="cotErrorObs" class="login-error" role="alert" hidden></div>
      <p class="cot-dato cot-calc" id="cotObsResumen"></p>
      <label class="login-campo"><span>Observaciones para la orden de compra (tallas, cantidades por talla, colores…)</span>
        <textarea id="cotObsTexto" rows="7" maxlength="${COT_OBS_MAX}" placeholder="Ej:&#10;Talla S: 20&#10;Talla M: 40&#10;Talla L: 30"></textarea></label>
      <p class="cot-dato cot-calc">Van en la columna NOTES de la orden de compra. Puedes escribir varias líneas (máx. ${COT_OBS_MAX} caracteres).</p>
      <div class="modal-acciones">
        <button type="button" class="btn secundario" data-cot-cerrar>Cancelar</button>
        <button type="submit" class="btn" id="cotBtnGuardarObs">Guardar</button>
      </div>
    </form>
  </div>

  <div id="ocModalDetalle" class="modal" hidden>
    <div class="modal-fondo" data-cot-cerrar></div>
    <div class="modal-tarjeta cot-modal cot-modal-detalle" role="dialog" aria-modal="true" aria-labelledby="ocDetalleTitulo">
      <div class="modal-cabecera"><h2 id="ocDetalleTitulo">Orden de compra</h2>
        <button type="button" class="modal-cerrar" data-cot-cerrar aria-label="Cerrar">✕</button></div>
      <div id="ocDetalleCuerpo" class="cot-detalle-tabla"></div>
    </div>
  </div>

  <div id="pgModalDetalle" class="modal" hidden>
    <div class="modal-fondo" data-cot-cerrar></div>
    <div class="modal-tarjeta cot-modal cot-modal-detalle" role="dialog" aria-modal="true" aria-labelledby="pgDetalleTitulo">
      <div class="modal-cabecera"><h2 id="pgDetalleTitulo">Cotización</h2>
        <button type="button" class="modal-cerrar" data-cot-cerrar aria-label="Cerrar">✕</button></div>
      <div id="pgDetalleCuerpo" class="cot-detalle-tabla"></div>
    </div>
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
/* Cantidad: acepta "100" o un rango "50-100" (también con – o —). Para los cálculos se usa el número mayor. */
function cotLeerCantidad(texto) {
  const m = /^\s*(\d+)\s*(?:[-–—]\s*(\d+))?\s*$/.exec(String(texto || '').replace(/[.,]/g, ''));
  if (!m) return null;
  const min = parseInt(m[1], 10), max = m[2] != null ? parseInt(m[2], 10) : min;
  return min > 0 && max > 0 ? { min, max } : null;
}
const cotEsRango = c => cotOk(c.cantidadMin) && c.cantidadMin > 0 && c.cantidadMin < c.cantidad;
const cotCantidadTxt = c => cotEsRango(c) ? `${cotNum(c.cantidadMin, 0)} - ${cotNum(c.cantidad, 0)}` : cotNum(c.cantidad, 0);
const cotAbrir = id => { document.getElementById(id).hidden = false; };
function cotCerrarModal(m) {
  if (!m) return;
  m.hidden = true;
  if (m.id === 'ocModalDetalle') ocDetalleBooth = null;
  if (m.id === 'cotModalObs') cotObsId = null;
  if (m.id === 'cotModalDetalle') cotDetalleId = null;
  if (m.id === 'pgModalDetalle') pgDetalleId = null;
}
function cotCerrarTodo() {
  document.querySelectorAll('#cotModalNueva, #cotModalCostos, #cotModalPago, #cotModalDetalle, #pgModalDetalle, #ocModalDetalle, #cotModalObs').forEach(m => { m.hidden = true; });
  cotDetalleId = null; pgDetalleId = null; ocDetalleBooth = null; cotObsId = null;
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
  const cant = cotLeerCantidad(document.getElementById('cotCantidad').value);
  const link = document.getElementById('cotLink').value.trim();

  const origen = document.getElementById('cotOrigen').value;
  const comercial = origen === 'Adma Company' ? document.getElementById('cotComercial').value.trim() : '';
  if (!cliente || !descripcion) return cotError('cotErrorNueva', 'Escribe el nombre del cliente y la descripción.');
  if (!origen) return cotError('cotErrorNueva', 'Indica de dónde proviene el cliente.');
  if (origen === 'Adma Company' && !comercial) return cotError('cotErrorNueva', 'Escribe el nombre del comercial que proporcionó el cliente.');
  if (!cant) return cotError('cotErrorNueva', 'Escribe la cantidad como un número (100) o como un rango (50-100), mayor que cero.');
  if (cant.min > cant.max) return cotError('cotErrorNueva', 'En el rango, el primer número debe ser menor que el segundo (por ejemplo 50-100).');
  const cantidad = cant.max;
  if (link && !cotLinkSeguro(link)) return cotError('cotErrorNueva', 'El link debe empezar por http:// o https://');
  if (cotImagenesForm.join('').length > COT_MAX_CHARS_TOTAL) return cotError('cotErrorNueva', 'Las imágenes pesan demasiado. Quita alguna.');

  const btn = document.getElementById('cotBtnGuardar');
  btn.disabled = true;
  try {
    await cotRef().add({
      cliente, origen, comercial, descripcion, cantidad, cantidadMin: cant.min < cant.max ? cant.min : null, link, imagenes: cotImagenesForm,
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
  if (c && (cotTieneOrden(c) || (c.pagos || []).length)) { alert('Esta cotización ya tiene orden de compra o pagos registrados y no se puede eliminar.'); return; }
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
      ['Total para el cliente con el aumento y la comisión de Adma', r.totalRmb, 2, '', Y],
      ['Comisión de Adma' + (pct ? ` (${pct.trim()})` : ''), r.comisionTotal, 2, '', Y]
      
    ]),
    grupo(r.viaAerea ? 'FLETE AÉREO SUJETO A IMPUESTOS' : 'FLETE MARÍTIMO', [
      ['CBM', r.cbmTotal, 4, ' m³'],
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
    const pajar = cotNorm([c.cliente, c.descripcion, c.origen, c.comercial, COT_ESTADOS[e], k.broker, c.link, c.creadoPor, c.cantidad, c.cantidadMin].join(' '));
    return q.every(t => pajar.includes(t));
  });
}

/* ---------- Tabla (al hacer clic en una fila se despliega el detalle) ---------- */
const COT_ICONOS = {
  excel:  '<rect x="4" y="4" width="16" height="16" rx="2"/><path d="M4 10h16M4 15h16M10 4v16"/>',
  pdf:    '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/>',
  enviar: '<path d="M5 12h14M13 6l6 6-6 6"/>',
  editar: '<path d="M4 20h4L19 9a2.8 2.8 0 0 0-4-4L4 16z"/><path d="M13.5 6.5l4 4"/>',
  borrar: '<path d="M4 7h16M10 11v6M14 11v6M6 7l1 12a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-12M9 7V4h6v3"/>'
};
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
  const ico = n => `<svg class="cot-svg" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${COT_ICONOS[n]}</svg>`;
  const filas = lista.map(c => {
    const est = cotEstadoDe(c);
    const puedeModulo = est === 'cotizada' || (est === 'finalizada' && !c.modulo);
    const etiquetaEditar = est === 'pendiente' ? 'Continuar: llenar los costos' : 'Editar costos';

    const acciones = [
      `<button type="button" class="cot-btn-tabla" data-cot-excel="${c.id}" title="A cotizar - Generar Excel">${ico('excel')}<span>A cotizar</span></button>`,
      c.costos ? `<button type="button" class="cot-btn-tabla" data-cot-pdf="${c.id}" data-tipo="cotizacion" title="${c.origen === 'Groupack' ? 'Cotización - Generar Excel (formato Groupack)' : 'Cotización - Generar PDF'}">${ico(c.origen === 'Groupack' ? 'excel' : 'pdf')}<span>${c.origen === 'Groupack' ? 'Cotización Excel' : 'Cotización PDF'}</span></button>` : '<span aria-hidden="true"></span>',
      puedeModulo ? `<button type="button" class="cot-btn-tabla cot-btn-modulo" data-cot-modulo="${c.id}" title="Mandar a módulo">${ico('enviar')}<span>Mandar a módulo</span></button>` : '<span aria-hidden="true"></span>',
      est !== 'finalizada' ? `<button type="button" class="cot-btn-tabla cot-btn-ico" data-cot-continuar="${c.id}" title="${etiquetaEditar}" aria-label="${etiquetaEditar} de ${cotEsc(c.cliente)}">${ico('editar')}</button>` : '<span aria-hidden="true"></span>',
      `<button type="button" class="cot-btn-tabla cot-btn-ico cot-btn-peligro" data-cot-eliminar="${c.id}" title="Eliminar" aria-label="Eliminar cotización de ${cotEsc(c.cliente)}">${ico('borrar')}</button>`
    ].join('');

    const fila = `<tr class="cot-fila" data-cot-toggle="${c.id}" tabindex="0" aria-haspopup="dialog" title="Ver detalle">
      <td class="cot-col-sel"><input type="checkbox" class="cot-sel" data-cot-sel="${c.id}"${cotSeleccion.has(c.id) ? ' checked' : ''} aria-label="Seleccionar la cotización de ${cotEsc(c.cliente)}"></td>
      <td class="cot-col-cliente"><strong>${cotEsc(c.cliente)}</strong></td>
      <td class="cot-col-prod"><span class="cot-prod-txt">${cotEsc(c.descripcion)}</span></td>
      <td class="cot-col-estado"><span class="cot-estado cot-e-${est}">${COT_ESTADOS[est]}</span>${cotTieneOrden(c) ? `<span class="cot-oc" title="Orden de compra generada">OC ${cotEsc(c.orden.booth)}</span>` : ''}</td>
      <td class="cot-col-acc"><div class="cot-acciones-fila">${acciones}</div></td>
    </tr>`;

    return fila;
  }).join('');

  grid.innerHTML = `<div class="cot-tabla-card"><table class="cot-tabla">
    <thead><tr><th class="cot-col-sel"><span class="sr-only">Seleccionar</span></th><th>Cliente</th><th>Producto</th><th>Estado</th><th class="cot-col-acc">Acciones</th></tr></thead>
    <tbody>${filas}</tbody></table></div>`;
}

function cotAlternar(id) {
  cotDetalleId = id;
  cotPintarDetalle();
  if (cotDetalleId) cotAbrir('cotModalDetalle');
}

/* Ventana con toda la información de una cotización */
function cotPintarDetalle() {
  const c = cotLista.find(x => x.id === cotDetalleId);
  const modal = document.getElementById('cotModalDetalle');
  if (!c) { cotDetalleId = null; if (modal) modal.hidden = true; return; }
  const est = cotEstadoDe(c);
  const link = cotLinkSeguro(c.link);
  const costos = c.costos ? cotLineasCostos(c).map(g => `<p class="cot-costos"><strong>${cotEsc(g.titulo)}</strong><br>${g.lineas.map(([a, b]) => `<strong>${cotEsc(a)}:</strong> ${cotEsc(b)}`).join('<br>')}</p>`).join('') : '';
  document.getElementById('cotDetalleTitulo').textContent = c.cliente;
  document.getElementById('cotDetalleCuerpo').innerHTML = `
    <p><span class="cot-estado cot-e-${est}">${COT_ESTADOS[est]}</span></p>
    <p>${cotEsc(c.descripcion)}</p>
    ${c.origen ? `<p class="cot-dato"><strong>Cliente de:</strong> ${cotEsc(c.origen)}${c.comercial ? ` · <strong>Comercial:</strong> ${cotEsc(c.comercial)}` : ''}</p>` : ''}
    <p class="cot-dato"><strong>Cantidad:</strong> ${cotCantidadTxt(c)}${(c.costos && c.costos.broker) ? ` · <strong>Broker:</strong> ${cotEsc(c.costos.broker)}` : ''}</p>
    ${c.modulo ? `<p class="cot-dato"><strong>Enviada a:</strong> ${cotEsc(cotNombreModulo(c.modulo))}</p>` : ''}
    ${est === 'finalizada' || c.observaciones ? cotObsHtml(c, 'data-cot-obs') : ''}
    ${link ? `<p class="cot-dato"><a href="${cotEsc(link)}" target="_blank" rel="noopener noreferrer">Ver link del producto</a></p>` : ''}
    ${(c.imagenes || []).length ? `<div class="cot-miniaturas">${c.imagenes.map((s, i) => `<img src="${cotEsc(s)}" alt="Referencia ${i + 1} de ${cotEsc(c.cliente)}">`).join('')}</div>` : ''}
    ${costos}
    ${cotTieneOrden(c) ? `<p class="cot-dato"><strong>Orden de compra:</strong> Booth ${cotEsc(c.orden.booth)} · Item No ${cotEsc(c.orden.itemNo)} · Shipping Mark ${cotEsc(c.orden.marca)}</p>` : ''}
    ${est === 'finalizada' ? `<footer><button type="button" class="cot-btn-tabla" data-cot-orden="${c.id}"><svg class="cot-svg" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${COT_ICONOS.pdf}</svg><span>${cotTieneOrden(c) ? 'Descargar orden de compra · Booth ' + c.orden.booth : 'Generar orden de compra'}</span></button>${c.costos && pgEsSuper() ? `<button type="button" class="cot-btn-tabla" data-cot-orden-prueba="${c.id}" title="PDF con datos ficticios: no cambia consecutivos ni guarda nada"><span>Orden de PRUEBA</span></button>` : ''}</footer>` : ''}`;
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
  if (tipo === 'cotizacion') return cotGenerarPDFCotizacion(c);
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
  escribir('Cantidad', 10, 'bold'); escribir(cotCantidadTxt(c) + ' unidades', 12); y += 4;

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
/* Logo de la empresa para el PDF (configuracion/empresa). jsPDF no acepta SVG: se pasa por un canvas a PNG. */
async function cotLogoParaPDF() {
  try {
    const d = await db.collection(APP_CONFIG.COLECCIONES.configuracion).doc('empresa').get();
    const src = d.exists ? d.data().logoBase64 : null;
    if (!src) return null;
    const img = await new Promise((ok, no) => { const i = new Image(); i.onload = () => ok(i); i.onerror = no; i.src = src; });
    let w = img.naturalWidth || 300, h = img.naturalHeight || 150;
    const esc = Math.min(1, 600 / Math.max(w, h));
    w = Math.round(w * esc); h = Math.round(h * esc);
    const cv = document.createElement('canvas');
    cv.width = w; cv.height = h;
    cv.getContext('2d').drawImage(img, 0, 0, w, h);
    return { data: cv.toDataURL('image/png'), w, h };
  } catch (err) { console.warn('No se pudo cargar el logo para el PDF:', err); return null; }
}

/* PDF de Cotización con el formato de ADMA: diseño de marca (amarillo/negro) pensado para la venta.
   Recibe una cotización o un arreglo de cotizaciones del MISMO cliente (un solo PDF con todos los productos). */
async function cotGenerarPDFCotizacion(entrada) {
  const todas = (Array.isArray(entrada) ? entrada : [entrada]).filter(c => c && c.costos);
  if (!todas.length) { alert('Estas cotizaciones aún no tienen costos guardados.'); return false; }
  /* Clientes de Groupack: la cotización sale en Excel. Una sola cotización: formato carta de Groupack (2 hojas). Varias: tabla con todos los productos en un solo archivo.
     Si la selección mezcla orígenes, los de Groupack van en Excel y el resto en PDF (dos descargas). */
  const grupales = todas.filter(c => c.origen === 'Groupack');
  if (grupales.length) {
    const okXls = grupales.length === 1 ? await cotGenerarExcelGrupalUna(grupales[0]) : await cotGenerarExcelGrupal(grupales);
    if (grupales.length === todas.length) return okXls;
  }
  const lista = grupales.length ? todas.filter(c => c.origen !== 'Groupack') : todas;
  if (!window.jspdf) { alert('No se cargó la librería de PDF. Revisa tu conexión y recarga la página.'); return false; }
  const N = lista.length, multi = N > 1;
  const rs = lista.map(c => cotCalculos(c) || {});
  const doc = new window.jspdf.jsPDF({ unit: 'mm', format: 'a4' });
  const M = 20, ANCHO = 210 - M * 2, ALTO = 297, PIE = 24;
  const LIMITE = ALTO - PIE - 6;           // espacio reservado para el pie
  const AMARILLO = [255, 236, 0], OSCURO = [24, 24, 27], GRIS = [100, 100, 110], CLARO = [244, 244, 245], BORDE = [225, 225, 228];
  const CY = M + 42;   // el contenido empieza debajo del encabezado, que se repite en todas las hojas
  const cop = v => v == null ? '—' : '$ ' + cotNum(v, 2) + ' COP';
  const recortar = (t, max) => { t = String(t || '').replace(/\s+/g, ' ').trim(); return t.length > max ? t.slice(0, max - 1) + '…' : t; };
  let y = CY, RH = 8;   // alto de cada fila (baja a 7 o 6 si hace falta para que un solo producto quepa en una hoja)

  const nuevaPaginaSi = h => { if (y + h > LIMITE) { doc.addPage(); y = CY; } };
  const seccion = t => {
    nuevaPaginaSi(14);
    doc.setFillColor(...AMARILLO); doc.rect(M, y, 2.2, 5.5, 'F');
    doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(...OSCURO);
    doc.text(t.toUpperCase(), M + 5, y + 4.2);
    y += 9;
  };
  const fila = (a, b) => {
    nuevaPaginaSi(RH);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(10.5); doc.setTextColor(...GRIS);
    doc.text(a, M + 1, y + RH - 3);
    doc.setFont('helvetica', 'bold'); doc.setTextColor(...OSCURO);
    doc.text(String(b), M + ANCHO - 1, y + RH - 3, { align: 'right' });
    y += RH;
    doc.setDrawColor(...BORDE); doc.setLineWidth(0.2); doc.line(M, y, M + ANCHO, y);
  };

  /* ---- Encabezado (se dibuja al final en TODAS las hojas): logo, título, fecha y línea amarilla ---- */
  const logo = await cotLogoParaPDF();
  const fecha = new Date().toLocaleDateString('es-CO', { timeZone: 'America/Bogota', day: '2-digit', month: 'long', year: 'numeric' });
  const dibujarEncabezado = () => {
    if (logo) {
      const k = Math.min(40 / logo.w, 30 / logo.h);
      doc.addImage(logo.data, 'PNG', M, M, logo.w * k, logo.h * k);
    }
    doc.setTextColor(...OSCURO); doc.setFont('helvetica', 'bold'); doc.setFontSize(26);
    doc.text('COTIZACIÓN', 210 - M, M + 11, { align: 'right' });
    doc.setFont('helvetica', 'normal'); doc.setFontSize(12); doc.setTextColor(...GRIS);
    doc.text('Adma Importaciones', 210 - M, M + 18, { align: 'right' });
    doc.setFontSize(10);
    doc.text(fecha, 210 - M, M + 24, { align: 'right' });
    doc.setFillColor(...AMARILLO); doc.rect(M, M + 33, ANCHO, 1.8, 'F');
  };

  /* ---- Cliente (con el producto si es uno solo; con el número de productos si son varios) ---- */
  doc.setFont('helvetica', 'normal'); doc.setFontSize(11);
  const desc1 = multi ? [] : doc.splitTextToSize(String(lista[0].descripcion || ''), ANCHO - 12);
  const hCliente = multi ? 20 : 30 + desc1.length * 5;
  nuevaPaginaSi(hCliente);
  doc.setFillColor(...CLARO); doc.roundedRect(M, y, ANCHO, hCliente, 3, 3, 'F');
  doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.setTextColor(...GRIS);
  doc.text('CLIENTE', M + 6, y + 7);
  doc.setFontSize(14); doc.setTextColor(...OSCURO);
  doc.text(String(lista[0].cliente || ''), M + 6, y + 13);
  if (multi) {
    doc.setFontSize(11);
    doc.text(`${N} productos cotizados`, M + ANCHO - 6, y + 13, { align: 'right' });
  } else {
    doc.setFontSize(8); doc.setTextColor(...GRIS);
    doc.text('PRODUCTO', M + 6, y + 21);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(11); doc.setTextColor(...OSCURO);
    desc1.forEach((l, i) => doc.text(l, M + 6, y + 27 + i * 5));
  }
  y += hCliente + (multi ? 8 : 6);
  /* ---- Un bloque por producto ---- */
  lista.forEach((c, i) => {
    const r = rs[i], kc = c.costos || {};
    if (multi && i > 0) { doc.addPage(); y = CY; }   // con varios productos, cada uno va en su propia hoja
    doc.setFont('helvetica', 'normal'); doc.setFontSize(11);
    const d = multi ? doc.splitTextToSize(String(c.descripcion || ''), ANCHO - 4) : [];
    const alto = rh => (multi ? 16 + 5 * d.length + (2 * rh + 6) + (9 + 2 * rh + 6) : 24 + (9 + 2 * rh + 5)) + (9 + 5 * rh + 6) + 36;
    RH = 8;                                           // si no cabe, las filas se achican para que el total quede en la misma hoja
    if (y + alto(8) > LIMITE) RH = 7;
    if (y + alto(7) > LIMITE) RH = 6;

    if (multi) {   // etiqueta "PRODUCTO i DE N" + descripción
      doc.setFillColor(...AMARILLO); doc.roundedRect(M, y, 42, 6.5, 1.5, 1.5, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.setTextColor(...OSCURO);
      doc.text(`PRODUCTO ${i + 1} DE ${N}`, M + 3, y + 4.5);
      doc.setFontSize(11);
      d.forEach((l, j) => doc.text(l, M + 1, y + 13 + j * 5));
      y += 12 + d.length * 5 + 4;
    }

    /* Cantidades */
    if (!multi) {
      nuevaPaginaSi(26);
      const cw = (ANCHO - 6) / 2;
      [['CANTIDAD SOLICITADA', cotCantidadTxt(c)], ['CANTIDAD COTIZADA', r.unidades != null ? cotNum(r.unidades, 0) : '—']].forEach(([t, v], j) => {
        const x = M + j * (cw + 6);
        doc.setFillColor(...CLARO); doc.roundedRect(x, y, cw, 18, 3, 3, 'F');
        doc.setFillColor(...AMARILLO); doc.rect(x, y + 3, 1.6, 12, 'F');
        doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.setTextColor(...GRIS);
        doc.text(t, x + 6, y + 6.5);
        doc.setFontSize(15); doc.setTextColor(...OSCURO);
        doc.text(String(v), x + 6, y + 14);
      });
      y += 24;
    } else {
      fila('Cantidad solicitada', cotCantidadTxt(c));
      fila('Cantidad cotizada', r.unidades != null ? cotNum(r.unidades, 0) : '—');
      y += 6;
    }

    /* Precio del producto (COP, ya incluye aumento y comisión) */
    seccion('Datos del precio del producto');
    fila('Costo Total', cop(r.totalCop));
    const comisionCop = r.comisionTotal != null && kc.usdRmb > 0 && kc.usdCop != null
      ? (r.comisionTotal / kc.usdRmb) * kc.usdCop : null;
    fila('Comisión de Adma' + (r.comisionPct != null ? ` del ${cotNum(r.comisionPct, 2)}%` : ''), cop(comisionCop));
    y += multi ? 6 : 5;

    /* Flete */
    seccion('Datos del flete');
    fila('Tipo de flete', r.viaAerea ? 'Aéreo (sujeto a impuestos)' : 'Marítimo');
    fila('Tiempo de llegada', r.viaAerea ? '20 a 25 días' : '90 días');
    fila('CBM', r.cbmTotal != null ? cotNum(r.cbmTotal, 4) + ' m³' : '—');
    fila('Cantidad de cajas', r.cajas != null ? cotNum(r.cajas, 0) : '—');
    const fleteTotal = r.logistica360 != null
      ? cop(r.logistica360) + (r.viaAerea && r.logistica360Usd != null ? ` (USD ${cotNum(r.logistica360Usd, 2)})` : '')
      : (r.logistica360Usd != null ? `USD ${cotNum(r.logistica360Usd, 2)}` : '—');
    fila('Costo Total', fleteTotal);
    y += 6;

    /* Total destacado del producto */
    nuevaPaginaSi(36);
    doc.setFillColor(...OSCURO); doc.roundedRect(M, y, ANCHO, 36, 3, 3, 'F');
    doc.setFillColor(...AMARILLO); doc.rect(M, y + 4, 2.2, 28, 'F');
    doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(255, 255, 255);
    doc.text('TOTAL PRODUCTOS + ENVÍO', M + 9, y + 13);
    doc.setFontSize(20); doc.setTextColor(...AMARILLO);
    doc.text(cop(r.totalConFlete), M + ANCHO - 8, y + 14, { align: 'right' });
    doc.setDrawColor(70, 70, 75); doc.setLineWidth(0.2);
    doc.line(M + 9, y + 19, M + ANCHO - 8, y + 19);
    doc.setFontSize(9); doc.setTextColor(255, 255, 255);
    doc.text('PRECIO POR UNIDAD', M + 9, y + 28);
    doc.setFontSize(12); doc.setTextColor(...AMARILLO);
    doc.text(cop(r.unitConFlete), M + ANCHO - 8, y + 29, { align: 'right' });
    y += 46;
  });

  /* ---- Varios productos: resumen y total de la cotización ---- */
  if (multi) {
    doc.addPage(); y = CY;   // el resumen total va en una hoja nueva
    seccion('Resumen de la cotización');
    lista.forEach((c, i) => fila(`${i + 1}. ${recortar(c.descripcion, 45)}`, cop(rs[i].totalConFlete)));
    y += 8;
    nuevaPaginaSi(34);
    const totalGeneral = rs.every(r => r.totalConFlete != null) ? rs.reduce((t, r) => t + r.totalConFlete, 0) : null;
    doc.setFillColor(...OSCURO); doc.roundedRect(M, y, ANCHO, 24, 3, 3, 'F');
    doc.setFillColor(...AMARILLO); doc.rect(M, y + 4, 2.2, 16, 'F');
    doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.setTextColor(255, 255, 255);
    doc.text('TOTAL DE LA COTIZACIÓN', M + 9, y + 14);
    doc.setFontSize(20); doc.setTextColor(...AMARILLO);
    doc.text(cop(totalGeneral), M + ANCHO - 8, y + 15, { align: 'right' });
    y += 34;
  }

  /* ---- Imágenes en hoja aparte (las reales cotizadas; si no hay, las de referencia) ---- */
  const grupos = lista.map((c, i) => ({ c, i, imgs: (c.imagenesReales || []).length ? c.imagenesReales : (c.imagenes || []) })).filter(g => g.imgs.length);
  if (grupos.length) {
    doc.addPage(); y = CY;
    for (const g of grupos) {
      if (multi) nuevaPaginaSi(84);
      seccion(multi ? `Imágenes · ${g.i + 1}. ${recortar(g.c.descripcion, 26)}` : 'Imágenes del producto');
      const colW = (ANCHO - 6) / 2, maxH = 70;
      let col = 0, altoFila = 0;
      for (const src of g.imgs) {
        const d = await cotDimensionesImagen(src);
        if (!d) continue;
        const k = Math.min((colW - 4) / d.w, (maxH - 4) / d.h);
        const w = d.w * k, h = d.h * k;
        if (col === 0) { nuevaPaginaSi(h + 8); altoFila = 0; }
        const x = M + col * (colW + 6);
        doc.setDrawColor(...BORDE); doc.setLineWidth(0.3); doc.roundedRect(x, y, w + 4, h + 4, 2, 2, 'S');
        doc.addImage(src, 'JPEG', x + 2, y + 2, w, h);
        altoFila = Math.max(altoFila, h + 4);
        if (col === 1) { y += altoFila + 6; col = 0; } else col = 1;
      }
      if (col === 1) y += altoFila + 6;
      if (multi) y += 4;
    }
  }

  /* ---- Pie de página en todas las hojas ---- */
  const n = doc.getNumberOfPages();
  for (let p = 1; p <= n; p++) {
    doc.setPage(p);
    dibujarEncabezado();
    const top = ALTO - PIE;
    doc.setFillColor(...OSCURO); doc.rect(0, top, 210, PIE, 'F');
    doc.setFillColor(...AMARILLO); doc.rect(0, top, 210, 1.5, 'F');
    doc.setTextColor(255, 255, 255);
    doc.setFont('helvetica', 'bold'); doc.setFontSize(10.5);
    doc.text('Juliana Meneses Correa', M, top + 10);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9);
    doc.text('Soporte Comercial · ¿Listo para avanzar? Escríbenos y confirmamos tu pedido.', M, top + 16);
    doc.setTextColor(...AMARILLO); doc.setFont('helvetica', 'bold'); doc.setFontSize(11);
    const wa = 'Whatsapp 3150168453';
    doc.textWithLink(wa, 210 - M - doc.getTextWidth(wa), top + 13, { url: 'https://wa.me/573150168453' });
  }

  const base = String(lista[0].cliente || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'cliente';
  doc.save(`cotizacion-${base}.pdf`);
  return true;
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
  if (!(upc > 0)) { el.textContent = `Cantidad solicitada: ${cotCantidadTxt(c)}${cotEsRango(c) ? ' (los cálculos usan ' + cotNum(c.cantidad, 0) + ')' : ''}`; return; }
  const cajas = Math.ceil(c.cantidad / upc);
  el.textContent = `Cantidad solicitada: ${cotCantidadTxt(c)}${cotEsRango(c) ? ' (calculado con ' + cotNum(c.cantidad, 0) + ')' : ''} · Cajas: ${cotNum(cajas, 0)} · Total unidades cotizadas: ${cotNum(cajas * upc, 0)}`;
}

/* ---------- Excel "A cotizar" (una o varias cotizaciones en un solo archivo) ---------- */
function cotActualizarAcciones() {
  const sel = [...cotSeleccion].map(id => cotLista.find(c => c.id === id)).filter(Boolean);
  const n = sel.length;
  document.getElementById('cotSelTexto').textContent = n ? `${n} seleccionada${n > 1 ? 's' : ''}` : '';
  const clientes = new Set(sel.map(c => cotNorm(c.cliente).trim()));
  const todasConCostos = sel.every(c => c.costos);
  const ok = {
    excel: n > 0 && sel.every(c => cotTipoSel(c) === 'excel'),
    pdf: n > 0 && todasConCostos && clientes.size === 1,
    orden: n > 0 && sel.every(c => cotTipoSel(c) === 'orden')
  };
  const motivo = {
    excel: 'Solo para cotizaciones pendientes',
    pdf: !todasConCostos ? 'Todas deben tener los costos llenos' : 'Deben ser del mismo cliente',
    orden: 'Solo para cotizaciones finalizadas sin orden de compra'
  };
  [['cotBtnExcelSel', 'excel', `A cotizar - Generar Excel (${n})`],
   ['cotBtnPdfSel', 'pdf', `Cotización ${n && sel.every(c => c.origen === 'Groupack') ? 'Excel' : 'PDF'} (${n})`],
   ['cotBtnOrdenSel', 'orden', `Generar orden de compra (${n})`]].forEach(([id, k, txt]) => {
    const b = document.getElementById(id);
    b.hidden = !n; b.textContent = txt; b.disabled = !ok[k]; b.title = ok[k] ? '' : motivo[k];
  });
  const bp = document.getElementById('cotBtnOrdenPrueba');   // solo Super usuario
  bp.hidden = !(ok.orden && pgEsSuper()); bp.textContent = `Orden de PRUEBA (${n})`;
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
      desc: c.descripcion, cant: cotEsRango(c) ? `${c.cantidadMin}-${c.cantidad}` : c.cantidad, link: link ? { text: link, hyperlink: link } : null,
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

/* ---------- Cotización en Excel para clientes de Groupack ----------
   Replica el "FORMATO DE COTIZACIÓN GRUPAL" de Groupack (encabezado con logo, tabla de 21 columnas con bloque Marítimo y bloque Aéreo,
   observaciones y firma). Se usa en lugar del PDF cuando el cliente es de Groupack (origen === 'Groupack').
   Solo se llena el bloque del tipo de flete de cada cotización; el otro queda en blanco. */
const COT_LOGO_GROUPACK = 'data:image/jpeg;base64,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27O97sCN/r6bCc1Honum0it86FKJfnc6JlyttZRP9RUwvhd4HNVP8Shq92uex3m42ypYrKihqJKeRq8rXMcrVT5UA6IAAvKygro7nlRgarhcjo5rLRORU/Us1M0NZdhfMWlxnkfb7N21vpphh7qGeLe77taqron6dCtXd8LFNmgABxzzR08L5pZGxxRtV73vXRGoiaqqrzIBVb2QSqZUbQDo2qirT2iljd1Kqvd5nIarklZ/4+izNzhxbiWmkc+hq6tWUqu54I0SONfG1qL4yNQBYP2NGnclDmLUe0WWijTw6TKV8FnfY6cOyW3KG93eVm6l2uzu1r7pkUbW6+Ur08QE7bRfsEZifEtV9WpScXY7RfsEZifEtV9WpScAAADUDQAANAAOejo6i41cFJSQST1U70jihiarnyOVdEaiJxVVXmJRyx2bsyM2J4vSLDlTFb3ro65VzVgpmJ077k77wNRVLGdnrZCwvkksV5rZG3zGG76/kj0ipdU4pCxeTo3l4r1cgHzZC2elySwVJW3mJndffEbJW6aL6FjTiyBF6tdXacrl6EQlrNrMGkyty6xDiysVqtttM58Uarp22Ze9jZ43q1DNOQr37Ivmr26ssWXFDN3kCJc7ijfdqipCxfAiucqfpNA0Wu91q75da66V8zpq6umfUTyu5XyPcrnL41VTpAAclO2N80bZnrHE5yI96JruprxXTnLOsJbbuSODMMWjD1sTEDKC1UsdLCnpemu6xqJqvf8AKumq+ErBAFq/+0Dyf91iD/x6fbIQ2p9pfKfPHLGSz2x15biChqWVdvfPQo1m8neva528uiKxzvGiGi4ALygACRsiMy5spM1MO4oY5/oSnnSKsY3+0pn97InyLqnWiF2VJVQ11NDU08rZaedjZI5GLq17VTVFRehUUoHLXdhbNbu+ygjsVbPv3jCrkon7y98+nVFWF3iRFZ+51gaudkRs0tDnVbLirFSG42eFWu04K5kkjVTwom78qGopZl2RLL+S+5c2PF1NEr5cPVSxVGia6QT6JvL1I9rE/eKzQBtv2PXGMNgzlrbJUORrcQW58MSqvLLEqSInkpIakHrYXxHcMH4jtV/tMyw3K2VDKmCToexdU16U4aKnQBfRygjzJbN6yZ04FocSWiVjZnNSOto97V9JOid8xydHOi86aKSGAAAAHUudzo7Nb6m4XCqh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';

async function cotGenerarExcelGrupal(lista) {
  lista = (lista || []).filter(c => c && c.costos);
  if (!lista.length) { alert('Estas cotizaciones aún no tienen costos guardados.'); return false; }
  if (!window.ExcelJS) { alert('No se cargó la librería de Excel. Revisa tu conexión y recarga la página.'); return false; }
  const N = lista.length;
  const rs = lista.map(c => cotCalculos(c) || {});
  const F = 'Nunito';
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Cotización', { views: [{ showGridLines: false, zoomScale: 40 }] });
  [10.33, 19.44, 44, 37.33, 31.44, 31.44, 31.44, 18.44, 26.55, 27.55, 27.55, 27.55, 38.44, 31.89, 28, 28, 28, 35.89, 33.66, 33.66, 31.55, 31.33]
    .forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  /* ---- Estilos ---- */
  const LET = n => String.fromCharCode(64 + n);                       // 1 -> A … 22 -> V
  const NEGRO = 'FF000000', AMARILLO = 'FFFDEE14', AMARILLO_M = 'FFFFFF00', VERDE = 'FF70AD47', GRIS = 'FFD9D9D9';
  const relleno = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const lado = s => (s ? { style: s } : undefined);
  const bordes = (l, r, t, b) => ({ left: lado(l), right: lado(r), top: lado(t), bottom: lado(b) });
  const MON_RMB = '_ [$¥-804]* #,##0.00_ ;_ [$¥-804]* \\-#,##0.00_ ;_ [$¥-804]* "-"??_ ;_ @_ ';
  const MON_COP = '_-[$$-409]* #,##0.00_ ;_-[$$-409]* \\-#,##0.00\\ ;_-[$$-409]* "-"??_ ;_-@_ ';
  const MON_USD = '"USD" #,##0.00';
  const poner = (ref, valor, o = {}) => {
    const c = ws.getCell(ref);
    if (valor !== undefined) c.value = valor;
    c.font = { name: o.fuente || F, size: o.tam || 11, bold: !!o.negrita, color: o.color ? { argb: o.color } : { argb: NEGRO }, underline: o.subrayado };
    if (o.relleno) c.fill = relleno(o.relleno);
    if (o.bordes) c.border = o.bordes;
    c.alignment = { horizontal: o.h || 'center', vertical: 'middle', wrapText: o.ajustar !== false };
    if (o.fmt) c.numFmt = o.fmt;
    return c;
  };

  /* ---- Encabezado: empresa, logo y datos de la cotización ---- */
  [[2, 'GROUPACK COLOMBIA'], [3, 'CALLE 64 No. 9-05 B/PAMPALINDA'], [4, 'TEL: 57 315-6487448'], [5, 'CALI- COLOMBIA']].forEach(([f, t]) => {
    ws.mergeCells(`B${f}:V${f}`);
    poner(`B${f}`, t, { tam: 24, negrita: true, ajustar: false, bordes: bordes('medium', 'medium', f === 2 ? 'medium' : null, null) });
  });
  [38.4, 35.4, 35.4, 35.4, 36, 28.95, 36.6, 24, 30, 24.6, 14.4, 9.6, 35.25, 49.5].forEach((h, i) => { ws.getRow(i + 2).height = h; });
  for (let f = 6; f <= 12; f++) {                                  // marco exterior del encabezado
    poner(`B${f}`, undefined, { bordes: bordes('medium', null, null, f === 12 ? 'medium' : null) });
    poner(`V${f}`, undefined, { bordes: bordes(null, 'medium', null, f === 12 ? 'medium' : null) });
  }
  for (let c = 3; c <= 21; c++) ws.getCell(`${LET(c)}12`).border = { bottom: { style: 'medium' } };

  const clientes = [...new Set(lista.map(c => String(c.cliente || '').trim()).filter(Boolean))];
  const vias = new Set(rs.map(r => (r.viaAerea ? 'Aereo' : 'Maritimo')));
  const modalidad = vias.size > 1 ? 'Maritimo-Aereo' : [...vias][0];
  const hoy = (typeof hoyBogota === 'function' ? hoyBogota() : new Date().toISOString().slice(0, 10)).split('-').map(Number);
  const info = [
    [7, 'FECHA', new Date(Date.UTC(hoy[0], hoy[1] - 1, hoy[2])), 'dd/mm/yyyy'],
    [8, 'CLIENTE:', clientes.join(' / '), null],
    [9, 'CONDICION DE PAGO', 'CONTADO', null],
    [10, 'MODALIDAD', modalidad, null],
    [11, '% NACIONALIZACION', null, null]
  ];
  info.forEach(([f, etq, val, fmt], i) => {
    const sup = i === 0 ? 'medium' : 'thin', inf = i >= 3 ? 'medium' : 'thin';
    ws.mergeCells(`O${f}:Q${f}`);
    poner(`O${f}`, etq, { tam: 10, negrita: true, color: AMARILLO, relleno: NEGRO, h: 'left', bordes: bordes('medium', 'thin', sup, inf) });
    ws.mergeCells(`R${f}:S${f}`);
    poner(`R${f}`, val, { tam: val && val.length > 22 ? 12 : 16, negrita: true, fmt, bordes: bordes('thin', 'thin', sup, inf) });
  });
  ws.getCell('R10').dataValidation = { type: 'list', allowBlank: true, formulae: ['"Maritimo,Aereo,Maritimo-Aereo,compras China"'] };
  ws.getCell('R11').dataValidation = { type: 'list', allowBlank: true, formulae: ['"30,40,50,60,70,80,90,100"'] };
  const idLogo = wb.addImage({ base64: COT_LOGO_GROUPACK, extension: 'jpeg' });
  ws.addImage(idLogo, { tl: { nativeCol: 1, nativeColOff: 536193, nativeRow: 2, nativeRowOff: 304800 }, ext: { width: 520, height: 278 } });

  /* ---- Encabezados de la tabla ---- */
  ws.mergeCells('M14:Q14'); ws.mergeCells('R14:V14');
  poner('M14', 'MARITIMO', { negrita: true, relleno: AMARILLO_M, bordes: bordes('medium', null, 'medium', 'medium') });
  for (let c = 14; c <= 17; c++) ws.getCell(`${LET(c)}14`).border = bordes(null, c === 17 ? 'medium' : null, 'medium', 'medium');
  poner('R14', 'AEREO', { negrita: true, color: 'FFFFFFFF', relleno: VERDE, bordes: bordes('medium', 'medium', 'medium', 'medium') });
  const TIT = ['ID COTIZACION', 'IMAGEN', 'NOMBRE DEL PRODUCTO', 'DESCRIPCION', 'ESPECIFICACIONES DEL PRODUCTO', 'LINK DE VISUALIZACION', 'SHIPPING MARK',
    'CAJAS', 'CANTIDAD', 'CBM TOTAL', 'PRECIO PRODUCTO EN RMB', 'PRECIO UNITARIO PRODUCTO', 'PRECIO UNITARIO FLETE', 'VALOR TOTAL UNITARIO',
    'VALOR TOTAL FLETE', 'VALOR TOTAL EN DESTINO', 'PRECIO UNITARIO AEREO', 'PRECIO FLETE AEREO', 'VALOR TOTAL UNITARIO AEREO',
    'VALOR TOTAL FLETE AEREO', 'VALOR TOTAL EN DESTINO'];
  TIT.forEach((t, i) => {
    const c = i + 2, aereo = c >= 18;
    poner(`${LET(c)}15`, t, { negrita: aereo, color: aereo ? 'FFFFFFFF' : NEGRO, relleno: aereo ? VERDE : AMARILLO,
      bordes: bordes(c === 2 || c === 13 || c === 18 ? 'medium' : null, c === 22 ? 'medium' : null, 'medium', 'medium') });
  });

  /* ---- Una fila por cotización ---- */
  const PRIM = 16, ULT = PRIM + N - 1;
  const MAXW = 300, MAXH = 330;
  for (let i = 0; i < N; i++) {
    const c = lista[i], r = rs[i], f = PRIM + i, k = c.costos || {};
    ws.getRow(f).height = 265.2;
    const link = cotLinkSeguro(c.link);
    const marca = (c.orden && c.orden.marca) || `${cotNombreClave(c.cliente)}-ADMA-GPK`;
    const celda = (col, valor, o = {}) => poner(`${col}${f}`, valor, { relleno: GRIS, bordes: bordes(col === 'B' ? 'medium' : 'thin', col === 'V' ? 'medium' : 'thin', 'thin', 'thin'), ...o });
    const num = v => (cotOk(v) ? v : null);
    celda('B', `COT-${String(c.id || '').slice(0, 5).toUpperCase()}`);
    celda('C', null);
    celda('D', c.descripcion || '', { tam: String(c.descripcion || '').length > 110 ? 12 : 16 });
    celda('E', null, { tam: 16 });
    celda('F', null, { tam: 16 });
    celda('G', link ? { text: link, hyperlink: link } : null, { fuente: 'Arial', tam: 14, color: 'FF1D4ED8', subrayado: true, h: 'left' });
    celda('H', marca, { tam: marca.length > 14 ? 12 : 16 });
    celda('I', num(r.cajas), { tam: 16, fmt: '#,##0' });
    celda('J', num(r.unidades), { tam: 16, fmt: '#,##0' });
    celda('K', num(r.cbmTotal), { tam: 16, fmt: '#,##0.0000' });
    celda('L', num(r.unitConComision), { tam: 18, fmt: MON_RMB });

    const aereo = !!r.viaAerea;
    const [cP, cF, cT, cTF, cD] = aereo ? ['R', 'S', 'T', 'U', 'V'] : ['M', 'N', 'O', 'P', 'Q'];
    const [oP, oF, oT, oTF, oD] = aereo ? ['M', 'N', 'O', 'P', 'Q'] : ['R', 'S', 'T', 'U', 'V'];
    [oP, oF, oT, oTF, oD].forEach(col => celda(col, null, { tam: 16, fmt: MON_COP }));       // el bloque que no aplica queda vacío
    const unit = num(r.precioUnitCop), fl = num(r.logisticaUnit);
    const B = aereo ? { negrita: true } : {};
    celda(cP, unit, { tam: 18, fmt: cP === 'M' ? '_-"$"\\ * #,##0.00_-;\\-"$"\\ * #,##0.00_-;_-"$"\\ * "-"??_-;_-@_-' : MON_COP, ...B });
    celda(cF, fl, { tam: aereo ? 18 : 16, fmt: MON_COP, ...B });
    celda(cT, unit != null && fl != null ? { formula: `${cP}${f}+${cF}${f}`, result: unit + fl } : num(r.totalUnitCop), { tam: aereo ? 18 : 16, fmt: MON_COP, ...B });
    celda(cTF, fl != null && num(r.unidades) != null ? { formula: `${cF}${f}*J${f}`, result: fl * r.unidades } : num(r.logistica360), { tam: aereo ? 18 : 16, fmt: MON_COP, ...B });
    celda(cD, unit != null && fl != null && num(r.unidades) != null ? { formula: `${cT}${f}*J${f}`, result: (unit + fl) * r.unidades } : num(r.totalConFlete), { tam: aereo ? 18 : 16, fmt: MON_COP, ...B });

    const src = (c.imagenesReales || []).length ? c.imagenesReales[0] : (c.imagenes || [])[0];
    const d = src ? await cotDimensionesImagen(src) : null;
    if (d) {
      const e = Math.min(MAXW / d.w, MAXH / d.h, 1);
      const w = d.w * e, h = d.h * e;
      const idImg = wb.addImage({ base64: src, extension: /^data:image\/png/.test(src) ? 'png' : 'jpeg' });
      ws.addImage(idImg, { tl: { nativeCol: 2, nativeColOff: Math.round(Math.max(0, (313 - w) / 2) * 9525), nativeRow: f - 1, nativeRowOff: Math.round(Math.max(0, (353 - h) / 2) * 9525) }, ext: { width: w, height: h } });
    }
  }

  /* ---- Fila de cierre ---- */
  const CIERRE = ULT + 1;
  ws.getRow(CIERRE).height = 24;
  for (let c = 2; c <= 22; c++) poner(`${LET(c)}${CIERRE}`, undefined, { bordes: bordes(c === 2 ? 'medium' : null, c === 22 ? 'medium' : null, 'thin', 'medium') });

  /* ---- Totales (cuadro a la derecha) ---- */
  const T0 = ULT + 4;                                               // fila de "OBSERVACIONES"
  const hayAereo = rs.some(r => r.viaAerea);
  const sumaDestino = rs.reduce((t, r) => t + (cotOk(r.totalConFlete) ? r.totalConFlete : 0), 0);
  const sumaCbm = rs.reduce((t, r) => t + (cotOk(r.cbmTotal) ? r.cbmTotal : 0), 0);
  const sumaUsd = rs.reduce((t, r) => t + (r.viaAerea && cotOk(r.logistica360Usd) ? r.logistica360Usd : 0), 0);
  const filasTotales = [
    hayAereo ? ['FLETE AEREO (USD)', 0, MON_USD] : null,
    ['TOTAL EN DESTINO (COP)', { formula: `SUM(Q${PRIM}:Q${ULT})+SUM(V${PRIM}:V${ULT})`, result: sumaDestino }, '"$" #,##0.00'],
    ['CBM TOTAL', { formula: `SUM(K${PRIM}:K${ULT})`, result: sumaCbm }, '#,##0.0000']
  ].filter(Boolean);
  if (hayAereo) filasTotales[0][1] = sumaUsd;                       // el flete aéreo en USD es el valor ingresado en cada cotización
  filasTotales.forEach(([etq, val, fmt], i) => {
    const f = T0 + i;
    ws.getRow(f).height = 43.5;
    ws.mergeCells(`O${f}:Q${f}`); ws.mergeCells(`R${f}:U${f}`);
    poner(`O${f}`, etq, { tam: 10, negrita: true, color: AMARILLO, relleno: NEGRO, h: 'left', bordes: bordes('medium', 'thin', 'thin', 'thin') });
    poner(`R${f}`, val, { tam: 16, negrita: true, fmt, bordes: bordes('thin', 'thin', 'thin', 'thin') });
  });

  /* ---- Observaciones y firma ---- */
  poner(`C${T0}`, 'OBSERVACIONES', { negrita: true, h: 'left', ajustar: false });
  [
    '➢ Cotización no incluye Fulfillment',
    '➢Valor es hasta la ciudad de Cali',
    '➢Confirmar aprobación de la cotización',
    '➢El tiempo de entrega maritimo es de 60 días y aereo es de 15 dias a partir de que la mercancía se encuentre en la bodega en origen',
    '➢Se calcula la tasa del dólar del día que realice el pago',
    '➢Valor sujeto a cambio de acuerdo con el comportamiento del flete'
  ].forEach((t, i) => { ws.getRow(T0 + 2 + i * 2).height = 24; poner(`C${T0 + 2 + i * 2}`, t, { tam: 16, h: 'left', ajustar: false }); });
  ['Cordialmente,', 'Libardo González', 'Groupack sas', 'Colombia'].forEach((t, i) => poner(`C${T0 + 17 + i}`, t, { tam: 16, h: 'left', ajustar: false }));

  ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, printArea: `B2:V${T0 + 20}` };

  const buf = await wb.xlsx.writeBuffer();
  const url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  const a = document.createElement('a');
  const slug = String(clientes[0] || 'cliente').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '').toLowerCase() || 'cliente';
  a.href = url; a.download = `cotizacion-groupack-${slug}${clientes.length > 1 ? '-y-otros' : ''}-${hoy.join('')}.xlsx`;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
  return true;
}

/* ---------- Cotización de UN producto de Groupack: formato carta (2 hojas) ----------
   Hoja 1: encabezado, datos de la cotización, producto con su imagen y valores. Hoja 2: observaciones y firma.
   Cuando son varias cotizaciones se usa cotGenerarExcelGrupal (tabla). */
const COT_GPK_CABECERA = 'data:image/png;base64,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t//888+VUg+v1/vcc889+OCD5mvVqFHjuuuuK7OG1zTt+eefN19eaTRrwK233tqyZUulKnv37m3Xrt23335rvkpmZuawYcNUp34jc6nfVVddVb9+faXNfvzxx0OGDHG73SbLr1ix4pZbbjE/+15Az549K+Lbfe3ataqjicvMsWPHVq5cGaRAr169VBfBPHDgQPv27RcuXGi+SnZ29j333PP//t//U72yFbnCQKmaOHGi0rBTKeWwYcNcLhc+9wAgjILPrXbBmdcAAKBCwJIFAFzMIFQ2SLp8ui2qjjOqM1m7EjlK9Tg0TXv88ccffvhhpVqZmZn/+te/Pv/885deeqlVq1bFzWse4Ha709LS3nrrLdU+bs8884ymaWX5qvTt2/ff//73li1bTJafP3/+2rVrlcZFPv3006qTrB04cKBnz559+/Z94YUXgqcYBQUFM2fOfO65506dOqX63G+++WaTT6R79+5vv/220sanT5++devW8ePHd+nSJUixEydOvPHGG5MmTfL7/Urbt9vtHTp0qIgXgvT0dKXyQ4YMUe0SdbapU6cqRbRpaWlt2rQp7lEhxKhRo4YOHap0DH/++We/fv26des2bty4Ro0aBb90TJ8+PTU11fwqGYWXtX/9619l/2rWq1fvmWeeURra+fvvv6empoYQggMAmNRo2x+I0gAALj6I1aCSM5jdpEdLGU1ksOYi9kudWbOw8Lg9TrfWKanKbSRan8nUZKkOXx08ePArr7wSwkjD1atXt2/fvmbNmr169WrXrl3NmjVr1KiRnJzsdrszMjJOnDixYcOGtWvXfvfddxkZGaobv+KKK4YMGVLGLwwzjx07Vmku9jFjxigt49ClS5emTZuaT+4Kpaenp6enN2nSpHfv3i1btqxRo0bNmjWjoqKOHz9+9OjRffv2LVy4cPHixQUFBaE9d/MDVO+444533nlHqa8iEW3fvr1v375Nmzbt2bNnhw4dGjRoEB8fz8xut/vYsWNr1qz55ptv5s2bF9ok7l27do2Nja1wFwIp5Zw5c5SqjBw5skGDBiHv0W63K8VqX3311VtvvWWxFPupPWDAgNTU1BBmAFywYMGCBQsaNmzYq1evVq1aBa4esbGx2dnZJ0+e3Lx588qVK+fPnx9CQExEw4YNu/rqq8vlNR05cuSMGTN27NhhvsrEiRP79Omj2o8VAMCMwAIFSNYAAC4+iNWg0pMakWBDIxakawa72CJ0v+b3OK1RLWKjh7JoSmQ/U7p0p4Sz2+0fffRR9+7dg8+jVJwjR45Mnjx58uTJ4T2qsWPHWq3Wsn9lunTpct111/38888my//0009ff/11586dTZZn5g8++KB9+/ZerzeEw9u2bVsJl4wo0p133hmkU9I5mjVrNmDAgC+//DKEHW3ZsmXLli2pqalEZLFYnE5nbm5uCQ/eZrONHz++Il4Gfvnll0OHDpkvX79+/ZJkakTUvn17p9NpftThqVOnli9ffuuttxZXwGq1vvfee71791aNWQN27dqlOtP/BVWrVk11dHYYWa3W999/v3379uYbRNf1oUOHrlmzpiJODggAEeucFT8DNxGuAQBcNDC3GlRu0kr+ZCKNtSwWeSylEBbSRL4/hq23WGPudtgaMFnK8p1yyy23PPfcc5HTQk2bNr399tvLa++qGc0LL7yglEi2bNnyzTffjKjWfv/995WqvPTSSyWfuMrv95c8UyOihx9+uF69ehXxSqA6AjT4EFozoqKi2rdvr1Ql+HqgRNSpU6dnnnkmclp1woQJ8fHx5XgArVq1uv/++5Wq/Prrr6+++io+GwGg5ALB2TmZWiGEawAAFw3EalCZSSImaWMpiH3E+aQVkKbl5sdZHTfYE+8iSzdJCUQ2ojKdVuzpp5/u2LFjJDRQdHT0f/7zn+DztZWqtm3b3nzzzebLb9u2TbXr1gMPPDB48OBIaO34+Pgvv/zS6XQq1apVq9YTTzwRCcdfvXp11fUoI4fqGqAlj9WIqFu3bkrl58yZc8GelWPGjFFN60pJnz59Bg4cWO6H8dJLL1WvXl2pyhtvvLF582Z8QAJAyRWXqZl5FAAAKgrEalCZMZFOIp+IyLAS+UhzuzxOi+3GqPgBpLUhjirjrmqnD4v5008/vfTSS8v56iDEZ5999s9//rN8D2PcuHFK5cePH+/xeJSqvPvuu02bNi3nc5H5k08+ueKKK0Ko+9hjj0XC8b/77rsVcVY1Ivrll1+UpiSLj49v3bp1yffbtWtXIRQuL9nZ2RecPVAI8fnnn5f71ePGG2/85JNPIuHFjY+PV+2R6vf7hw4d6vP58BkJACVhJjVDsgYAcBFArAYVlCxm+U7VMMBga4HkAtIkWZwuXyJr10Qn9CfRnjimHJ9eYmLizJkzk5KSyvEYJkyYoNqbpjRce+21Sodx4MCBDz/8UGkXTqdzxowZ1apVK8en+dxzz4XcAcrpdM6bNy+0SC5c3n777Z49e1bQq4nqCNAOHTqEZbbB5OTkG264QanKrFmzLlgmKSnpiy++KPnQ4JBdddVVaWlpdrs9Ql7fvn37durUSanKli1bImp4OAAAAABELMRqUMlJQ/eRxeeVvlxXAot2zmr3kdaWRHR4YrsSuOaaa1asWPGPf/yjXPb+4IMPPvzwwxHyIo0dO1ZpIOqECRNycnKUdlGnTp2lS5fWrl0+U5yMGzfu2WefLckWqlevvnDhwpo1a5bL8b/44ovDhg2ruFeBchkBGtC9e3el8gsXLnS73Rcsdu211y5cuLBc5jW7/PLL582bV75Tqp3v3XffjYqKUqry6quvbt++HZ+RAAAAABAcYjWo7G8BJjtpdreRSPZWjvgBRB2Iq5R7phZQt27d4Gv/lQZN01JTU999993IeZGaNGmitGxCRkbG22+/rbqXevXqLVu2rIwHvdpstkmTJo0ePbrkm6pTp878+fOrVq1axq/OyJEjK+6UakS0fv36AwcOKL1BVLs+BaEaq+Xm5n7zzTdmSl5//fVLliy57LLLyrIxb7311tWrV5f7ENTz1a5d+/nnn1eq4vV6hw0bpus6PiYBAAAAIFimgCaAyoyJiGVBbrzd1jY28S6ydyCKPr2UAXEkHGF8fPzs2bOHDh1aNrurWbPm4sWLR40aVY7LFBTp+eeft1gs5su/9957x44dC+HpL1++fMSIEWXzpOrVq7d8+fIhQ4aEa4ONGzdet25dGPtSBZeUlDRjxoxXXnmlQl8EVEeAXn/99YmJiWE8B6666iqlKjNnzjRZ8uqrr/7555/LZig3Mz/77LNz584NY+OE10MPPaQ6BeH69esnTpyID0oAAAAACAKxGlRcJc59JEsyvD6X3XmVPfZ24lvOytQiiMVieffdd7/44ovS7gNyyy23rF279sYbb4zAF7t+/fpK63Xm5+eHFvfY7fY333xzzpw5pTpVmaZpI0aMWL169TXXXBPeLVerVi09Pf29995THfKmqlOnTuvXr+/Ro0dFv46ojgDt2rVreA9ANfZavHhxfn6+ycJVqlSZNWvWBx98UKppV61atebMmfP8888rrcBQ9hfSSZMmqR7h+PHjf/vtN3zcAgAAAECxwYSUEq0A5cC/W/eO8xZ865A6SUnSyYaFSCfhM4RfskHERIKlZPIzkU5xJCVzNgvD463tSL6XrHcRXVJUCuY2XOPzsr6MsZ1iXZKMIy2bDCfpscSSRB5pOknN8BuaQxa4tALZObF6X6GlECUTUQTGaoXy8/PfeOONyZMnZ2dnh3fLN9xww5NPPlmSXk579+5t3LixycLdu3c33+Om0KFDh5o0aWJ+lU+LxbJx48b69euH9ow8Hs9777335ptvZmVlhbe1b7311gkTJphvrtD8/vvv48ePnz17tt/vD++WGzZs+Pjjj999992RfLaYtGHDhpSUFNUqqv3Lglu3bt1NN92kVGXKlClKw6KJKDMzMzU19ZNPPlFdJze4WrVqjRo16r777gttgYKyPxOeeOKJSZMmKVVp1arVd999F8mJIQAAAACUI3xNhIhhCDI0ksySiNggIaUgEiyZDGK/1SCftOZJzU+GlTiK6IJDAs+kY4aFSBAxSUHSKnXdIA9bLPk5Vqvjupjkf5HoSJRwZj41jtgWio6OHjt27G+//fbqq6+Ga3L9jh07LlmyZNmyZWU2cjBkl1566YMPPmi+vN/vf/HFF0Pend1uHzVq1O+///7mm2+GpeeazWYbOHDg2rVr58+fX9qZGhHVr19/ypQpO3fuHDVqVEJCQsk3yMwdOnSYO3fuxo0bS5KpRRTVEaB16tQJb6ZGRC1atFBda8LMeqDnSEhIeOutt3777bennnoqLEsM16xZ880339yxY8fw4cMjZ9HPC3rxxRcvueQSpSqrV6+ePHkyPqIBAAAAoEiI1aDcydP/1dykuYh1yUJKOxsONhysO8hwknSwzLewi5g8hpan2yUnkIwxtXFpIX8ySQtrWSxymSSRMIgK9Dhha2+Nu89hu1KwhUiL5EDtbHFxcY899tjOnTt/+OGHhx56KITeWHFxcV26dHn99dc3btw4d+7cyBz1WaTRo0dHR0ebL//VV19t2LChJHuMjo4eMWLE9u3bV6xY8cgjj9SrVy+ELdx2223//e9/Dxw48Mknn6jO7lRCl1xySWpq6p49ez766KNBgwaFkA/GxMS0b9/++eef37hx47x58zp27Bhp8+6VRLmPACUiZlYdB/rtt9+G1mW1WrVqL7744r59++bOnTt48OBq1aqpHmrz5s2ffvrp77///rfffhsxYkQFCtQCYmNjQ1jP5IUXXti3bx8+rQEAAACgiC/JGAQK5eOvQaB+lmRIBws3k98gG0mbQRaSmpDEZLDUifxELrJ6PMwuo6bkG52OgY6YTsXkwoWDQDNY18lIJsPJIp84h9hHkqVmK9ATdGoeV/0BorZE9ooSqBXnxIkTq1ev3rJly759+/bu3XvkyJG8vLyCggKv15uQkJCcnJyUlJSUlJScnFy3bt22bds2b95c0zScg6E5fvz46tWrN23atHfv3n379h05ciQ/P7+goMDv90dFRTmdzipVqtSpU+eKK6648sorb7jhhquvvlppsYXSdvTo0Z9++mnt2rVHjx7NyMg4depUZmZmRkaG2+1OSEhISkoK/DcxMbFx48YpKSnNmjXD2XIR27Nnz5o1a7Zt27Zv3759+/adPHkycD5LKaOjo6OiopKTk+vUqVO3bt1rrrnm1ltvTU5ORqMBAAAAABRCrAbl5O+xmpR2Q8snNqR0kmEjsggpmfzEXpaSyCB2k3DkeaobjhZxSQMkNWcqbiDP2bGaQUYCETFJIjeJHLJIt6+G1NrY4/sL6y1EMXgpAAAAAAAAACAEFjQBlLvTo0CZJAmWFiKNpWTNLamAyMuskRTEVldBVc2WEhPfg6gVk/m5gXQSHsksDb+wx7oKrFL8M6pKX7J0IHJG8gIFAAAAAAAAABDJEKtBhGCSFiZiqbEkJp9BBcQFpJEuNWnY3a4Eq/1ae2IvstxCFK2yYdalj4XONnt2bozVen1U9YFENyNTAwAAAAAAAICSQKwG5Y+JDCKWxMRsCCImoUvpI41IOn3+GI83QbO2sFe7neg6omi1OEyyJhw6+z3+RLK1iEq4iygQzBlYsgMAAAAAAAAAQoZYDSKFYMmSiHSSGkli1qS0+vUqXl8tq/VKZ0JfouuIqqp3MZNEUvcmG9wivsYgonZEUYEdos0BAAAAAAAAIGSI1aDcFGZjkohJsjRIEklJJEmSIJsunV5vssVyZXS1nkQ3E1UxueG/tiyZWfoMj8Va3x47UMp2zFEY+wkAAAAAAAAAJYdYDcpYINLyExXoXpfQJQuW7DeEl2UCSS9rWSwMkhYpovK98dJxZVxST6KbiKqcWdug+ERMSmIm0lk6nX6DhVvao7LcMV7ZLTm+K1lbMyFTAwAAAAAAAIDwQKwG5UIn8kldMlnJMKSQkkg3SJDQNEnCIw2ZX+C0OJtFJ/Ulamd27Ofpx71kIZYetjpz8+Mtcc1ttkFkaUFkO7sQAAAAAAAAAEBJIFaDMhZItTTSHRpHG4aDyM/EgpkoWxMsZbRft7q8Ds3RLDquH1FnohjFLmZesrrdmm71JTNfGxvbn6ihlFYiDa0PAAAAAAAAAOGCWA3K6cQTdqE5JAmSzJKJyMI6kfRLp9tflayNoqsOItlJPVMjIoPY6tOqSboitmo/onZEVZjRSQ0AAAAAAAAAwgmxGpQTqbGQxB4SXmIfkUZklyw9foewNYlOvItkR+I4IkN9yc5otytBpxYJ1btJeSNTQmB/GP4JAAAAAAAAAGGEWA3KidDJ4tJFjiAXMxnsFFylwO1gW73opIFEt5xZo0ApUwsEZzFu15VxCY2ImjBX+/tDAAAAAAAAAADhgVgNyovfkAW6KLAInaTVIHuuO9pmaxqV1IOoHVFyCbbM8VXasLARWYmIyJDEjFgNAAAAAAAAAMIKsRqUsTODMdlvkJc1Icnh12M8vmrCeVVUYj+iW8/0U+OQu5ixiD57j8jUAAAAAAAAACDsEKtBOTEMYouUUT49RvfWtGj/iErsRdSGqAoRlXjMpofI9teqowAAAAAAAAAA4SbQBFBONJJ23Yj2eRIF14uq0VPKW0o29vMvBmVJcqGJAQAAAAAAAKD0oLcalLEz3dCkJkV8vhHrdDaMiu9H1Ja5Srj2IShRYuAnAAAAAAAAAJQmxGpQxk7PrWYYlF+QICxXxyf2IWpHlPTXtGthYEWoBgAAAAAAAACliqWUaAUoQ5JIJ9LJ/4fXvV6zaZrteqLaaBcAAAAAAAAAqFgQq0FZk5TP5JK6j8lLnEgilihwEqKHGQAAAAAAAABUGIjVAAAAAAAAAAAAlGElUAAAAAAAAAAAAGWI1QAAAAAAAAAAAJQhVgMAAAAAAAAAAFCGWA0AAAAAAAAAAEAZYjUAAAAAAAAAAABliNUAAAAAAAAAAACUIVYDAAAAAAAAAABQhlgNAAAAAAAAAABAGWI1AAAAAAAAAAAAZYjVAAAAAAAAAAAAlCFWAwAAAAAAAAAAUIZYDQAAAAAAAAAAQBliNQAAAAAAAAAAAGWI1QAAAAAAAAAAAJQhVgMAAAAAAAAAAFCGWA0AAAAAAAAAAEAZYjUAAAAAAAAAAABliNUAAAAAAAAAAACUIVYDAAAAAAAAAABQhlgNAAAAAAAAAABAGWI1AAAAAAAAAAAAZYjVAAAAAAAAAAAAlCFWAwAAAAAAAAAAUIZYDQAAAAAAAAAAQBliNQAAAAAAAAAAAGWI1QAAAAAAAAAAAJQhVgMAAAAAAAAAAFCGWA0AAAAAAAAAAEAZYjUAAAAAAAAAAABliNUAAAAAAAAAAACUIVYDAAAAAAAAAABQhlgNAAAAAAAAAABAGWI1AAAAAAAAAAAAZYjVAAAAAAAAAAAAlCFWAwAAAAAAAAAAUIZYDQAAAAAAAAAAQBliNQAAAAAAAAAAAGWI1QAAAAAAAAAAAJQhVgMAAAAAAAAAAFCGWA0AAAAAAAAAAEAZYjUAAAAAAAAAAABliNUAAAAAAAAAAACUIVYDAAAAAAAAAABQhlgNAAAAAAAAAABA2f8HUbRM/4vVk0QAAAAASUVORK5CYII=';
const COT_GPK_PIE = 'data:image/png;base64,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tgIoV66cqNnso0ePatWq1blz5+Dg4Nu3b+/du1d2nc0sKNFPk55gLXo2MHPmzH379omeW7x8+XL48OERERH6Kb179548ebLojeAbN260atWqUKFCRYsWffny5e3bt6UD++bLl+/NN99UsGQUWRNfX986deoYxnOXLl3Sv9UVFBTUu3dvw84HXr58ac5wK9JHTYIg+Pv7S7fIyJEjpTfqgiCsXLlSet/rUJTavZWyfv160dO7jIyMFStWiB6bKXU0KXj6cjqOVnvadJAfxzlaqTGdscYkU4PD4iVQl03WTBzJjl8r66IrxTM1QRACAwNXrVpl3TPA/v37DxgwQDp9+vTpTZs2tXRp06ZN0z10kmWLZcoaOHCgiV7VjSlevPiGDRtk+ztISUmRNlVzc3P7/PPPOTBdm926hvG1PSvWqlChQjt27MjywayssmXLbtq0Sf8anfU1upvbmjVrLL3DbNmypfTaNygoyJwL4jJlysieFbPUvXv3LOcJCgoSDU8mO3LLs2fPVq1aNXPmzE2bNunuELJ868S6gpo8ebJFnd+FhIR888030un79+83zJUCAgKWLFni7e0tnTMxMfHChQs3b96UJlmCICxYsEC/zyhSMoqsiSAIzZs3N/xUo9EYPiWaOnWq7LjSIk2bNi1ZsqSJGVQq1Zw5c4oUKSKa3qFDB3M2U926defMmeNQp1Cldm+lyPa4JHopWMGjScHTF7Wns8jZo5Ua0xlrTCdFppYbEKu5LAfpZy07d842WvLrr78eGRlp+npdxNPTc8KECQsWLJAv6jx5Nm/eLPss3djS5s2b98EHH5jafDZYpjFLliwZOXKktK24MTVr1ty5c6ex+6LFixfHx8eLJr755ptVqlThqHRtohuDfPny5bZRX8uXL3/48OG2bdta9K033njj6NGjsj2zWKFIkSJ79+411umhVJcuXdavXy8bkQ8bNsxEC1NBEPz8/H766SfrztWTJ082PYyav7//mjVrRN2Qt27dOsu3vYKCgmQ7cs5mQQ0YMEDacUyWBg4cKDsu8yeffBIbG2t4b7lhwwbzS9LDw2PBggXdunVTvGSyvyaCIEif7sydO1f/9lP+/Pm3bdtmesTbfPnyzZ0794svvjBWMfn6+s6fP79Xr17GKjVjw+no964ff/zR/FrPPhTcvRURGhoqnSh7TlDqaFLw9EXtaedL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';

async function cotGenerarExcelGrupalUna(c) {
  if (!c || !c.costos) { alert('Esta cotización aún no tiene costos guardados.'); return false; }
  if (!window.ExcelJS) { alert('No se cargó la librería de Excel. Revisa tu conexión y recarga la página.'); return false; }
  const r = cotCalculos(c) || {}, k = c.costos;
  const aereo = !!r.viaAerea;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Cotización', { views: [{ showGridLines: false }] });
  const PX = [190, 190, 15, 215, 215];                              // ancho de A..E en píxeles (825 en total)
  PX.forEach((p, i) => { ws.getColumn(i + 1).width = (p - 5) / 7; });
  const ANCHO = PX.reduce((a, b) => a + b, 0);
  const F = 'Arial', NEGRO = 'FF000000', AMARILLO = 'FFFFFF00';
  const fino = { style: 'thin', color: { argb: NEGRO } };
  const caja = { top: fino, left: fino, bottom: fino, right: fino };
  const poner = (ref, valor, o = {}) => {
    const x = ws.getCell(ref);
    if (valor !== undefined) x.value = valor;
    x.font = { name: F, size: o.tam || 10, bold: !!o.negrita, color: { argb: o.color || NEGRO } };
    if (o.relleno) x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: o.relleno } };
    if (o.caja) x.border = caja;
    x.alignment = { horizontal: o.h || 'left', vertical: 'middle', wrapText: o.ajustar !== false, indent: o.indent || 0 };
    if (o.fmt) x.numFmt = o.fmt;
    return x;
  };
  const alto = (fila, pt) => { ws.getRow(fila).height = pt; };
  const imagen = (b64, fila, w, h, ext = 'png') => {
    const id = wb.addImage({ base64: b64, extension: ext });
    ws.addImage(id, { tl: { nativeCol: 0, nativeColOff: 0, nativeRow: fila - 1, nativeRowOff: 0 }, ext: { width: w, height: h } });
  };
  const PAGINA = 850;                                               // alto de cada hoja impresa, en puntos

  const hoy = (typeof hoyBogota === 'function' ? hoyBogota() : new Date().toISOString().slice(0, 10)).split('-').map(Number);
  const fmtCop = '"$" #,##0.00';

  /* ===== Hoja 1 ===== */
  let f = 1, usado = 0;
  const fila = pt => { alto(f, pt); usado += pt; return f++; };
  imagen(COT_GPK_CABECERA, fila(75), ANCHO, 100);
  fila(14);
  [`Estimado Cliente: ${String(c.cliente || '').trim()}`, 'A continuación nos permitimos enviar respuesta de su solicitud:'].forEach(t => {
    const n = fila(17); ws.mergeCells(`A${n}:E${n}`); poner(`A${n}`, t, { tam: 11, indent: 5, ajustar: false });
  });
  fila(8);
  ['GROUPACK COLOMBIA', 'CALI'].forEach(t => {
    const n = fila(20); ws.mergeCells(`A${n}:E${n}`); poner(`A${n}`, t, { tam: 12, h: 'center', relleno: AMARILLO, caja: true });
    for (const col of 'BCDE') ws.getCell(`${col}${n}`).border = caja;
  });
  fila(14);
  const tiempo = aereo ? '15 días' : '60 días';
  const cod = `COT-${String(c.id || '').slice(0, 5).toUpperCase()}`;
  const izq = [['Condiciones de Pago', 'Contado'], ['EMPRESA', 'Groupack'], ['Contacto', ''], ['Modalidad', aereo ? 'Aereo' : 'Maritimo']];
  const der = [['No. Cotización', cod], ['Fecha', new Date(Date.UTC(hoy[0], hoy[1] - 1, hoy[2]))], ['Validez', ''], ['Tiempo de entrega estimado', tiempo]];
  const PRIMERA = f;
  izq.forEach((par, i) => {
    const n = fila(18);
    poner(`A${n}`, par[0], { caja: true }); poner(`B${n}`, par[1], { caja: true });
    poner(`D${n}`, der[i][0], { relleno: AMARILLO, caja: true });
    poner(`E${n}`, der[i][1], { caja: true, fmt: 'dd/mm/yyyy', h: 'left' });
  });
  fila(14);
  const ARRIBA_IMG = f;                                             // la imagen del producto ocupa el espacio de la izquierda
  const desc = String(c.descripcion || '');
  const nDesc = fila(Math.max(90, Math.ceil(desc.length / 26) * 13 + 10));
  poner(`D${nDesc}`, 'Descripción del Producto', { relleno: AMARILLO, caja: true, h: 'center' });
  poner(`E${nDesc}`, desc, { caja: true, tam: 9 });
  const num = v => (cotOk(v) ? v : null);
  const campo = (etq, val, fmt, o = {}) => {
    const n = fila(18);
    poner(`D${n}`, etq, { relleno: o.negro ? NEGRO : AMARILLO, color: o.negro ? 'FFFFFFFF' : NEGRO, caja: true });
    poner(`E${n}`, val, { caja: true, fmt, h: 'left' });
    return n;
  };
  const nCant = campo('Cantidad por comprar', num(r.unidades), '#,##0');
  campo('Cantidad por caja', num(k.unidadesPorCaja), '#,##0');
  campo('No. De cajas', num(r.cajas), '#,##0');
  campo('CBM/Caja', num(k.cbmCaja), '0.0000');
  campo('Peso/Caja', num(k.pesoCaja), '#,##0.00" kg"');
  fila(14);
  // Valor de la unidad (con aumento, sin comisión) y comisión, en COP por unidad
  let sinCom = null, com = null;
  if (cotOk(r.precioUnitCop) && cotOk(k.valorUnidad) && cotOk(k.aumentoRmb) && k.usdRmb > 0 && cotOk(k.usdCop)) {
    const unitAum = k.aumentoTipo === 'porcentaje' ? k.valorUnidad * (1 + k.aumentoRmb / 100) : k.valorUnidad + k.aumentoRmb;
    sinCom = unitAum / k.usdRmb * k.usdCop;
    com = r.precioUnitCop - sinCom;
  }
  const pct = r.comisionPct != null ? String(Math.round(r.comisionPct * 100) / 100).replace('.', ',') : '';
  const nVu = campo('Valor Unidad Producto', sinCom, fmtCop);
  const nCom = campo(`Comisión Groupack${pct ? ' ' + pct + '%' : ''}`, com, fmtCop);
  fila(14);
  const nTu = campo('Valor Total Unidad Producto', sinCom != null ? { formula: `E${nVu}+E${nCom}`, result: sinCom + com } : num(r.precioUnitCop), fmtCop);
  const nFl = campo('Valor Flete por Producto', num(r.logisticaUnit), fmtCop);
  const tuVal = num(r.precioUnitCop), flVal = num(r.logisticaUnit);
  const unit360 = tuVal != null && flVal != null ? tuVal + flVal : num(r.totalUnitCop);
  const n360 = campo('Valor Unitario Servicio 360', tuVal != null && flVal != null ? { formula: `E${nTu}+E${nFl}`, result: unit360 } : unit360, fmtCop, { negro: true });
  fila(14);
  campo('Valor Total a Pagar', unit360 != null && num(r.unidades) != null ? { formula: `E${n360}*E${nCant}`, result: unit360 * r.unidades } : num(r.totalConFlete), fmtCop, { negro: true });
  const ABAJO_IMG = f;
  // imagen del producto (real si existe; si no, la de referencia), en el espacio libre de la izquierda
  const src = (c.imagenesReales || []).length ? c.imagenesReales[0] : (c.imagenes || [])[0];
  const d = src ? await cotDimensionesImagen(src) : null;
  if (d) {
    let hDisp = 0; for (let i = ARRIBA_IMG; i < ABAJO_IMG; i++) hDisp += ws.getRow(i).height * 4 / 3;
    const e = Math.min(360 / d.w, (hDisp - 20) / d.h, 1);
    const w = d.w * e, h = d.h * e;
    const idImg = wb.addImage({ base64: src, extension: /^data:image\/png/.test(src) ? 'png' : 'jpeg' });
    ws.addImage(idImg, { tl: { nativeCol: 0, nativeColOff: Math.round(Math.max(0, (PX[0] + PX[1] - w) / 2) * 9525), nativeRow: ARRIBA_IMG - 1, nativeRowOff: Math.round(10 * 9525) }, ext: { width: w, height: h } });
  }
  // relleno hasta el pie de la hoja y pie
  fila(Math.max(8, PAGINA - usado - 86));
  imagen(COT_GPK_PIE, fila(86), ANCHO, 115);
  ws.getRow(f - 1).addPageBreak();

  /* ===== Hoja 2 ===== */
  usado = 0;
  imagen(COT_GPK_CABECERA, fila(75), ANCHO, 100);
  fila(40);
  { const n = fila(18); ws.mergeCells(`A${n}:E${n}`); poner(`A${n}`, 'OBSERVACIONES', { tam: 11, indent: 5, ajustar: false }); }
  fila(14);
  const obs = [
    'Cotización no incluye Fulfillment',
    'Confirmar aprobación de la cotización',
    aereo ? 'El tiempo de entrega Aerea es a partir de 15 dias que la mercancía se encuentre en la bodega en origen'
          : 'El tiempo de entrega Maritima es a partir de 60 dias que la mercancía se encuentre en la bodega en origen',
    ...(aereo ? ['Flete aéreo sujeto a impuestos'] : []),
    'Valor es hasta la ciudad de Cali',
    'Se calcula la tasa del dolar del día que realice el pago',
    'Valor sujeto a cambio de acuerdo con el comportamiento del flete'
  ];
  obs.forEach(t => { const n = fila(t.length > 70 ? 32 : 18); ws.mergeCells(`A${n}:E${n}`); poner(`A${n}`, '➢   ' + t, { indent: 8 }); });
  fila(60);
  { const n = fila(18); ws.mergeCells(`A${n}:E${n}`); poner(`A${n}`, 'Cordialmente,', { tam: 11, indent: 5, ajustar: false }); }
  fila(70);
  ['Libardo González', 'Groupack sas', 'Colombia'].forEach(t => { const n = fila(17); ws.mergeCells(`A${n}:E${n}`); poner(`A${n}`, t, { tam: 11, indent: 5, ajustar: false }); });
  fila(Math.max(8, PAGINA - usado - 86));
  imagen(COT_GPK_PIE, fila(86), ANCHO, 115);

  ws.pageSetup = { paperSize: 1, orientation: 'portrait', scale: 84, horizontalCentered: true,
    margins: { left: 0.4, right: 0.4, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 }, printArea: `A1:E${f - 1}` };

  const buf = await wb.xlsx.writeBuffer();
  const url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  const a = document.createElement('a');
  const slug = String(c.cliente || 'cliente').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '').toLowerCase() || 'cliente';
  a.href = url; a.download = `cotizacion-groupack-${slug}-${hoy.join('')}.xlsx`;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
  return true;
}

/* ---------- Mandar a módulo (Adma Company / Groupack) y pagos ----------
   Al mandar una cotización a su módulo se guarda: modulo ('adma-company' | 'groupack'), estado 'finalizada' y pagos: [].
   Cada pago: { id, concepto: 'producto'|'flete'|'todo', monto, montoProducto, montoFlete, fecha, nota, comprobante, por, registradoEn }
   El estado de pago (sin pagar / abonado / pagado) se calcula siempre a partir de los pagos y de los totales de la cotización. */
const COT_MODULO_DE_ORIGEN = { 'Adma Company': 'adma-company', 'Groupack': 'groupack' };
const cotNombreModulo = id => (APP_CONFIG.MODULOS.find(m => m.id === id) || {}).nombre || id;
const PG_ESTADOS = { sinpagar: 'Sin pagar', abonado: 'Abonado', pagado: 'Pagado' };
const PG_CONCEPTOS = { producto: 'Producto', flete: 'Flete', todo: 'Producto + flete' };
let pgDetalleId = null;
let pgPagoId = null;
const pgEsSuper = () => !!sesionActual && sesionActual.rol === APP_CONFIG.ROLES.SUPER;
const pgCop = v => '$ ' + cotNum(v, 0) + ' COP';
const pgSuma = (lista, campo) => lista.reduce((t, p) => t + (Number(p[campo]) || 0), 0);

/* Mandar a módulo: primero se piden las observaciones (tallas, cantidades por talla, etc.), que van en la columna NOTES de la orden de compra.
   La cotización se guarda con `observaciones` (texto, puede tener varias líneas). Se pueden editar mientras no tenga orden de compra generada. */
const COT_OBS_MAX = 400;
let cotObsId = null, cotObsModo = 'enviar';
const cotObsHtml = (c, attr) => `<p class="cot-dato cot-obs"><strong>Observaciones:</strong> ${c.observaciones ? cotEsc(c.observaciones) : '<em>Sin observaciones</em>'}${attr && !cotTieneOrden(c) ? ` <button type="button" class="cot-enlace" ${attr}="${c.id}">Editar</button>` : ''}</p>`;

function cotMandarAModulo(id) {
  const c = cotLista.find(x => x.id === id);
  if (!c) return;
  if (!COT_MODULO_DE_ORIGEN[c.origen]) { alert('Esta cotización no tiene indicado si el cliente viene de Adma Company o Groupack, por eso no se puede mandar a un módulo.'); return; }
  cotAbrirObs(id, 'enviar');
}

function cotAbrirObs(id, modo) {
  const c = cotLista.find(x => x.id === id);
  if (!c) return;
  cotObsId = id; cotObsModo = modo;
  document.getElementById('cotFormObs').reset();
  cotError('cotErrorObs', '');
  const nombre = cotNombreModulo(COT_MODULO_DE_ORIGEN[c.origen] || '');
  document.getElementById('cotObsTitulo').textContent = (modo === 'enviar' ? 'Mandar a módulo · ' : 'Observaciones · ') + c.cliente;
  document.getElementById('cotObsResumen').textContent = modo === 'enviar'
    ? `Se mandará al módulo ${nombre} y la cotización quedará Finalizada.` : '';
  document.getElementById('cotObsTexto').value = c.observaciones || '';
  document.getElementById('cotBtnGuardarObs').textContent = modo === 'enviar' ? `Mandar a ${nombre}` : 'Guardar';
  cotAbrir('cotModalObs');
  document.getElementById('cotObsTexto').focus();
}

async function cotGuardarObs(e) {
  e.preventDefault();
  const c = cotLista.find(x => x.id === cotObsId);
  if (!c) return cotError('cotErrorObs', 'No se encontró la cotización.');
  const observaciones = document.getElementById('cotObsTexto').value.trim();
  const btn = document.getElementById('cotBtnGuardarObs');
  btn.disabled = true;
  try {
    if (cotObsModo === 'editar') {
      if (cotTieneOrden(c)) return cotError('cotErrorObs', 'Esta cotización ya tiene orden de compra: las observaciones ya no se pueden cambiar.');
      await cotRef().doc(c.id).update({ observaciones });
    } else {
      const mod = COT_MODULO_DE_ORIGEN[c.origen];
      if (!mod) return cotError('cotErrorObs', 'Esta cotización no tiene indicado el origen del cliente.');
      const datos = {
        estado: 'finalizada', modulo: mod, observaciones,
        enviadaEn: firebase.firestore.FieldValue.serverTimestamp(),
        enviadaPor: sesionActual ? sesionActual.nombre : ''
      };
      if (!Array.isArray(c.pagos)) datos.pagos = [];
      if (!c.finalizadaEn) datos.finalizadaEn = firebase.firestore.FieldValue.serverTimestamp();
      await cotRef().doc(c.id).update(datos);
    }
    cotCerrarModal(document.getElementById('cotModalObs'));
  } catch (err) {
    console.error('Error al guardar las observaciones:', err);
    cotError('cotErrorObs', 'No se pudo guardar. Inténtalo de nuevo.');
  } finally { btn.disabled = false; }
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
  return { viaAerea: !!r.viaAerea, unidades: r.unidades, producto, flete, fleteTotal, esGroupack, parteAdma, total: producto + flete, pagadoP, pagadoF, pendP, pendF, pagado, pendiente, estado };
}
const pgViaTxt = R => R.viaAerea ? 'aéreo' : 'marítimo';
const pgFleteTxt = R => 'Flete ' + pgViaTxt(R);
const pgConceptoTxt = (concepto, R) => concepto === 'flete' ? pgFleteTxt(R) : concepto === 'todo' ? 'Producto + ' + pgFleteTxt(R).toLowerCase() : 'Producto';
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
        <input type="search" data-pg-buscar placeholder="Buscar por cliente, flete o estado de pago…" aria-label="Buscar en ${cotEsc(cotNombreModulo(mod))}" autocomplete="off">
        <select data-pg-estado aria-label="Filtrar por estado de pago"><option value="">Todos los pagos</option><option value="sinpagar">Sin pagar</option><option value="abonado">Abonado</option><option value="pagado">Pagado</option></select>
      </div>
      <p class="cot-conteo" data-pg-conteo aria-live="polite"></p></div>
    <div class="cot-tabla-wrap" data-pg-grid aria-live="polite"><p class="lista-vacia">Cargando…</p></div>`;
  sec.addEventListener('input', e => { if (e.target.closest('[data-pg-buscar]')) pgPintar(mod); });
  sec.addEventListener('change', e => { if (e.target.closest('[data-pg-estado]')) pgPintar(mod); });
  sec.addEventListener('click', e => {
    const pag = e.target.closest('[data-pg-pagar]');
    if (pag) { pgAbrirPago(pag.dataset.pgPagar); return; }
    if (e.target.closest('input, a, button, select')) return;
    const fila = e.target.closest('tr[data-pg-toggle]');
    if (fila) pgAbrirDetalle(fila.dataset.pgToggle);
  });
  sec.addEventListener('keydown', e => {
    if ((e.key === 'Enter' || e.key === ' ') && e.target.matches('tr[data-pg-toggle]')) { e.preventDefault(); pgAbrirDetalle(e.target.dataset.pgToggle); }
  });
}

/* Ventana de detalle (compartida por Adma Company y Groupack) */
function pgInitDetalle() {
  document.getElementById('pgDetalleCuerpo').addEventListener('click', e => {
    const pag = e.target.closest('[data-pg-pagar]');
    const comp = e.target.closest('[data-pg-comp]');
    const bor = e.target.closest('[data-pg-borrar]');
    const parte = e.target.closest('[data-pg-parte-guardar]');
    const obs = e.target.closest('[data-pg-obs]');
    if (obs) { cotCerrarModal(document.getElementById('pgModalDetalle')); cotAbrirObs(obs.dataset.pgObs, 'editar'); }
    else if (parte) pgGuardarParteAdma(parte.dataset.pgParteGuardar);
    else if (pag) pgAbrirPago(pag.dataset.pgPagar);
    else if (comp) pgVerComprobante(comp.dataset.pgComp, comp.dataset.pid);
    else if (bor) pgQuitarPago(bor.dataset.pgBorrar, bor.dataset.pid);
  });
}

function pgAbrirDetalle(id) {
  pgDetalleId = id;
  pgPintarDetalle();
  if (pgDetalleId) cotAbrir('pgModalDetalle');
}

function pgPintarDetalle() {
  const modal = document.getElementById('pgModalDetalle');
  const c = cotLista.find(x => x.id === pgDetalleId);
  const R = c && pgResumen(c);
  if (!R) { pgDetalleId = null; if (modal) modal.hidden = true; return; }
  const fotos = c.imagenesReales || [];
  const pagos = (c.pagos || []).slice().sort((a, b) => (a.registradoEn || 0) - (b.registradoEn || 0));
  const lista = pagos.map(p => `<li class="pg-pago"><span>${cotEsc(p.fecha || '')} · ${cotEsc(pgConceptoTxt(p.concepto, R))} · <strong>${pgCop(p.monto)}</strong>${p.nota ? ` · ${cotEsc(p.nota)}` : ''}${p.por ? ` · ${cotEsc(p.por)}` : ''}</span>
      ${p.comprobante ? `<button type="button" class="cot-enlace" data-pg-comp="${c.id}" data-pid="${cotEsc(p.id)}">Comprobante</button>` : ''}
      ${pgEsSuper() ? `<button type="button" class="cot-eliminar" data-pg-borrar="${c.id}" data-pid="${cotEsc(p.id)}" aria-label="Quitar este pago">Quitar</button>` : ''}</li>`).join('');
  document.getElementById('pgDetalleTitulo').textContent = c.cliente;
  document.getElementById('pgDetalleCuerpo').innerHTML = `
    <p><span class="cot-estado pg-${R.estado}">${PG_ESTADOS[R.estado]}</span></p>
    <p class="cot-dato"><strong>Cliente:</strong> ${cotEsc(c.cliente)}</p>
    ${fotos.length ? `<div class="cot-miniaturas">${fotos.map((s, i) => `<img src="${cotEsc(s)}" alt="Foto ${i + 1} de ${cotEsc(c.cliente)}">`).join('')}</div>` : ''}
    <p class="cot-dato"><strong>Cantidad cotizada:</strong> ${cotNum(R.unidades, 0)}</p>
    ${cotObsHtml(c, 'data-pg-obs')}
    <p class="cot-dato"><strong>Precio total del producto:</strong> ${pgCop(R.producto)}</p>
    <p class="cot-dato"><strong>Tipo de flete:</strong> ${R.viaAerea ? 'Aéreo' : 'Marítimo'}</p>
    <p class="cot-dato pg-flete-fila"><span><strong>Precio total del ${pgFleteTxt(R).toLowerCase()}:</strong> ${pgCop(R.fleteTotal)}</span>
      ${R.esGroupack ? `<span class="pg-parte"><strong>Parte correspondiente del flete a Adma:</strong>
        <input type="number" min="0" step="any" data-pg-parte-input="${c.id}" value="${R.parteAdma ?? ''}" placeholder="Valor en COP" aria-label="Parte correspondiente del flete a Adma en COP">
        <button type="button" class="btn btn-chico" data-pg-parte-guardar="${c.id}">Guardar</button></span>` : ''}</p>
    <p class="cot-dato"><strong>Precio total:</strong> ${pgCop(R.total)}</p>
    <p class="cot-costos"><strong>PAGOS</strong><br>
      <strong>Producto:</strong> pagado ${pgCop(R.pagadoP)} · falta ${pgCop(R.pendP)}<br>
      <strong>${pgFleteTxt(R)}:</strong> pagado ${pgCop(R.pagadoF)} · falta ${pgCop(R.pendF)}<br>
      <strong>Total:</strong> pagado ${pgCop(R.pagado)} · falta ${pgCop(R.pendiente)}</p>
    ${lista ? `<ul class="pg-pagos">${lista}</ul>` : ''}
    ${R.estado === 'pagado' ? '' : `<footer><button type="button" class="btn btn-chico" data-pg-pagar="${c.id}">Registrar pago</button></footer>`}`;
}

function pgFila(c) {
  const R = pgResumen(c);
  if (!R) return '';
  const foto = (c.imagenesReales || [])[0];
  return `<tr class="cot-fila" data-pg-toggle="${c.id}" tabindex="0" aria-haspopup="dialog" title="Ver detalle">
    <td class="cot-col-cliente"><div class="pg-cliente">${foto ? `<img class="pg-thumb" src="${cotEsc(foto)}" alt="">` : '<span class="pg-thumb" aria-hidden="true"></span>'}<strong>${cotEsc(c.cliente)}</strong></div></td>
    <td><span class="cot-via cot-via-${R.viaAerea ? 'aereo' : 'maritimo'}">${R.viaAerea ? 'Aéreo' : 'Marítimo'}</span></td>
    <td class="cot-num"><span class="cot-precio">${R.esGroupack && R.parteAdma == null ? '<span class="cot-alerta" title="Aún no se ha registrado la parte correspondiente del flete a Adma. Ábrela y escríbela en la ventana de detalle.">⚠ Falta parte del flete</span>' : ''}<span>${pgCop(R.total)}</span></span></td>
    <td class="cot-num">${pgCop(R.pendiente)}</td>
    <td><span class="cot-estado pg-${R.estado}">${PG_ESTADOS[R.estado]}</span></td>
    <td class="cot-col-acc"><div class="cot-acciones-fila cot-acciones-pg">${R.estado === 'pagado' ? '<span aria-hidden="true"></span>'
      : `<button type="button" class="cot-btn-tabla cot-btn-modulo" data-pg-pagar="${c.id}" title="Registrar pago">Registrar pago</button>`}</div></td>
  </tr>`;
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
    const pajar = cotNorm(c.cliente + ' ' + PG_ESTADOS[R.estado] + ' ' + (R.viaAerea ? 'aereo' : 'maritimo'));
    return q.every(t => pajar.includes(t));
  });
  conteo.textContent = lista.length === todas.length ? `${todas.length} cotizaciones` : `${lista.length} de ${todas.length} cotizaciones`;
  grid.innerHTML = lista.length ? `<div class="cot-tabla-card"><table class="cot-tabla">
    <thead><tr><th>Cliente</th><th>Tipo de flete</th><th class="cot-num">Precio total para Adma</th><th class="cot-num">Falta por pagar</th><th>Estado de pago</th><th class="cot-col-acc">Acciones</th></tr></thead>
    <tbody>${lista.map(pgFila).join('')}</tbody></table></div>` : '<p class="lista-vacia">Ninguna cotización coincide con la búsqueda.</p>';
}
const pgPintarTodos = () => {
  Object.values(COT_MODULO_DE_ORIGEN).forEach(pgPintar);
  if (pgDetalleId) pgPintarDetalle();
  ocPintar();
  if (ocDetalleBooth != null) ocPintarDetalle();
};

/* ---------- Groupack: parte correspondiente a Adma ---------- */
async function pgGuardarParteAdma(id) {
  const input = document.querySelector(`#pgDetalleCuerpo [data-pg-parte-input="${id}"]`);
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
    ? `Falta por pagar · Producto: ${pgCop(R.pendP)} · ${pgFleteTxt(R)}: ${pgCop(R.pendF)} · Total: ${pgCop(R.pendP + R.pendF)}` : '';
}

function pgAbrirPago(id) {
  const c = cotLista.find(x => x.id === id);
  if (!c) return;
  pgPagoId = id;
  document.getElementById('cotFormPago').reset();
  cotError('cotErrorPago', '');
  document.getElementById('cotPagoTitulo').textContent = 'Registrar pago · ' + c.cliente;
  document.getElementById('cotPagoFecha').value = typeof hoyBogota === 'function' ? hoyBogota() : '';
  const Rp = pgResumen(c);
  if (Rp) {
    document.querySelector('#cotPagoConcepto option[value="flete"]').textContent = 'Precio del ' + pgFleteTxt(Rp).toLowerCase();
    document.querySelector('#cotPagoConcepto option[value="todo"]').textContent = 'Todo (producto + ' + pgFleteTxt(Rp).toLowerCase() + ')';
  }
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
  if (max <= 1) return cotError('cotErrorPago', `Ya no falta nada por pagar en "${pgConceptoTxt(concepto, R)}".`);
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
    cotCerrarModal(document.getElementById('cotModalPago'));
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
  if (!p || !/^data:image\/(png|jpe?g|webp|gif);base64,[A-Za-z0-9+\/=]+$/.test(p.comprobante || '')) return;
  const w = window.open('', '_blank');
  if (!w) { alert('El navegador bloqueó la ventana. Permite las ventanas emergentes e inténtalo de nuevo.'); return; }
  w.document.write(`<title>Comprobante</title><body style="margin:0;background:#222"><img src="${cotEsc(p.comprobante)}" style="max-width:100%;display:block;margin:auto">`);
  w.document.close();
}

/* ---------- Escucha en tiempo real ---------- */
function initEscuchaCotizaciones() {
  if (!document.getElementById('cotGrid')) return;
  detenerEscuchaCotizaciones();
  cotDesuscribir = cotRef().orderBy('creadoEn', 'desc').onSnapshot(snap => {
    cotLista = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    // Si una cotización se borra, sale de la selección
    [...cotSeleccion].forEach(id => { if (!cotLista.some(c => c.id === id)) cotSeleccion.delete(id); });
    cotPintar();
    if (cotDetalleId) cotPintarDetalle();
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

/* ---------- Orden de compra: réplica de la plantilla de Excel (ORDEN_DE_COMPRA_PLANTILLA.xlsx) ----------
   Se dibuja en un canvas con las mismas medidas de columnas/filas de la plantilla (así se ven igual los textos en chino
   con las fuentes del dispositivo) y se inserta como página horizontal en el PDF. */
const OC_LOGO = 'data:image/png;base64,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gajQMXjD58ViqcPdK0fq3KYRKPkpAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIIIAAAggggAACCCCAAAIIdEHg/wGQEm7LPw2QogAAAABJRU5ErkJggg==';
const OC_FUENTE = 'Arial, "Microsoft YaHei", "PingFang SC", "Noto Sans CJK SC", "SimSun", "Songti SC", sans-serif';
const OC_COLS = [9, 27.85546875, 17.140625, 14.42578125, 19.42578125, 11.28515625, 9, 9, 9, 9, 9, 13.42578125, 9, 9, 15.140625, 22]; // A..P (la P, NOTES, se ensanchó de 9 a 22 para que quepan tallas y cantidades)
const OC_COLS_ENC = [['NO.', '序号'], ['PICTURE', '产品照片'], ['SHIPPING MARK'], ['ITEM NO', '货号'], ['DESCRIPTION', 'SPANISH', '品名及规格'],
  ['DESCRIPTION', 'CHINESE', '品名及规格'], ['CTN', '箱数'], ['QTY/', 'CTN', '装箱数量'], ['UNIT', '单位'], ['QTY', '总数量'], ['PRICE', '单价'],
  ['AMOUNT', '金额'], ['CBM', '箱规'], ['T.CBM', '总体积'], ['G.W', '毛重'], ['NOTES', '备注']];

function ocCargarImagen(src) {
  return new Promise(ok => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => ok(null); i.src = src; });
}

/* items: [{ foto, descripcion, ctn, qtyCtn, price, cbm, pesoCaja }] -> canvas con la página completa */
async function ocDibujar(items, fechaTxt, booth) {
  const S = 2, PX = 96 / 72, nItems = Math.max(5, items.length);
  const colsPx = OC_COLS.map(w => w * 7);
  const filasPt = [101.1, 20.1, 20.1, 20.1, 20.1, 65.1, ...Array(nItems).fill(96), 65.25, 15, 15, 14.25, 14.25, 14.25, 14.25, 14.25, 14.25, 14.25, 14.25, 20.25];
  const xs = [0]; colsPx.forEach(w => xs.push(xs[xs.length - 1] + w));
  const ys = [0]; filasPt.forEach(h => ys.push(ys[ys.length - 1] + h * PX));
  const W = xs[xs.length - 1] + 8, H = ys[ys.length - 1];
  const T = 6 + nItems;                    // fila (0-based) de totales
  const cv = document.createElement('canvas');
  cv.width = Math.ceil(W * S); cv.height = Math.ceil(H * S);
  const g = cv.getContext('2d');
  g.scale(S, S);
  g.fillStyle = '#fff'; g.fillRect(0, 0, W, H);
  g.textBaseline = 'alphabetic';

  const R = (c0, c1, r0, r1) => ({ x: xs[c0], y: ys[r0], w: xs[c1 + 1] - xs[c0], h: ys[r1 + 1] - ys[r0] });
  const fuente = (pt, b) => `${b ? 'bold ' : ''}${pt * PX}px ${OC_FUENTE}`;
  const partir = (txt, maxW) => {            // ajuste de línea: por palabras y, si una palabra no cabe, por caracteres
    const out = [];
    String(txt).split('\n').forEach(parr => {
      let linea = '';
      (parr.match(/[^\s-]+-?|\s+/g) || []).forEach(tok => {   // corta por espacios y después de cada guion
        if (!tok) return;
        if (g.measureText(linea + tok).width <= maxW) { linea += tok; return; }
        if (linea.trim()) { out.push(linea.trimEnd()); linea = ''; }
        if (g.measureText(tok.trimStart()).width <= maxW) { linea = tok.trimStart(); return; }
        for (const ch of tok.trimStart()) {
          if (g.measureText(linea + ch).width > maxW && linea) { out.push(linea); linea = ''; }
          linea += ch;
        }
      });
      out.push(linea.trimEnd());
    });
    return out;
  };
  /* Texto dentro de un rectángulo: align left|center|right, valign top|middle|bottom, wrap y recorte opcionales */
  const txt = (r, t, o = {}) => {
    const { pt = 11, b = false, color = '#000', align = 'left', valign = 'middle', wrap = false, clip = false, pad = 3 } = o;
    g.save();
    if (clip || wrap) { g.beginPath(); g.rect(r.x, r.y, r.w, r.h); g.clip(); }
    g.font = fuente(pt, b); g.fillStyle = color;
    const lh = pt * PX * 1.2;
    const lineas = wrap ? partir(t, r.w - pad * 2) : String(t).split('\n');
    const total = lineas.length * lh;
    let y0 = valign === 'top' ? r.y + 2 : valign === 'bottom' ? r.y + r.h - total - 2 : r.y + (r.h - total) / 2;
    lineas.forEach((l, i) => {
      const w = g.measureText(l).width;
      const x = align === 'center' ? r.x + (r.w - w) / 2 : align === 'right' ? r.x + r.w - w - pad : r.x + pad;
      g.fillText(l, x, y0 + i * lh + lh * 0.78);
    });
    g.restore();
  };
  const caja = r => { g.strokeStyle = '#000'; g.lineWidth = 1; g.strokeRect(r.x + 0.5, r.y + 0.5, r.w - 1, r.h - 1); };

  /* --- Encabezado --- */
  const logo = await ocCargarImagen(OC_LOGO);
  if (logo) g.drawImage(logo, 16, 7, 134, 116);
  txt(R(3, 14, 0, 0), 'GROUPACK\nYIWU ADSUKA TRADING CO.,LIMITED\n义乌市阿肃咔贸易有限公司', { pt: 16, b: true, align: 'center' });
  txt(R(15, 15, 0, 0), 'NO:', { pt: 12, valign: 'bottom' });
  txt(R(0, 4, 1, 1), 'BOOTH(店面）: ' + booth, { pt: 10 });
  txt(R(0, 4, 2, 2), 'TEL(电话）：', { pt: 10 });
  txt(R(0, 4, 3, 3), 'RECEIVED BY(收货人）：', { pt: 10 });
  txt(R(0, 4, 4, 4), 'PAYMENT TIME(付款时间）：现金', { pt: 10 });
  txt(R(5, 12, 2, 2), 'ORDER DATE(订货日期）：  ' + fechaTxt, { pt: 10, clip: true });
  txt(R(5, 12, 3, 3), 'DELIVERY DATE(交货日期）：        年      月      日', { pt: 10, clip: true });
  txt(R(5, 12, 4, 4), 'DELIVERY ADD(交货地址）：浙江省义乌市商博路518号(Chinagoods)共享云仓7号1单元  ', { pt: 11, clip: true });
  txt(R(13, 15, 1, 1), '地址（add）：浙江省义乌市稠州北路1121号', { pt: 10, wrap: true });
  txt(R(14, 15, 2, 2), '      福田大厦A座2307室', { pt: 10, wrap: true });
  txt(R(13, 15, 3, 3), '电话：', { pt: 10 });
  txt(R(13, 15, 4, 4), '工作时间：周一到周六9：00-17:00', { pt: 10 });

  /* --- Tabla --- */
  for (let c = 0; c < 16; c++) {
    const r = R(c, c, 5, 5); caja(r);
    const lin = OC_COLS_ENC[c];
    txt(r, lin.join('\n'), { pt: 11, align: 'center', wrap: true, pad: 1 });
  }
  const fmt2 = v => Number(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const fmtN = (v, d) => String(+Number(v).toFixed(d));
  let sCtn = 0, sPrice = 0, sAmt = 0, sCbm = 0;
  for (let i = 0; i < nItems; i++) {
    const f = 6 + i, it = items[i];
    for (let c = 0; c < 16; c++) caja(R(c, c, f, f));
    txt(R(0, 0, f, f), String(i + 1), { align: 'center' });
    if (!it) continue;
    const qty = it.ctn * it.qtyCtn, amt = it.price * qty, tcbm = it.cbm * it.ctn;
    sCtn += it.ctn; sPrice += it.price; sAmt += amt; sCbm += tcbm;
    const ct = (c, t, o = {}) => txt(R(c, c, f, f), t, { align: 'center', ...o });
    const fotoImg = it.foto ? await ocCargarImagen(it.foto) : null;
    if (fotoImg) {
      const rb = R(1, 1, f, f), pad = 8, k = Math.min((rb.w - pad * 2) / fotoImg.width, (rb.h - pad * 2) / fotoImg.height, 1.6);
      const w = fotoImg.width * k, h = fotoImg.height * k;
      g.drawImage(fotoImg, rb.x + (rb.w - w) / 2, rb.y + (rb.h - h) / 2, w, h);
    }
    ct(2, it.marca || '', { align: 'left', wrap: true });
    ct(3, it.itemNo || '', { align: 'left', pt: 10, wrap: true });
    ct(4, it.descripcion, { wrap: true });
    ct(6, fmtN(it.ctn, 0)); ct(7, fmtN(it.qtyCtn, 0)); ct(8, 'PCS'); ct(9, fmtN(qty, 0));
    ct(10, '￥' + fmt2(it.price)); txt(R(11, 11, f, f), '￥' + fmt2(amt), { align: 'right', pt: 11 });
    ct(12, fmtN(it.cbm, 5)); ct(13, tcbm.toFixed(3));
    if (it.pesoCaja != null) ct(14, fmtN(it.pesoCaja * it.ctn, 2));
    if (it.notas) {                                   // observaciones: se reduce la letra hasta que el texto quepa en la celda
      const rn = R(15, 15, f, f);
      let pt = 10;
      for (const p of [10, 9, 8, 7, 6.5, 6]) {
        pt = p; g.font = fuente(p);
        if (partir(it.notas, rn.w - 6).length * p * PX * 1.2 <= rn.h - 4) break;
      }
      txt(rn, it.notas, { pt, wrap: true, valign: 'top', align: 'left' });
    }
  }
  for (let c = 0; c < 16; c++) caja(R(c, c, T, T));
  [6, 11, 13].forEach(c => { const r = R(c, c, T, T); g.fillStyle = '#FFFF00'; g.fillRect(r.x + 1, r.y + 1, r.w - 2, r.h - 2); caja(r); });
  txt(R(6, 6, T, T), fmtN(sCtn, 0), { pt: 11, b: true, align: 'right' });
  txt(R(10, 10, T, T), '￥' + fmt2(sPrice), { align: 'center' });
  txt(R(11, 11, T, T), '￥' + fmt2(sAmt), { pt: 11, b: true, align: 'right' });
  txt(R(13, 13, T, T), sCbm.toFixed(3), { pt: 11, b: true, align: 'right' });

  /* --- Total, firmas y notas --- */
  const rTot = R(0, 15, T + 1, T + 1); caja(rTot);
  const base = 'Total Amount（总金额）：   拾   万   仟   佰   拾   元           ¥：';
  txt(rTot, base, { pt: 10 });
  g.font = fuente(10); txt({ x: rTot.x + g.measureText(base).width + 8, y: rTot.y, w: 400, h: rTot.h }, fmt2(sAmt), { pt: 10, b: true });
  const rFir = R(0, 15, T + 2, T + 2); caja(rFir);
  txt(rFir, '采购方签名（Buyer Signature）：                            供货方签名（Seller Signatiue）：', { pt: 10 });
  txt(R(0, 15, T + 3, T + 3), '注意事项：', { pt: 12 });
  const notas = ['1.质量要求技术标准：产品品质、规格应完全与样品及合同要求相符.否则拒绝收货', '2.买方已付订金或订单被卖方取消时，卖方需返还三倍订金作为买方损失.',
    '3.订单的任何变更需取得买方经理的认可.未经买方经理认可，对订单的修改将不被买方接受.', '4.供方必须按照买方签订的时间准时交货.若不能如期交货，一切责任由供货方承担.', '5.货品要保质、保量、不良货品可以退掉.'];
  notas.forEach((n, i) => txt(R(3, 15, T + 4 + i, T + 4 + i), n, { pt: 8 }));
  txt(R(10, 13, T + 4, T + 4), '正唛和侧唛必须两正两侧', { pt: 10, color: '#FF0000' });
  const lab = [['ITEM NO:', '填货号'], ['QTY:', '     PCS'], ['G.W:', '     KG'], ['N.W:', '     KG'], ['MEAS:', '   X  X  CM']];
  lab.forEach(([a, b], i) => {
    txt(R(14, 14, T + 4 + i, T + 4 + i), a, { pt: 12, align: 'right' });
    txt(R(15, 15, T + 4 + i, T + 4 + i), b, { pt: i === 0 ? 11 : 12, color: i === 0 ? '#FF0000' : '#000' });
  });
  // Rombo de marcas (de la mitad de la columna J hasta la N, filas T+5 a T+10)
  const dx0 = xs[9] + 56.3, dy0 = ys[T + 5] + 2.7, dx1 = xs[13] + 16.3, dy1 = ys[T + 10] + 12;
  g.beginPath(); g.moveTo((dx0 + dx1) / 2, dy0); g.lineTo(dx1, (dy0 + dy1) / 2); g.lineTo((dx0 + dx1) / 2, dy1); g.lineTo(dx0, (dy0 + dy1) / 2); g.closePath();
  g.fillStyle = '#fff'; g.fill(); g.strokeStyle = '#000'; g.lineWidth = 1.33; g.stroke();
  txt(R(1, 9, T + 10, T + 10), 'se paga la totalidad de la orden 4 días despues de entregar los productos en bodega', { pt: 12, align: 'center' });
  txt(R(1, 9, T + 11, T + 11), '整個訂單在產品交付倉庫後 4 天支付。', { pt: 16, align: 'center', color: '#FF0000' });
  return { cv, W, H };
}

/* ---------- Consecutivos y generación (una sola vez por cotización) ----------
   Firestore: configuracion/consecutivos = { booth: <próximo Booth a usar>, clientes: { JD: 2, ... } }  (el contador de Item No va por nombre completo del cliente)
   Al generar la orden se guarda en cada cotización: orden = { booth, fila, itemNo, marca, fecha, creadaPor, creadaEn }.
   Una cotización con orden ya no se puede volver a generar: solo se vuelve a descargar el mismo PDF (no cambia ningún consecutivo). */
const OC_BOOTH_INICIAL = 17906;   // el 17905 ya se usó en una orden real fuera de la app
const ocRefContador = () => db.collection(APP_CONFIG.COLECCIONES.configuracion).doc('consecutivos');
const cotTieneOrden = c => !!(c.orden && c.orden.booth);
/* Qué se puede hacer con la selección: pendientes -> Excel "A cotizar"; finalizadas con costos y sin orden -> orden de compra */
const cotTipoSel = c => {
  const e = cotEstadoDe(c);
  if (e === 'pendiente') return 'excel';
  if (e === 'finalizada' && c.costos && !cotTieneOrden(c)) return 'orden';
  return null;
};

/* Nombre del cliente sin espacios ni tildes y con cada palabra en mayúscula inicial: "jorge duarte" -> "JorgeDuarte".
   Se usa para el Shipping Mark (JorgeDuarte-ADMA-GPK) y para el Item No (JorgeDuarte1, JorgeDuarte2...). */
function cotNombreClave(nombre) {
  const t = String(nombre || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z0-9\s]/g, ' ')
    .split(/\s+/).filter(Boolean).map(p => p[0].toUpperCase() + p.slice(1).toLowerCase()).join('');
  return t || 'Cliente';
}

function cotSeleccionar(id, marcar) {
  if (marcar && cotLista.some(c => c.id === id)) cotSeleccion.add(id); else cotSeleccion.delete(id);
  cotActualizarAcciones();
}
function cotSeleccionarVisibles() {
  cotSeleccion.clear();
  cotFiltradas().forEach(c => cotSeleccion.add(c.id));
  cotPintar();
}

/* Reserva el Booth y los Item No en una transacción (si dos personas generan a la vez, nunca se repiten) y marca las cotizaciones */
async function ocReservar(lista) {
  const d0 = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Bogota' }));
  const fecha = `${d0.getFullYear()}年  ${String(d0.getMonth() + 1).padStart(2, '0')}月  ${String(d0.getDate()).padStart(2, '0')}日`;
  const ordenes = {};
  await db.runTransaction(async tx => {
    const rc = ocRefContador();
    const lecturas = await Promise.all([tx.get(rc), ...lista.map(c => tx.get(cotRef().doc(c.id)))]);
    const sc = lecturas[0], docs = lecturas.slice(1);
    if (docs.some(s => !s.exists)) throw new Error('NO_EXISTE');
    if (docs.some(s => s.data().orden && s.data().orden.booth)) throw new Error('YA_TIENE_ORDEN');
    const d = sc.exists ? sc.data() : {};
    const booth = Math.max(Number.isInteger(d.booth) ? d.booth : 0, OC_BOOTH_INICIAL);
    const clientes = { ...(d.clientes || {}) };
    lista.forEach((c, i) => {
      const clave = cotNombreClave(c.cliente);
      clientes[clave] = (clientes[clave] || 0) + 1;
      ordenes[c.id] = { booth, fila: i, itemNo: `${clave}${clientes[clave]}`, marca: `${clave}-ADMA-GPK`, fecha, creadaPor: sesionActual ? sesionActual.nombre : '' };
    });
    tx.set(rc, { booth: booth + 1, clientes }, { merge: true });
    lista.forEach(c => tx.update(cotRef().doc(c.id), { orden: { ...ordenes[c.id], creadaEn: firebase.firestore.FieldValue.serverTimestamp() } }));
  });
  return ordenes;
}

/* Dibuja y descarga el PDF. lista: cotizaciones en el orden de las filas; ord: id -> { booth, itemNo, marca, fecha } */
async function ocImprimir(lista, ord) {
  const items = lista.map(c => {
    const k = c.costos, r = cotCalculos(c) || {};
    return {
      foto: (c.imagenesReales || [])[0] || (c.imagenes || [])[0] || '',
      descripcion: c.descripcion,
      marca: ord[c.id].marca, itemNo: ord[c.id].itemNo,
      ctn: r.cajas ?? k.cajas ?? 0, qtyCtn: k.unidadesPorCaja || 0,
      price: k.valorUnidad || 0, cbm: k.cbmCaja || 0,
      pesoCaja: cotOk(k.pesoCaja) ? k.pesoCaja : null,
      notas: c.observaciones || ''
    };
  });
  const o0 = ord[lista[0].id];
  const { cv, W, H } = await ocDibujar(items, o0.fecha, o0.booth);
  const doc = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'pt', format: 'letter' });
  const M = 24, k = Math.min((792 - M * 2) / W, (612 - M * 2) / H);
  doc.addImage(cv.toDataURL('image/jpeg', 0.93), 'JPEG', (792 - W * k) / 2, M, W * k, H * k);
  doc.save(`orden-de-compra-${o0.booth}.pdf`);
  return true;
}

/* Orden de compra: recibe el id de una cotización o un arreglo de ids (varias cotizaciones en una sola orden). */
async function cotGenerarOrden(ids) {
  if (!window.jspdf) { alert('No se cargó la librería de PDF. Revisa tu conexión y recarga la página.'); return false; }
  const lista = (Array.isArray(ids) ? ids : [ids]).map(id => cotLista.find(x => x.id === id)).filter(c => c && c.costos);
  if (!lista.length) { alert('Esta cotización aún no tiene costos guardados.'); return false; }
  const conOrden = lista.filter(cotTieneOrden);
  if (conOrden.length === lista.length) {            // ya generada: se vuelve a descargar el mismo PDF, sin tocar los consecutivos
    const booth = lista[0].orden.booth;
    const grupo = cotLista.filter(c => c.orden && c.orden.booth === booth).sort((a, b) => (a.orden.fila || 0) - (b.orden.fila || 0));
    const ord = {}; grupo.forEach(c => { ord[c.id] = c.orden; });
    return ocImprimir(grupo, ord);
  }
  if (conOrden.length) { alert('Algunas de las cotizaciones seleccionadas ya tienen orden de compra. Selecciona solo cotizaciones sin orden.'); return false; }
  const n = lista.length;
  if (!confirm(`Se generará UNA orden de compra con ${n} cotización${n > 1 ? 'es' : ''}.\n\nSe asignarán el Booth y los Item No consecutivos y NO se podrá generar otra vez (después solo podrás volver a descargar el mismo PDF).\n\n¿Continuar?`)) return false;
  let ord;
  try { ord = await ocReservar(lista); }
  catch (err) {
    console.error('No se pudo reservar el consecutivo:', err);
    alert(err.message === 'YA_TIENE_ORDEN' ? 'Alguien más acaba de generar la orden de compra de una de estas cotizaciones. Actualiza e inténtalo de nuevo.'
      : 'No se pudo reservar el consecutivo de la orden. Revisa tu conexión y las reglas de Firestore (configuracion/consecutivos).');
    return false;
  }
  cotSeleccion.clear();
  return ocImprimir(lista, ord);
}

/* ---------- Módulo Órdenes de compra ----------
   Lista las órdenes ya generadas (una por Booth/consecutivo) a partir de las cotizaciones que tienen `orden`.
   No guarda nada propio: todo sale de cotizaciones/{id}.orden, así que siempre está al día. */
const OC_MODULO = 'ordenes-compra';
let ocDetalleBooth = null;

const ocMonto = c => {                      // mismo cálculo que el PDF: precio por unidad × cajas × unidades por caja (RMB)
  const k = c.costos || {}, r = cotCalculos(c) || {};
  return (k.valorUnidad || 0) * (r.cajas ?? k.cajas ?? 0) * (k.unidadesPorCaja || 0);
};
const ocFechaTxt = ts => ts && ts.toDate
  ? ts.toDate().toLocaleDateString('es-CO', { timeZone: 'America/Bogota', day: '2-digit', month: 'short', year: 'numeric' }) : '';

function ocGrupos() {
  const mapa = new Map();
  cotLista.forEach(c => {
    if (!cotTieneOrden(c)) return;
    if (!mapa.has(c.orden.booth)) mapa.set(c.orden.booth, []);
    mapa.get(c.orden.booth).push(c);
  });
  return [...mapa.entries()].map(([booth, items]) => {
    items.sort((a, b) => (a.orden.fila || 0) - (b.orden.fila || 0));
    const o0 = items[0].orden;
    return { booth, items, total: items.reduce((t, c) => t + ocMonto(c), 0), fecha: ocFechaTxt(o0.creadaEn), por: o0.creadaPor || '' };
  }).sort((a, b) => b.booth - a.booth);
}

function ocConstruirModulo() {
  const sec = document.getElementById(OC_MODULO);
  if (!sec) return;
  sec.innerHTML = `
    <div class="cot-barra"><h2 class="cot-titulo">Órdenes de compra</h2>
      <div class="cot-filtros">
        <input type="search" data-oc-buscar placeholder="Buscar por Booth, cliente, Item No o producto…" aria-label="Buscar órdenes de compra" autocomplete="off">
      </div>
      <p class="cot-conteo" data-oc-conteo aria-live="polite"></p></div>
    <div class="cot-tabla-wrap" data-oc-grid aria-live="polite"><p class="lista-vacia">Cargando…</p></div>`;
  sec.addEventListener('input', e => { if (e.target.closest('[data-oc-buscar]')) ocPintar(); });
  sec.addEventListener('click', e => {
    const pdf = e.target.closest('[data-oc-pdf]');
    if (pdf) { cotGenerarOrden([pdf.dataset.ocPdf]); return; }
    if (e.target.closest('input, a, button, select')) return;
    const fila = e.target.closest('tr[data-oc-toggle]');
    if (fila) ocAbrirDetalle(fila.dataset.ocToggle);
  });
  sec.addEventListener('keydown', e => {
    if ((e.key === 'Enter' || e.key === ' ') && e.target.matches('tr[data-oc-toggle]')) { e.preventDefault(); ocAbrirDetalle(e.target.dataset.ocToggle); }
  });
  document.getElementById('ocDetalleCuerpo').addEventListener('click', e => {
    const pdf = e.target.closest('[data-oc-pdf]');
    const ver = e.target.closest('[data-oc-ver]');
    if (pdf) cotGenerarOrden([pdf.dataset.ocPdf]);
    else if (ver) { cotCerrarModal(document.getElementById('ocModalDetalle')); cotAlternar(ver.dataset.ocVer); }
  });
}

function ocPintar() {
  const sec = document.getElementById(OC_MODULO);
  const grid = sec && sec.querySelector('[data-oc-grid]');
  if (!grid) return;
  const conteo = sec.querySelector('[data-oc-conteo]');
  const todos = ocGrupos();
  if (!todos.length) {
    conteo.textContent = '';
    grid.innerHTML = '<p class="lista-vacia">Aún no hay órdenes de compra. Se generan desde Cotizaciones, con las cotizaciones finalizadas.</p>';
    return;
  }
  const q = cotNorm(sec.querySelector('[data-oc-buscar]').value).split(/\s+/).filter(Boolean);
  const lista = todos.filter(g => {
    const pajar = cotNorm([g.booth, g.por, g.fecha, ...g.items.flatMap(c => [c.cliente, c.descripcion, c.orden.itemNo, c.orden.marca])].join(' '));
    return q.every(t => pajar.includes(t));
  });
  conteo.textContent = lista.length === todos.length ? `${todos.length} órdenes` : `${lista.length} de ${todos.length} órdenes`;
  const pdfIco = `<svg class="cot-svg" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${COT_ICONOS.pdf}</svg>`;
  const filas = lista.map(g => {
    const clientes = [...new Set(g.items.map(c => c.cliente))].join(', ');
    return `<tr class="cot-fila" data-oc-toggle="${cotEsc(g.booth)}" tabindex="0" aria-haspopup="dialog" title="Ver detalle">
      <td><strong>Booth ${cotEsc(g.booth)}</strong></td>
      <td>${cotEsc(g.fecha)}</td>
      <td class="cot-col-prod"><span class="cot-prod-txt" title="${cotEsc(clientes)}">${cotEsc(clientes)}</span></td>
      <td class="cot-num">${g.items.length}</td>
      <td class="cot-num">¥ ${cotNum(g.total, 2)}</td>
      <td>${cotEsc(g.por)}</td>
      <td class="cot-col-acc"><div class="cot-acciones-fila cot-acciones-pg"><button type="button" class="cot-btn-tabla" data-oc-pdf="${g.items[0].id}" title="Descargar la orden de compra en PDF">${pdfIco}<span>Descargar PDF</span></button></div></td>
    </tr>`;
  }).join('');
  grid.innerHTML = lista.length ? `<div class="cot-tabla-card"><table class="cot-tabla">
    <thead><tr><th>Consecutivo</th><th>Fecha</th><th>Clientes</th><th class="cot-num">Cotizaciones</th><th class="cot-num">Total (RMB)</th><th>Creada por</th><th class="cot-col-acc">Acciones</th></tr></thead>
    <tbody>${filas}</tbody></table></div>` : '<p class="lista-vacia">Ninguna orden coincide con la búsqueda.</p>';
}

function ocAbrirDetalle(booth) {
  ocDetalleBooth = Number(booth);
  ocPintarDetalle();
  if (ocDetalleBooth != null) cotAbrir('ocModalDetalle');
}

function ocPintarDetalle() {
  const modal = document.getElementById('ocModalDetalle');
  const g = ocGrupos().find(x => x.booth === ocDetalleBooth);
  if (!g) { ocDetalleBooth = null; if (modal) modal.hidden = true; return; }
  const items = g.items.map(c => {
    const k = c.costos || {}, r = cotCalculos(c) || {};
    const ctn = r.cajas ?? k.cajas ?? 0, upc = k.unidadesPorCaja || 0;
    return `<div class="oc-item">
      <p><strong>${cotEsc(c.orden.itemNo)}</strong> · ${cotEsc(c.orden.marca)}</p>
      <p class="cot-dato"><strong>Cliente:</strong> ${cotEsc(c.cliente)}</p>
      <p>${cotEsc(c.descripcion)}</p>
      <p class="cot-dato"><strong>Cajas:</strong> ${cotNum(ctn, 0)} · <strong>Unidades por caja:</strong> ${cotNum(upc, 0)} · <strong>Cantidad:</strong> ${cotNum(ctn * upc, 0)}</p>
      <p class="cot-dato"><strong>Precio por unidad:</strong> ¥ ${cotNum(k.valorUnidad || 0, 2)} · <strong>Monto:</strong> ¥ ${cotNum(ocMonto(c), 2)}</p>
      ${c.observaciones ? `<p class="cot-dato cot-obs"><strong>Notes:</strong> ${cotEsc(c.observaciones)}</p>` : ''}
      <button type="button" class="cot-enlace" data-oc-ver="${c.id}" style="align-self:flex-start">Ver cotización</button>
    </div>`;
  }).join('');
  document.getElementById('ocDetalleTitulo').textContent = `Orden de compra · Booth ${g.booth}`;
  document.getElementById('ocDetalleCuerpo').innerHTML = `
    <p class="cot-dato"><strong>Fecha:</strong> ${cotEsc(g.fecha || '—')}${g.por ? ` · <strong>Creada por:</strong> ${cotEsc(g.por)}` : ''}</p>
    <p class="cot-dato"><strong>Cotizaciones:</strong> ${g.items.length} · <strong>Total:</strong> ¥ ${cotNum(g.total, 2)}</p>
    ${items}
    <footer><button type="button" class="btn btn-chico" data-oc-pdf="${g.items[0].id}">Descargar PDF</button></footer>`;
}

/* MODO PRUEBA (solo Super usuario): dibuja y descarga la orden de compra con Booth "PRUEBA" e Item No ficticios.
   No lee ni escribe Firestore: no sube el consecutivo, no marca la cotización y no limpia la selección. */
async function cotGenerarOrdenPrueba(ids) {
  if (!pgEsSuper()) { alert('El modo prueba es solo para el Super usuario.'); return false; }
  if (!window.jspdf) { alert('No se cargó la librería de PDF. Revisa tu conexión y recarga la página.'); return false; }
  const lista = (Array.isArray(ids) ? ids : [ids]).map(id => cotLista.find(x => x.id === id)).filter(c => c && c.costos);
  if (!lista.length) { alert('Esta cotización aún no tiene costos guardados.'); return false; }
  const d0 = new Date(new Date().toLocaleString('en-US', { timeZone: 'America/Bogota' }));
  const fecha = `${d0.getFullYear()}年  ${String(d0.getMonth() + 1).padStart(2, '0')}月  ${String(d0.getDate()).padStart(2, '0')}日`;
  const cont = {}, ord = {};
  lista.forEach((c, i) => {
    const clave = cotNombreClave(c.cliente);
    cont[clave] = (cont[clave] || 0) + 1;
    ord[c.id] = { booth: 'PRUEBA', fila: i, itemNo: `${clave}${cont[clave]}`, marca: `${clave}-ADMA-GPK`, fecha };
  });
  try { return await ocImprimir(lista, ord); }
  catch (err) { console.error('Error en la orden de prueba:', err); alert('No se pudo generar la orden de prueba.'); return false; }
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
  document.querySelectorAll('[data-cot-cerrar]').forEach(el => el.addEventListener('click', () => cotCerrarModal(el.closest('.modal'))));
  document.addEventListener('keydown', e => {
    if (e.key !== 'Escape') return;
    const abiertas = [...document.querySelectorAll('.modal')].filter(m => !m.hidden);
    if (abiertas.length) cotCerrarModal(abiertas[abiertas.length - 1]);
  });

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
    cotSeleccionar(cb.dataset.cotSel, cb.checked);
  });
  document.getElementById('cotBtnSelVisibles').addEventListener('click', () => cotSeleccionarVisibles());
  document.getElementById('cotBtnPdfSel').addEventListener('click',
    () => cotGenerarPDFCotizacion(cotLista.filter(c => cotSeleccion.has(c.id)).reverse()));
  document.getElementById('cotBtnOrdenSel').addEventListener('click', () => cotGenerarOrden([...cotSeleccion]));
  document.getElementById('cotBtnOrdenPrueba').addEventListener('click', () => cotGenerarOrdenPrueba([...cotSeleccion]));
  document.getElementById('cotBtnLimpiarSel').addEventListener('click', () => { cotSeleccion.clear(); cotPintar(); });
  document.getElementById('cotBtnExcelSel').addEventListener('click', () => cotGenerarExcel([...cotSeleccion]));

  document.getElementById('cotGrid').addEventListener('click', e => {
    const xls = e.target.closest('[data-cot-excel]');
    const pdf = e.target.closest('[data-cot-pdf]');
    const ord = e.target.closest('[data-cot-orden]');
    const mdl = e.target.closest('[data-cot-modulo]');
    const cont = e.target.closest('[data-cot-continuar]');
    const del = e.target.closest('[data-cot-eliminar]');
    if (xls) cotGenerarExcel([xls.dataset.cotExcel]);
    else if (pdf) cotGenerarPDF(pdf.dataset.cotPdf, pdf.dataset.tipo);
    else if (ord) cotGenerarOrden(ord.dataset.cotOrden);
    else if (mdl) cotMandarAModulo(mdl.dataset.cotModulo);
    else if (cont) cotAbrirCostos(cont.dataset.cotContinuar);
    else if (del) cotEliminar(del.dataset.cotEliminar);
    else if (!e.target.closest('input, a, button')) {
      const fila = e.target.closest('tr[data-cot-toggle]');
      if (fila) cotAlternar(fila.dataset.cotToggle);
    }
  });
  document.getElementById('cotDetalleCuerpo').addEventListener('click', e => {
    const ord = e.target.closest('[data-cot-orden]');
    const pru = e.target.closest('[data-cot-orden-prueba]');
    const obs = e.target.closest('[data-cot-obs]');
    if (obs) { cotCerrarModal(document.getElementById('cotModalDetalle')); cotAbrirObs(obs.dataset.cotObs, 'editar'); }
    else if (ord) cotGenerarOrden(ord.dataset.cotOrden);
    else if (pru) cotGenerarOrdenPrueba(pru.dataset.cotOrdenPrueba);
  });
  document.getElementById('cotGrid').addEventListener('keydown', e => {
    if ((e.key === 'Enter' || e.key === ' ') && e.target.matches('tr[data-cot-toggle]')) {
      e.preventDefault(); cotAlternar(e.target.dataset.cotToggle);
    }
  });

  /* Evita que la rueda del mouse o el panel táctil cambien, sin querer, el valor de un campo numérico:
     si el campo tiene el foco al desplazarse sobre él, se le quita el foco (la página sigue desplazándose normal). */
  document.addEventListener('wheel', e => {
    const a = document.activeElement;
    if (a && a.type === 'number' && a === e.target) a.blur();
  }, { capture: true, passive: true });

  /* Módulos Adma Company / Groupack y formulario de pagos */
  Object.values(COT_MODULO_DE_ORIGEN).forEach(pgConstruirModulo);
  ocConstruirModulo();
  pgInitDetalle();
  document.getElementById('cotFormPago').addEventListener('submit', pgGuardarPago);
  document.getElementById('cotFormObs').addEventListener('submit', cotGuardarObs);
  document.getElementById('cotPagoConcepto').addEventListener('change', () => { document.getElementById('cotPagoMonto').value = ''; });
  document.getElementById('cotPagoSaldo').addEventListener('click', () => {
    const c = cotLista.find(x => x.id === pgPagoId), R = c && pgResumen(c);
    if (R) document.getElementById('cotPagoMonto').value = Math.round(pgMaximo(document.getElementById('cotPagoConcepto').value, R));
  });
});
