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
        <button type="button" class="cot-enlace" id="cotBtnSelVisibles">Seleccionar pendientes visibles</button>
        <button type="button" class="cot-enlace" id="cotBtnSelFinal">Seleccionar finalizadas visibles</button>
        <span id="cotSelTexto" class="cot-conteo" style="flex-basis:auto"></span>
        <button type="button" class="btn btn-chico" id="cotBtnExcelSel" hidden></button>
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
      ['Total para el cliente con el aumento y la comision de Adma', r.totalRmb, 2, '', Y],
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
      c.costos ? `<button type="button" class="cot-btn-tabla" data-cot-pdf="${c.id}" data-tipo="cotizacion" title="Cotización - Generar PDF">${ico('pdf')}<span>Cotización PDF</span></button>` : '<span aria-hidden="true"></span>',
      puedeModulo ? `<button type="button" class="cot-btn-tabla cot-btn-modulo" data-cot-modulo="${c.id}" title="Mandar a módulo">${ico('enviar')}<span>Mandar a módulo</span></button>` : '<span aria-hidden="true"></span>',
      est !== 'finalizada' ? `<button type="button" class="cot-btn-tabla cot-btn-ico" data-cot-continuar="${c.id}" title="${etiquetaEditar}" aria-label="${etiquetaEditar} de ${cotEsc(c.cliente)}">${ico('editar')}</button>` : '<span aria-hidden="true"></span>',
      `<button type="button" class="cot-btn-tabla cot-btn-ico cot-btn-peligro" data-cot-eliminar="${c.id}" title="Eliminar" aria-label="Eliminar cotización de ${cotEsc(c.cliente)}">${ico('borrar')}</button>`
    ].join('');

    const fila = `<tr class="cot-fila" data-cot-toggle="${c.id}" tabindex="0" aria-haspopup="dialog" title="Ver detalle">
      <td class="cot-col-sel">${cotTipoSel(c) ? `<input type="checkbox" class="cot-sel" data-cot-sel="${c.id}"${cotSeleccion.has(c.id) ? ' checked' : ''} aria-label="Seleccionar la cotización de ${cotEsc(c.cliente)} para el Excel o la orden de compra">` : ''}</td>
      <td class="cot-col-cliente"><strong>${cotEsc(c.cliente)}</strong></td>
      <td class="cot-col-prod"><span class="cot-prod-txt">${cotEsc(c.descripcion)}</span></td>
      <td class="cot-col-estado"><span class="cot-estado cot-e-${est}">${COT_ESTADOS[est]}</span>${cotTieneOrden(c) ? `<span class="cot-oc" title="Orden de compra generada">OC ${c.orden.booth}</span>` : ''}</td>
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
    <p class="cot-dato"><strong>Cantidad:</strong> ${cotCantidadTxt(c)}</p>
    ${c.modulo ? `<p class="cot-dato"><strong>Enviada a:</strong> ${cotEsc(cotNombreModulo(c.modulo))}</p>` : ''}
    ${est === 'finalizada' || c.observaciones ? cotObsHtml(c, 'data-cot-obs') : ''}
    ${link ? `<p class="cot-dato"><a href="${cotEsc(link)}" target="_blank" rel="noopener noreferrer">Ver link del producto</a></p>` : ''}
    ${(c.imagenes || []).length ? `<div class="cot-miniaturas">${c.imagenes.map((s, i) => `<img src="${s}" alt="Referencia ${i + 1} de ${cotEsc(c.cliente)}">`).join('')}</div>` : ''}
    ${costos}
    ${cotTieneOrden(c) ? `<p class="cot-dato"><strong>Orden de compra:</strong> Booth ${c.orden.booth} · Item No ${cotEsc(c.orden.itemNo)} · Shipping Mark ${cotEsc(c.orden.marca)}</p>` : ''}
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

/* PDF de Cotización con el formato de ADMA (logo, datos, precio, flete, total, imágenes y pie). */
/* PDF de Cotización con el formato de ADMA: diseño de marca (amarillo/negro) pensado para la venta. */
async function cotGenerarPDFCotizacion(c) {
  const r = cotCalculos(c) || {};
  const doc = new window.jspdf.jsPDF({ unit: 'mm', format: 'a4' });
  const M = 20, ANCHO = 210 - M * 2, ALTO = 297, PIE = 24;
  const LIMITE = ALTO - PIE - 6;           // espacio reservado para el pie
  const AMARILLO = [255, 236, 0], OSCURO = [24, 24, 27], GRIS = [100, 100, 110], CLARO = [244, 244, 245], BORDE = [225, 225, 228];
  const mm = pt => pt * 0.3528;
  const cop = v => v == null ? '—' : '$ ' + cotNum(v, 2) + ' COP';
  let y = M;

  const nuevaPaginaSi = h => { if (y + h > LIMITE) { doc.addPage(); y = M; } };
  const seccion = t => {
    nuevaPaginaSi(14);
    doc.setFillColor(...AMARILLO); doc.rect(M, y, 2.2, 5.5, 'F');
    doc.setFont('helvetica', 'bold'); doc.setFontSize(11); doc.setTextColor(...OSCURO);
    doc.text(t.toUpperCase(), M + 5, y + 4.2);
    y += 9;
  };
  const fila = (a, b) => {
    nuevaPaginaSi(8);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(10.5); doc.setTextColor(...GRIS);
    doc.text(a, M + 1, y + 5);
    doc.setFont('helvetica', 'bold'); doc.setTextColor(...OSCURO);
    doc.text(String(b), M + ANCHO - 1, y + 5, { align: 'right' });
    y += 8;
    doc.setDrawColor(...BORDE); doc.setLineWidth(0.2); doc.line(M, y, M + ANCHO, y);
  };

  /* ---- Encabezado: logo grande a la izquierda, título y fecha a la derecha ---- */
  const logo = await cotLogoParaPDF();
  if (logo) {
    const k = Math.min(40 / logo.w, 30 / logo.h);
    doc.addImage(logo.data, 'PNG', M, y, logo.w * k, logo.h * k);
  }
  const fecha = new Date().toLocaleDateString('es-CO', { timeZone: 'America/Bogota', day: '2-digit', month: 'long', year: 'numeric' });
  doc.setTextColor(...OSCURO); doc.setFont('helvetica', 'bold'); doc.setFontSize(26);
  doc.text('COTIZACIÓN', 210 - M, y + 11, { align: 'right' });
  doc.setFont('helvetica', 'normal'); doc.setFontSize(12); doc.setTextColor(...GRIS);
  doc.text('Adma Importaciones', 210 - M, y + 18, { align: 'right' });
  doc.setFontSize(10);
  doc.text(fecha, 210 - M, y + 24, { align: 'right' });
  y += 36;
  doc.setFillColor(...AMARILLO); doc.rect(M, y, ANCHO, 1.8, 'F');
  y += 8;

  /* ---- Cliente y producto en una tarjeta ---- */
  doc.setFont('helvetica', 'normal'); doc.setFontSize(11);
  const desc = doc.splitTextToSize(String(c.descripcion || ''), ANCHO - 12);
  const hCliente = 30 + desc.length * 5;
  nuevaPaginaSi(hCliente);
  doc.setFillColor(...CLARO); doc.roundedRect(M, y, ANCHO, hCliente, 3, 3, 'F');
  doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.setTextColor(...GRIS);
  doc.text('CLIENTE', M + 6, y + 7);
  doc.setFontSize(14); doc.setTextColor(...OSCURO);
  doc.text(String(c.cliente || ''), M + 6, y + 13);
  doc.setFontSize(8); doc.setTextColor(...GRIS);
  doc.text('PRODUCTO', M + 6, y + 21);
  doc.setFont('helvetica', 'normal'); doc.setFontSize(11); doc.setTextColor(...OSCURO);
  desc.forEach((l, i) => doc.text(l, M + 6, y + 27 + i * 5));
  y += hCliente + 8;

  /* ---- Cantidades: dos tarjetas ---- */
  nuevaPaginaSi(26);
  const cw = (ANCHO - 6) / 2;
  [['CANTIDAD SOLICITADA', cotCantidadTxt(c)], ['CANTIDAD COTIZADA', r.unidades != null ? cotNum(r.unidades, 0) : '—']].forEach(([t, v], i) => {
    const x = M + i * (cw + 6);
    doc.setFillColor(...CLARO); doc.roundedRect(x, y, cw, 18, 3, 3, 'F');
    doc.setFillColor(...AMARILLO); doc.rect(x, y + 3, 1.6, 12, 'F');
    doc.setFont('helvetica', 'bold'); doc.setFontSize(8); doc.setTextColor(...GRIS);
    doc.text(t, x + 6, y + 6.5);
    doc.setFontSize(15); doc.setTextColor(...OSCURO);
    doc.text(String(v), x + 6, y + 14);
  });
  y += 26;

  /* ---- Precio del producto (COP, ya incluye aumento y comisión) ---- */
    seccion('Datos del precio del producto');
  fila('Costo Total', cop(r.totalCop));
  const kc = c.costos || {};
  const comisionCop = r.comisionTotal != null && kc.usdRmb > 0 && kc.usdCop != null
    ? (r.comisionTotal / kc.usdRmb) * kc.usdCop : null;
  fila('Comisión de Adma' + (r.comisionPct != null ? ` del ${cotNum(r.comisionPct, 2)}%` : ''), cop(comisionCop));
  y += 8;

  /* ---- Flete ---- */
  seccion('Datos del flete');
  fila('Tipo de flete', r.viaAerea ? 'Aéreo (sujeto a impuestos)' : 'Marítimo');
  fila('CBM', r.cbmTotal != null ? cotNum(r.cbmTotal, 4) + ' m³' : '—');
  fila('Cantidad de cajas', r.cajas != null ? cotNum(r.cajas, 0) : '—');
  const fleteTotal = r.logistica360 != null
    ? cop(r.logistica360) + (r.viaAerea && r.logistica360Usd != null ? ` (USD ${cotNum(r.logistica360Usd, 2)})` : '')
    : (r.logistica360Usd != null ? `USD ${cotNum(r.logistica360Usd, 2)}` : '—');
  fila('Costo Total', fleteTotal);
  y += 10;

  /* ---- Total destacado ---- */
  /* ---- Total destacado ---- */
  nuevaPaginaSi(46);
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

  /* ---- Imágenes (las reales cotizadas; si no hay, las de referencia) ---- */
  const imgs = (c.imagenesReales || []).length ? c.imagenesReales : (c.imagenes || []);
  if (imgs.length) {
    doc.addPage(); y = M;
    seccion('Imágenes del producto');
    const colW = (ANCHO - 6) / 2, maxH = 70;
    let col = 0, altoFila = 0;
    for (const src of imgs) {
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
  }

  /* ---- Pie de página en todas las hojas ---- */
  const n = doc.getNumberOfPages();
  for (let p = 1; p <= n; p++) {
    doc.setPage(p);
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

  const base = c.cliente.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'cliente';
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
  const n = sel.length, tipo = n ? cotTipoSel(sel[0]) : null;
  document.getElementById('cotSelTexto').textContent = n ? `${n} seleccionada${n > 1 ? 's' : ''}` : '';
  const bx = document.getElementById('cotBtnExcelSel'), bo = document.getElementById('cotBtnOrdenSel');
  bx.hidden = tipo !== 'excel'; bx.textContent = `A cotizar - Generar Excel (${n})`;
  bo.hidden = tipo !== 'orden'; bo.textContent = `Generar orden de compra (${n})`;
  const bp = document.getElementById('cotBtnOrdenPrueba');   // solo Super usuario
  bp.hidden = !(tipo === 'orden' && pgEsSuper()); bp.textContent = `Orden de PRUEBA (${n})`;
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
    ${fotos.length ? `<div class="cot-miniaturas">${fotos.map((s, i) => `<img src="${s}" alt="Foto ${i + 1} de ${cotEsc(c.cliente)}">`).join('')}</div>` : ''}
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
    <td class="cot-col-cliente"><div class="pg-cliente">${foto ? `<img class="pg-thumb" src="${foto}" alt="">` : '<span class="pg-thumb" aria-hidden="true"></span>'}<strong>${cotEsc(c.cliente)}</strong></div></td>
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
    [...cotSeleccion].forEach(id => { if (!cotLista.some(c => c.id === id && cotTipoSel(c))) cotSeleccion.delete(id); });
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
  const c = cotLista.find(x => x.id === id);
  if (!c) return;
  if (!marcar) { cotSeleccion.delete(id); cotActualizarAcciones(); return; }
  const t = cotTipoSel(c);
  if (!t) return;
  let cambio = false;   // no se mezclan pendientes (Excel) con finalizadas (orden de compra)
  [...cotSeleccion].forEach(o => { const oc = cotLista.find(x => x.id === o); if (!oc || cotTipoSel(oc) !== t) { cotSeleccion.delete(o); cambio = true; } });
  cotSeleccion.add(id);
  cambio ? cotPintar() : cotActualizarAcciones();
}
function cotSeleccionarVisibles(tipo) {
  cotSeleccion.clear();
  cotFiltradas().filter(c => cotTipoSel(c) === tipo).forEach(c => cotSeleccion.add(c.id));
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
    return `<tr class="cot-fila" data-oc-toggle="${g.booth}" tabindex="0" aria-haspopup="dialog" title="Ver detalle">
      <td><strong>Booth ${g.booth}</strong></td>
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
    cotSeleccionar(cb.dataset.cotSel, cb.checked);
  });
  document.getElementById('cotBtnSelVisibles').addEventListener('click', () => cotSeleccionarVisibles('excel'));
  document.getElementById('cotBtnSelFinal').addEventListener('click', () => cotSeleccionarVisibles('orden'));
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
