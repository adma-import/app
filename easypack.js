/* ===== Módulo EasyPack =====
   Réplica del módulo Cotizaciones, pero solo con cotizaciones de EasyPack (colección `cotizaciones`, origen === 'EasyPack';
   cotizaciones.js las separa y las entrega a epRecibir()).
   Documento:
   {
     origen: 'EasyPack', consecutivo (número, 1,2,3…; lo asigna una transacción sobre configuracion/consecutivos.easypack),
     cliente, empresa (opcional), nit (opcional), nombreProducto (español), nombreProductoEn (inglés),
     descripcion, cantidad, cantidadMin, link, imagenes[],
     estado: 'pendiente' | 'cotizada', anulada: true (conserva el consecutivo; se puede crear una copia con el siguiente número),
     costosEp: { ...entradas del formulario de edición + resultados calculados (ver epCalcular) },
     creadoPor, creadoEn
   } */

const EP_ORIGEN = 'EasyPack';
const EP_VALIDEZ_DIAS = 15;
const EP_ENTREGA_DIAS = { aerea: 25, maritima: 90 };
const EP_MODALIDAD = { aerea: 'Aérea', maritima: 'Marítima' };
const EP_FIRMA = { nombre: 'Danna Galeano Grandas', cargo: 'Directora Comercial EasyPack Solutions', correo: 'directorcomercial@easypacklatam.com' };
const EP_OPCIONES_PAGINA = [10, 50, 100];
const EP_MESES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];

/* Constantes de las fórmulas del Excel "datos formulario edicion" (se dejan juntas para ajustarlas fácil) */
const EP_K = {
  DIV_PRECIO: 0.14,          // L = USD / 0.14
  RECARGO_1: 1.05,           // N = L × 1.05
  RECARGO_2: 1.1,            // O = N × 1.1
  FACTOR_COSTO: 0.14,        // U = N × 0.14  (costo del producto en USD)
  AEREO_USD_KG: { NORMAL: 18, ESPECIAL: 23 },          // tarifa aérea en USD por kg de caja (columna X del Excel: NORMAL 18, ESPECIAL 23)
  BASE_IMPUESTOS: 0.3,       // IVA y arancel se calculan sobre el 30 % del valor en COP
  IVA: 0.19,
  ARANCEL: 0.1,
  GASTO_DEFECTO: { NORMAL: 2, ESPECIAL: 3 }           // USD, "gasto de emergencia" por tipo de tarifa
};

/* Logo del módulo: un frasco de perfume y un termo (envases) */
const EP_LOGO = '<svg class="ep-logo" viewBox="0 0 48 32" width="44" height="30" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="14" width="16" height="15" rx="3"/><rect x="8" y="9" width="6" height="5"/><rect x="7" y="4" width="8" height="5" rx="1.5"/><path d="M7 21h8"/><path d="M32 6h8v4h-8z"/><path d="M31 10h10v16a3 3 0 0 1-3 3h-4a3 3 0 0 1-3-3z"/><path d="M31 17h10"/></svg>';

let epLista = [];
let epImagenes = [];
let epEditId = null;                   // cotización cuyos datos se editan (null = creando)
let epCostosId = null;                 // cotización a la que se le llenan costos
let epPagina = 1, epPorPagina = 10;
const epSel = new Set();

/* ---------- Utilidades ---------- */
/* Sin redondeo: se muestran todos los decimales del valor (hasta 10, solo para quitar el ruido de coma flotante) */
const epNum = (v, dec = 10) => Number(v).toLocaleString('es-CO', { minimumFractionDigits: 0, maximumFractionDigits: dec });
const epCop = v => '$ ' + epNum(v, 2);   // valores en pesos: máximo 2 decimales, sin rellenar con ceros
const epEstadoDe = c => c.anulada ? 'anulada' : (c.modulo === 'easypack' && c.estado === 'finalizada' ? 'finalizada' : (c.costosEp ? 'cotizada' : 'pendiente'));
const EP_ESTADOS = { pendiente: 'Pendiente', cotizada: 'Cotizada', finalizada: 'Finalizada', anulada: 'Anulada' };
const epIco = n => `<svg class="cot-svg" viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${COT_ICONOS[n]}</svg>`;
const epFechaHoy = () => {                               // dd/mmm/aaaa en hora de Bogotá
  const [a, m, d] = (typeof hoyBogota === 'function' ? hoyBogota() : new Date().toISOString().slice(0, 10)).split('-');
  return `${d}/${EP_MESES[Number(m) - 1]}/${a}`;
};

/* ---------- Fórmulas del Excel ----------
   Entradas (k): qtyCaja, alto, ancho, largo (cm), pesoCaja (kg, se escribe a mano como en el Excel), precioUsd, tipo ('NORMAL'|'ESPECIAL'), gasto (USD), fleteBarco (COP por m³).
   trm = dólar a pesos. Devuelve null en lo que no se puede calcular (p. ej. el flete aéreo sin peso: "SIN PESO" en el Excel). */
function epCalcular(k, trm) {
  const qty = k.qtyCaja > 0 ? k.qtyCaja : 0;
  const cbmCaja = (k.alto * k.ancho * k.largo) / 1e6;                       // H
  const pesoCaja = k.pesoCaja > 0 ? k.pesoCaja : 0;                          // J (kg, escrito a mano)
  const L = k.precioUsd / EP_K.DIV_PRECIO;
  const N = L * EP_K.RECARGO_1;
  const O = N * EP_K.RECARGO_2;
  const precioChinaUsd = O + (k.gasto || 0);                                 // Q (informativo: no entra al precio)
  const costoProdUsd = N * EP_K.FACTOR_COSTO;                                // U
  const costoProdCop = costoProdUsd * trm;                                   // V
  const fleteMarUnidad = qty ? (k.fleteBarco * cbmCaja) / qty : 0;           // T
  const costoImportBarco = costoProdCop + fleteMarUnidad;                    // W (producto + flete, antes de impuestos)
  const fleteAereoCajaUsd = (EP_K.AEREO_USD_KG[k.tipo] || EP_K.AEREO_USD_KG.NORMAL) * pesoCaja;   // X
  const fleteAereoCajaCop = fleteAereoCajaUsd * trm;                         // Y
  const fleteAereoUnidad = qty && fleteAereoCajaCop > 0 ? fleteAereoCajaCop / qty : null;   // Z
  const precioAereo = fleteAereoUnidad == null ? null : fleteAereoUnidad + costoProdCop;    // AB
  const iva = k.precioUsd * trm * EP_K.BASE_IMPUESTOS * EP_K.IVA;            // AC
  const arancel = k.precioUsd * trm * EP_K.BASE_IMPUESTOS * EP_K.ARANCEL;    // AE
  const precioMaritimo = costoImportBarco - iva + arancel;                   // AF: COSTO APROX NACIONALIZADO BARCO (es el precio marítimo)
  return { cbmCaja, pesoCaja, precioChinaUsd, costoProdUsd, costoProdCop, fleteMarUnidad, precioMaritimo,
           fleteAereoUnidad, precioAereo, iva, arancel, costoImportBarco, nacionalizado: precioMaritimo };
}
const epPrecioUnidad = (r, modalidad) => modalidad === 'aerea' ? r.precioAereo : r.precioMaritimo;

/* ---------- Interfaz ---------- */
function epConstruirUI() {
  const sec = document.getElementById('easypack-cot');   // pestaña Cotizaciones del módulo (la arma cotizaciones.js)
  if (!sec) return false;
  sec.innerHTML = `
    <div class="cot-barra">
      <div class="cot-cabecera"><h2 class="cot-titulo">${EP_LOGO}EasyPack · Cotizaciones</h2>
        <button type="button" class="btn cot-btn-crear" id="epBtnNueva" title="Crear una cotización de EasyPack"><svg class="cot-svg" viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${COT_ICONOS.crear}</svg><span>Crear cotización</span></button></div>
      <div class="cot-filtros">
        <input type="search" id="epBuscar" placeholder="Buscar por cliente, empresa, NIT, producto…" aria-label="Buscar cotizaciones de EasyPack" autocomplete="off">
        <label class="cot-filtro"><span>Filtrar por estado</span><select id="epFiltroEstado" aria-label="Filtrar por estado"><option value="">Todos</option><option value="pendiente">Pendiente</option><option value="cotizada">Cotizada</option><option value="finalizada">Finalizada</option><option value="anulada">Anulada</option></select></label>
      </div>
      <p id="epConteo" class="cot-conteo" aria-live="polite"></p>
      <div class="cot-acciones">
        <button type="button" class="btn btn-chico" id="epBtnPdfSel" hidden></button>
        <button type="button" class="cot-enlace" id="epBtnLimpiarSel" hidden>Quitar selección</button>
        <button type="button" class="btn btn-chico secundario" id="epBtnRecalcular" hidden title="Recalcula los costos de todas las cotizaciones con las fórmulas actuales (solo Super usuario)">Recalcular costos</button>
      </div></div>
    <div id="epGrid" class="cot-tabla-wrap" aria-live="polite"><p class="lista-vacia">Cargando cotizaciones…</p></div>`;

  const m = document.createElement('div');
  m.innerHTML = `
  <div id="epModalNueva" class="modal" hidden>
    <div class="modal-fondo"></div>
    <form id="epFormNueva" class="modal-tarjeta cot-modal" novalidate>
      <div class="modal-cabecera"><h2 id="epNuevaTitulo">Nueva cotización EasyPack</h2>
        <button type="button" class="modal-cerrar" data-ep-cerrar aria-label="Cerrar">✕</button></div>
      <div id="epErrorNueva" class="login-error" role="alert" hidden></div>
      <label class="login-campo"><span>Nombre del cliente</span><input type="text" id="epCliente" required autocomplete="off"></label>
      <label class="login-campo"><span>Nombre de la empresa (opcional)</span><input type="text" id="epEmpresa" autocomplete="off"></label>
      <label class="login-campo"><span>NIT (opcional)</span><input type="text" id="epNit" autocomplete="off"></label>
      <label class="login-campo"><span>Nombre del producto (español)</span><input type="text" id="epProductoEs" required autocomplete="off"></label>
      <label class="login-campo"><span>Nombre del producto (inglés)</span><input type="text" id="epProductoEn" autocomplete="off"></label>
      <label class="login-campo"><span>Descripción del producto</span><textarea id="epDescripcion" rows="3" required></textarea></label>
      <label class="login-campo"><span>Cantidad</span><input type="text" id="epCantidad" inputmode="numeric" autocomplete="off" placeholder="Ej: 100 o un rango 50-100" required></label>
      <label class="login-campo"><span>Link</span><input type="url" id="epLink" placeholder="https://" autocapitalize="none"></label>
      <div class="login-campo"><span>Imágenes de referencia (máx. ${COT_MAX_IMAGENES})</span>
        <input type="file" id="epArchivos" accept="image/*" multiple>
        <div id="epPrevias" class="cot-previas"></div></div>
      <div class="modal-acciones">
        <button type="button" class="btn secundario" data-ep-cerrar>Cancelar</button>
        <button type="submit" class="btn" id="epBtnGuardar">Crear cotización</button>
      </div>
    </form>
  </div>

  <div id="epModalCostos" class="modal" hidden>
    <div class="modal-fondo"></div>
    <form id="epFormCostos" class="modal-tarjeta cot-modal" novalidate>
      <div class="modal-cabecera"><h2 id="epCostosTitulo">Costos</h2>
        <button type="button" class="modal-cerrar" data-ep-cerrar aria-label="Cerrar">✕</button></div>
      <div id="epErrorCostos" class="login-error" role="alert" hidden></div>
      <label class="login-campo"><span>Unidades por caja</span><input type="number" id="epQtyCaja" min="1" step="1" required></label>
      <div class="ep-trio">
        <label class="login-campo"><span>Alto (cm)</span><input type="number" id="epAlto" min="0" step="any" required></label>
        <label class="login-campo"><span>Ancho (cm)</span><input type="number" id="epAncho" min="0" step="any" required></label>
        <label class="login-campo"><span>Largo (cm)</span><input type="number" id="epLargo" min="0" step="any" required></label>
      </div>
      <label class="login-campo"><span>Peso por caja (kg)</span><input type="number" id="epPesoUnit" min="0" step="any" required></label>
      <label class="login-campo"><span>Precio del producto (USD)</span><input type="number" id="epPrecioUsd" min="0" step="any" required></label>
      <label class="login-campo"><span>Tipo de tarifa</span><select id="epTipo"><option value="NORMAL">Normal</option><option value="ESPECIAL">Especial</option></select></label>
      <label class="login-campo"><span>Gasto de emergencia (USD, informativo)</span><input type="number" id="epGasto" min="0" step="any"></label>
      <label class="login-campo"><span>Precio del flete marítimo (COP por m³)</span><select id="epFlete"></select></label>
      <label class="login-campo"><span>Modalidad de la cotización</span><select id="epModalidad"><option value="aerea">Aérea (${EP_ENTREGA_DIAS.aerea} días)</option><option value="maritima">Marítima (${EP_ENTREGA_DIAS.maritima} días)</option></select></label>
      <label class="login-campo cot-check"><span><input type="checkbox" id="epStock"> Hay stock disponible</span></label>
      <div class="cot-dato cot-calc ep-resumen" id="epResumen" aria-live="polite"></div>
      <div class="modal-acciones">
        <button type="button" class="btn secundario" data-ep-cerrar>Cancelar</button>
        <button type="submit" class="btn" id="epBtnGuardarCostos">Guardar</button>
      </div>
    </form>
  </div>`;
  document.getElementById('appShell').appendChild(m);
  if (!document.getElementById('pan-ep')) {
    const s = document.createElement('section');
    s.id = 'pan-ep'; s.hidden = true; s.className = 'pan';
    document.querySelector('.modulos').appendChild(s);
  }
  return true;
}

const epCerrar = el => { if (el) el.hidden = true; if (el && el.id === 'epModalNueva') epEditId = null; if (el && el.id === 'epModalCostos') epCostosId = null; };
function epCerrarTodo() { document.querySelectorAll('#epModalNueva, #epModalCostos').forEach(epCerrar); }
const epError = (id, msg) => { const e = document.getElementById(id); e.textContent = msg || ''; e.hidden = !msg; };

/* ---------- Datos de la lista ---------- */
function epRecibir(lista) {                                     // lo llama cotizaciones.js desde la escucha en tiempo real
  epLista = lista.slice().sort((a, b) => (b.consecutivo || 0) - (a.consecutivo || 0));
  [...epSel].forEach(id => { if (!epLista.some(c => c.id === id && !c.anulada)) epSel.delete(id); });
  if (document.getElementById('epGrid')) epPintar();
  if (epDetalleId) epPintarDetalle();
}

function epFiltradas() {
  const q = cotNorm(document.getElementById('epBuscar').value).split(/\s+/).filter(Boolean);
  const est = document.getElementById('epFiltroEstado').value;
  return epLista.filter(c => {
    const e = epEstadoDe(c);
    if (est && e !== est) return false;
    if (!q.length) return true;
    const pajar = cotNorm([c.consecutivo, c.cliente, c.empresa, c.nit, c.nombreProducto, c.nombreProductoEn, c.descripcion, EP_ESTADOS[e], c.creadoPor].join(' '));
    return q.every(t => pajar.includes(t));
  });
}

function epControlesPagina(total, ini, mostradas, paginas) {
  return `<div class="cot-paginacion">
    <label class="cot-pag-tam"><span>Filas por página</span><select data-ep-por-pagina aria-label="Filas por página">${EP_OPCIONES_PAGINA.map(n => `<option value="${n}"${n === epPorPagina ? ' selected' : ''}>${n}</option>`).join('')}</select></label>
    <span class="cot-pag-info">${ini + 1}–${ini + mostradas} de ${total}</span>
    <div class="cot-pag-nav">
      <button type="button" class="btn btn-chico secundario" data-ep-pag="prev"${epPagina <= 1 ? ' disabled' : ''}>Anterior</button>
      <span class="cot-pag-num">Página ${epPagina} de ${paginas}</span>
      <button type="button" class="btn btn-chico secundario" data-ep-pag="next"${epPagina >= paginas ? ' disabled' : ''}>Siguiente</button>
    </div></div>`;
}

function epPintar() {
  epActualizarAcciones();
  const grid = document.getElementById('epGrid');
  const conteo = document.getElementById('epConteo');
  if (!epLista.length) {
    conteo.textContent = '';
    grid.innerHTML = '<p class="lista-vacia">Aún no hay cotizaciones de EasyPack. Toca el botón + para crear la primera.</p>';
    return;
  }
  const lista = epFiltradas();
  conteo.textContent = lista.length === epLista.length ? `${epLista.length} cotizaciones` : `${lista.length} de ${epLista.length} cotizaciones`;
  if (!lista.length) { grid.innerHTML = '<p class="lista-vacia">Ninguna cotización coincide con la búsqueda.</p>'; return; }

  const paginas = Math.max(1, Math.ceil(lista.length / epPorPagina));
  epPagina = Math.min(Math.max(1, epPagina), paginas);
  const ini = (epPagina - 1) * epPorPagina;
  const pagina = lista.slice(ini, ini + epPorPagina);

  const filas = pagina.map(c => {
    const est = epEstadoDe(c);
    const k = c.costosEp;
    const precio = k && cotOk(k.precioUnidad) ? `$ ${epNum(k.precioUnidad, 2)} <small>(${EP_MODALIDAD[k.modalidad] || ''})</small>` : '—';
    const ver = `<button type="button" class="cot-btn-tabla" data-ep-ver="${c.id}" title="Ver el detalle de la cotización" aria-label="Ver la cotización ${c.consecutivo}">${epIco('ver')}<span>Ver</span></button>`;
    // En la tabla solo Ver y Anular; costos, PDF, compra finalizada, editar y copia están en el detalle
    const acc = ver + (est === 'anulada' || est === 'finalizada' ? ''
      : `<button type="button" class="cot-btn-tabla cot-btn-ico cot-btn-peligro" data-ep-anular="${c.id}" title="Anular (conserva el consecutivo)" aria-label="Anular la cotización ${c.consecutivo}">${epIco('borrar')}</button>`);
    return `<tr class="cot-fila${est === 'anulada' ? ' ep-anulada' : ''}" data-ep-ver="${c.id}" tabindex="0" title="Ver detalle">
      <td class="cot-col-sel"><input type="checkbox" data-ep-sel="${c.id}"${epSel.has(c.id) ? ' checked' : ''}${k && est !== 'anulada' ? '' : ' disabled'} aria-label="Seleccionar la cotización ${c.consecutivo}"></td>
      <td class="cot-col-cliente"><strong>${cotEsc(c.cliente)}</strong></td>
      <td class="cot-col-prod"><span class="cot-prod-txt">${cotEsc(c.nombreProducto)}${c.nombreProductoEn ? `<br><small>${cotEsc(c.nombreProductoEn)}</small>` : ''}</span></td>
      <td class="cot-num">${cotCantidadTxt(c)}</td>
      <td>${cotEsc(ocFechaTxt(c.creadoEn) || '—')}</td>
      <td class="cot-num">${precio}</td>
      <td class="cot-col-estado"><span class="cot-estado cot-e-${est}">${EP_ESTADOS[est]}</span></td>
      <td class="cot-col-acc"><div class="cot-acciones-fila">${acc}</div></td>
    </tr>`;
  }).join('');

  grid.innerHTML = `<div class="cot-tabla-card"><table class="cot-tabla">
    <thead><tr><th class="cot-col-sel"><span class="sr-only">Seleccionar</span></th><th>Cliente</th><th>Producto</th><th class="cot-num">Cantidad</th><th>Fecha</th><th class="cot-num">Precio unidad</th><th>Estado</th><th class="cot-col-acc">Acciones</th></tr></thead>
    <tbody>${filas}</tbody></table></div>${epControlesPagina(lista.length, ini, pagina.length, paginas)}`;
}

/* Recalcula costosEp de todas las cotizaciones con las fórmulas actuales (con las mismas entradas y el dólar que ya tenían guardados). Solo Super usuario. */
async function epRecalcularTodas() {
  if (typeof pgEsSuper !== 'function' || !pgEsSuper()) { alert('Solo el Super usuario puede recalcular los costos.'); return; }
  const trmActual = typeof TASAS !== 'undefined' ? TASAS.usdCop : null;
  const plan = [], omitidas = [];
  epLista.filter(c => c.costosEp).forEach(c => {
    const k = c.costosEp;
    const pesoCaja = k.pesoCaja > 0 ? k.pesoCaja : (k.pesoUnit > 0 && k.qtyCaja > 0 ? (k.pesoUnit * k.qtyCaja) / 1000 : 0);
    const trm = k.usdCop > 0 ? k.usdCop : trmActual;
    const entrada = { qtyCaja: k.qtyCaja, alto: k.alto, ancho: k.ancho, largo: k.largo, pesoCaja, precioUsd: k.precioUsd, tipo: k.tipo, gasto: k.gasto || 0, fleteBarco: k.fleteBarco };
    if (!(trm > 0) || ![entrada.qtyCaja, entrada.alto, entrada.ancho, entrada.largo, entrada.precioUsd, entrada.fleteBarco].every(x => x >= 0) || !(entrada.qtyCaja > 0)) { omitidas.push(c); return; }
    const r = epCalcular(entrada, trm);
    const precioUnidad = epPrecioUnidad(r, k.modalidad);
    if (precioUnidad == null) { omitidas.push(c); return; }
    const nuevo = { ...k, ...r, pesoCaja, precioUnidad, usdCop: trm, recalculadoEn: new Date().toISOString().slice(0, 10) };
    delete nuevo.pesoUnit;
    Object.keys(nuevo).forEach(key => { if (typeof nuevo[key] === 'number' && !Number.isFinite(nuevo[key])) nuevo[key] = null; });
    plan.push({ c, nuevo, cambia: Math.abs((k.precioUnidad || 0) - precioUnidad) > 1e-9 });
  });
  if (!plan.length) { alert('No hay cotizaciones de EasyPack con costos para recalcular.'); return; }
  const cambian = plan.filter(p => p.cambia).length;
  if (!confirm(`Se recalcularán los costos de ${plan.length} cotización(es) de EasyPack con las fórmulas actuales (el precio por unidad cambiará en ${cambian}).` +
    (omitidas.length ? `\n\nSe omiten ${omitidas.length} por datos incompletos (por ejemplo, aéreas sin peso).` : '') +
    '\n\nLos PDF que ya enviaste no cambian, pero los nuevos saldrán con los valores recalculados.\n\n¿Continuar?')) return;
  const btn = document.getElementById('epBtnRecalcular'); btn.disabled = true;
  try {
    for (let i = 0; i < plan.length; i += 400) {
      const batch = db.batch();
      plan.slice(i, i + 400).forEach(p => batch.update(cotRef().doc(p.c.id), { costosEp: p.nuevo }));
      await batch.commit();
    }
    alert(`Listo: se recalcularon ${plan.length} cotizaciones.` + (omitidas.length ? ` Omitidas: ${omitidas.map(c => c.consecutivo + ' (' + c.cliente + ')').join(', ')}.` : ''));
  } catch (err) { console.error(err); alert('No se pudo recalcular. Revisa tu conexión y las reglas de Firestore.'); }
  finally { btn.disabled = false; }
}

/* ---------- Pantalla de detalle (igual que la de Groupack, con los datos propios de EasyPack) ---------- */
let epDetalleId = null;
function epAbrirDetalle(id) {
  if (!epLista.some(c => c.id === id)) return;
  epDetalleId = id;
  epPintarDetalle();
  if (epDetalleId) panAbrir('pan-ep');
}

function epPintarDetalle() {
  const sec = document.getElementById('pan-ep');
  const c = epLista.find(x => x.id === epDetalleId);
  if (!c) { epDetalleId = null; if (sec && !sec.hidden && typeof panVolver === 'function') panVolver(); return; }
  const est = epEstadoDe(c), k = c.costosEp, link = cotLinkSeguro(c.link);
  const btn = (ico, txt, attr, cls = '') => `<button type="button" class="cot-btn-tabla ${cls}" ${attr} title="${cotEsc(txt)}">${epIco(ico)}<span>${cotEsc(txt)}</span></button>`;
  const acciones = est === 'anulada'
    ? btn('enviar', 'Crear copia', `data-ep-copia="${c.id}"`)
    : est === 'finalizada'
    ? btn('pdf', 'Cotización PDF', `data-ep-pdf="${c.id}"`)
    : [
      btn('excel', k ? 'Editar costos' : 'Llenar costos', `data-ep-costos="${c.id}"`),
      k ? btn('pdf', 'Cotización PDF', `data-ep-pdf="${c.id}"`) : '',
      k ? btn('enviar', 'Compra finalizada', `data-ep-finalizar="${c.id}"`, 'cot-btn-modulo') : '',
      btn('editar', 'Editar datos', `data-ep-editar="${c.id}"`),
      btn('borrar', 'Anular', `data-ep-anular="${c.id}"`, 'cot-btn-peligro')
    ].join('');

  const dato = (t, v) => v ? `<div class="pan-dato"><dt>${t}</dt><dd>${v}</dd></div>` : '';
  const datos = [
    dato('No. de cotización', cotEsc(c.consecutivo)),
    dato('Producto', cotEsc(c.nombreProducto || c.descripcion)),
    dato('Producto (inglés)', cotEsc(c.nombreProductoEn || '')),
    c.nombreProducto ? dato('Descripción', cotEsc(c.descripcion || '')) : '',
    dato('Cliente', cotEsc(c.cliente)),
    dato('Empresa', cotEsc(c.empresa || '')),
    dato('NIT', cotEsc(c.nit || '')),
    dato('Cliente de', cotEsc(c.origen || EP_ORIGEN)),
    dato('Cantidad', cotCantidadTxt(c)),
    dato('Fecha de creación', cotEsc(ocFechaTxt(c.creadoEn) || '—')),
    dato('Creada por', cotEsc(c.creadoPor || '')),
    dato('Modalidad', k ? cotEsc(EP_MODALIDAD[k.modalidad] || '') : ''),
    dato('Stock', k ? (k.stock ? 'Hay stock disponible' : 'Sin stock') : ''),
    c.copiaDe ? dato('Copia de', 'No. ' + cotEsc(c.copiaDe)) : '',
    c.anulada && c.anuladaPor ? dato('Anulada por', cotEsc(c.anuladaPor)) : '',
    link ? dato('Link', `<a href="${cotEsc(link)}" target="_blank" rel="noopener noreferrer">Ver link del producto</a>`) : ''
  ].join('');

  const imgs = (c.imagenes || []).length
    ? `<div class="pan-card"><h3 class="pan-sub">Imágenes de referencia</h3><div class="cot-miniaturas pan-imgs">${c.imagenes.map((src, i) => `<img src="${cotEsc(src)}" alt="Referencia ${i + 1} de ${cotEsc(c.cliente)}">`).join('')}</div></div>` : '';

  const ok = v => v != null && Number.isFinite(Number(v));
  const l = (a, v, fmt) => ok(v) ? [a, fmt(v)] : null;
  const n0 = v => epNum(v, 0), n2 = v => epNum(v), n3 = v => epNum(v), n4 = v => epNum(v);
  const grupos = k ? [
    ['Caja y producto', [l('Unidades por caja', k.qtyCaja, n0), l('Alto (cm)', k.alto, n2), l('Ancho (cm)', k.ancho, n2), l('Largo (cm)', k.largo, n2),
      l('CBM por caja', k.cbmCaja, n4), l('Peso por caja (kg)', k.pesoCaja, n3), ['Tipo de tarifa', k.tipo === 'ESPECIAL' ? 'Especial' : 'Normal']]],
    ['Precios', [l('Precio del producto (USD)', k.precioUsd, n2), l('Gasto de emergencia (USD)', k.gasto, n2), l('Dólar a pesos (TRM)', k.usdCop, n2),
      l('Flete marítimo (COP por m³)', k.fleteBarco, epCop), l('Costo del producto (COP)', k.costoProdCop, epCop), l('IVA por unidad', k.iva, epCop), l('Arancel por unidad', k.arancel, epCop),
      l('Costo de importación barco por unidad (producto + flete)', k.costoImportBarco, epCop), l('Costo de la cotización por unidad', k.precioUnidad, epCop),
      l('Costo marítimo por unidad', k.precioMaritimo, epCop), l('Costo aéreo por unidad', k.precioAereo, epCop)]]
  ].map(([t, ls]) => [t, ls.filter(Boolean)]) : [];
  const tablas = grupos.map(([t, ls]) => ls.length ? `
    <div class="pan-card"><h3 class="pan-sub">${cotEsc(t)}</h3>
      <div class="cot-tabla-card"><table class="cot-tabla pan-tabla"><thead><tr><th>Concepto</th><th class="cot-num">Valor</th></tr></thead>
        <tbody>${ls.map(([a, b]) => `<tr><td>${cotEsc(a)}</td><td class="cot-num">${cotEsc(b)}</td></tr>`).join('')}</tbody></table></div></div>` : '').join('');

  sec.innerHTML = `
    <div class="pan-barra">
      <button type="button" class="btn secundario btn-chico pan-volver" data-pan-volver>${epIco('volver')}<span>Volver</span></button>
      <div class="pan-titulo"><h2>${cotEsc(c.cliente)}</h2><span class="cot-estado cot-e-${est}">${EP_ESTADOS[est]}</span></div>
      <div class="pan-acciones">${acciones}</div>
    </div>
    <div class="pan-card"><dl class="pan-datos">${datos}</dl></div>
    ${imgs}
    ${tablas}`;
}

function epActualizarAcciones() {
  const n = epSel.size;
  const b = document.getElementById('epBtnPdfSel');
  b.hidden = !n; b.textContent = `Cotización PDF (${n})`;
  document.getElementById('epBtnLimpiarSel').hidden = !n;
  document.getElementById('epBtnRecalcular').hidden = !(typeof pgEsSuper === 'function' && pgEsSuper());
}

/* ---------- Crear / editar datos ---------- */
function epPintarPrevias() {
  document.getElementById('epPrevias').innerHTML = epImagenes.map((src, i) =>
    `<div class="cot-previa"><img src="${src}" alt="Imagen ${i + 1}"><button type="button" data-ep-quitar="${i}" aria-label="Quitar imagen ${i + 1}">✕</button></div>`).join('');
}

async function epManejarArchivos(e) {
  const archivos = [...e.target.files];
  e.target.value = '';
  epError('epErrorNueva', '');
  for (const f of archivos) {
    if (epImagenes.length >= COT_MAX_IMAGENES) { epError('epErrorNueva', `Máximo ${COT_MAX_IMAGENES} imágenes.`); break; }
    if (!f.type.startsWith('image/')) { epError('epErrorNueva', 'Solo se permiten imágenes.'); continue; }
    try { epImagenes.push(await cotComprimir(f)); }
    catch (err) { console.error(err); epError('epErrorNueva', 'No se pudo leer una de las imágenes.'); }
  }
  epPintarPrevias();
}

function epAbrirNueva() {
  epEditId = null;
  document.getElementById('epFormNueva').reset();
  document.getElementById('epNuevaTitulo').textContent = 'Nueva cotización EasyPack';
  document.getElementById('epBtnGuardar').textContent = 'Crear cotización';
  epImagenes = []; epPintarPrevias(); epError('epErrorNueva', '');
  document.getElementById('epModalNueva').hidden = false;
  document.getElementById('epCliente').focus();
}

function epAbrirEditar(id) {
  const c = epLista.find(x => x.id === id);
  if (!c || c.anulada) return;
  epEditId = id;
  document.getElementById('epFormNueva').reset();
  epError('epErrorNueva', '');
  document.getElementById('epNuevaTitulo').textContent = `Editar cotización EasyPack No. ${c.consecutivo}`;
  document.getElementById('epBtnGuardar').textContent = 'Guardar cambios';
  document.getElementById('epCliente').value = c.cliente || '';
  document.getElementById('epEmpresa').value = c.empresa || '';
  document.getElementById('epNit').value = c.nit || '';
  document.getElementById('epProductoEs').value = c.nombreProducto || '';
  document.getElementById('epProductoEn').value = c.nombreProductoEn || '';
  document.getElementById('epDescripcion').value = c.descripcion || '';
  document.getElementById('epCantidad').value = cotEsRango(c) ? `${c.cantidadMin}-${c.cantidad}` : String(c.cantidad ?? '');
  document.getElementById('epLink').value = c.link || '';
  epImagenes = (c.imagenes || []).slice(); epPintarPrevias();
  document.getElementById('epModalNueva').hidden = false;
  document.getElementById('epCliente').focus();
}

async function epGuardarNueva(e) {
  e.preventDefault();
  const cliente = document.getElementById('epCliente').value.trim();
  const nombreProducto = document.getElementById('epProductoEs').value.trim();
  const descripcion = document.getElementById('epDescripcion').value.trim();
  const cant = cotLeerCantidad(document.getElementById('epCantidad').value);
  const link = document.getElementById('epLink').value.trim();
  if (!cliente) return epError('epErrorNueva', 'Escribe el nombre del cliente.');
  if (!nombreProducto) return epError('epErrorNueva', 'Escribe el nombre del producto en español.');
  if (!descripcion) return epError('epErrorNueva', 'Escribe la descripción del producto.');
  if (!cant) return epError('epErrorNueva', 'La cantidad debe ser un número mayor que cero, o un rango como 50-100.');
  if (link && !cotLinkSeguro(link)) return epError('epErrorNueva', 'El link debe empezar con http:// o https://');
  if (epImagenes.reduce((s, i) => s + i.length, 0) > COT_MAX_CHARS_TOTAL) return epError('epErrorNueva', 'Las imágenes pesan demasiado. Quita alguna o usa imágenes más pequeñas.');
  const datos = {
    cliente, empresa: document.getElementById('epEmpresa').value.trim(), nit: document.getElementById('epNit').value.trim(),
    nombreProducto, nombreProductoEn: document.getElementById('epProductoEn').value.trim(), descripcion,
    cantidad: cant.max, cantidadMin: cant.min < cant.max ? cant.min : null, link, imagenes: epImagenes
  };
  const btn = document.getElementById('epBtnGuardar');
  btn.disabled = true;
  try {
    if (epEditId) {
      await cotRef().doc(epEditId).update({ ...datos, editadoPor: sesionActual ? sesionActual.nombre : '', editadoEn: firebase.firestore.FieldValue.serverTimestamp() });
    } else {
      await epCrear(datos);
    }
    epCerrar(document.getElementById('epModalNueva'));
  } catch (err) {
    console.error(err);
    epError('epErrorNueva', 'No se pudo guardar. Revisa tu conexión y las reglas de Firestore.');
  } finally { btn.disabled = false; }
}

/* Crea la cotización con el siguiente consecutivo (transacción: dos personas a la vez nunca repiten número). */
async function epCrear(datos) {
  const rc = db.collection(APP_CONFIG.COLECCIONES.configuracion).doc('consecutivos');
  const ref = cotRef().doc();
  let n = 0;
  await db.runTransaction(async tx => {
    const s = await tx.get(rc);
    const d = s.exists ? s.data() : {};
    n = (Number.isInteger(d.easypack) ? d.easypack : 0) + 1;
    tx.set(rc, { easypack: n }, { merge: true });
    tx.set(ref, {
      estado: 'pendiente', ...datos, origen: EP_ORIGEN, consecutivo: n,
      creadoPor: sesionActual ? sesionActual.nombre : '', creadoEn: firebase.firestore.FieldValue.serverTimestamp()
    });
  });
  return n;
}

/* ---------- Compra finalizada: la cotización pasa a Compras Finalizadas de EasyPack (pagos del cliente) ---------- */
async function epFinalizar(id) {
  const c = epLista.find(x => x.id === id);
  if (!c || c.anulada || !c.costosEp || c.modulo) return;
  if (!confirm(`¿Pasar la cotización No. ${c.consecutivo} de ${c.cliente} a Compras Finalizadas?

Desde allí se registran los pagos.`)) return;
  const datos = {
    estado: 'finalizada', modulo: 'easypack',
    enviadaEn: firebase.firestore.FieldValue.serverTimestamp(), enviadaPor: sesionActual ? sesionActual.nombre : '',
    finalizadaEn: firebase.firestore.FieldValue.serverTimestamp()
  };
  if (!Array.isArray(c.pagos)) datos.pagos = [];
  try { await cotRef().doc(id).update(datos); epSel.delete(id); }
  catch (err) { console.error(err); alert('No se pudo pasar la cotización a Compras Finalizadas.'); }
}

/* ---------- Anular y copiar ---------- */
async function epAnular(id) {
  const c = epLista.find(x => x.id === id);
  if (!c || c.anulada) return;
  if (!confirm(`¿Anular la cotización No. ${c.consecutivo} de ${c.cliente}?\n\nQueda en la lista como "Anulada" y conserva su consecutivo (no se reutiliza). Desde ella podrás crear una copia con el siguiente número.`)) return;
  try {
    await cotRef().doc(id).update({ anulada: true, anuladaPor: sesionActual ? sesionActual.nombre : '', anuladaEn: firebase.firestore.FieldValue.serverTimestamp() });
    epSel.delete(id);
  } catch (err) { console.error(err); alert('No se pudo anular la cotización.'); }
}

async function epCopiar(id) {
  const c = epLista.find(x => x.id === id);
  if (!c || !c.anulada) return;
  if (!confirm(`Se creará una copia de la cotización anulada No. ${c.consecutivo} con el siguiente consecutivo disponible.\n\n¿Continuar?`)) return;
  const { id: _id, consecutivo, anulada, anuladaPor, anuladaEn, creadoEn, creadoPor, editadoEn, editadoPor, ...datos } = c;
  try {
    const n = await epCrear({ ...datos, copiaDe: consecutivo });
    alert(`Copia creada con el consecutivo ${n}.`);
  } catch (err) { console.error(err); alert('No se pudo crear la copia.'); }
}

/* ---------- Formulario de costos (fórmulas del Excel) ---------- */
function epLeerCostos() {
  const v = id => parseFloat(document.getElementById(id).value);
  const fleteSel = document.getElementById('epFlete');
  const fleteValor = parseFloat(fleteSel.selectedOptions[0] && fleteSel.selectedOptions[0].dataset.valor);
  return {
    qtyCaja: v('epQtyCaja'), alto: v('epAlto'), ancho: v('epAncho'), largo: v('epLargo'), pesoCaja: v('epPesoUnit'),
    precioUsd: v('epPrecioUsd'), tipo: document.getElementById('epTipo').value, gasto: v('epGasto') || 0,
    fleteId: fleteSel.value, fleteBarco: fleteValor,
    modalidad: document.getElementById('epModalidad').value, stock: document.getElementById('epStock').checked
  };
}

function epRefrescarFletes(seleccionado) {
  const sel = document.getElementById('epFlete');
  const fletes = (typeof DATOS !== 'undefined' && DATOS.fletes) || [];
  sel.innerHTML = fletes.length
    ? fletes.map(f => `<option value="${cotEsc(f.id)}" data-valor="${f.valor}">${epNum(f.valor)}</option>`).join('')
    : '<option value="" data-valor="">Agrega un flete en Configuración › Datos</option>';
  if (seleccionado && fletes.some(f => f.id === seleccionado)) sel.value = seleccionado;
}

function epActualizarResumen() {
  const k = epLeerCostos();
  const el = document.getElementById('epResumen');
  const trm = typeof TASAS !== 'undefined' ? TASAS.usdCop : null;
  const faltan = [k.qtyCaja, k.alto, k.ancho, k.largo, k.pesoCaja, k.precioUsd, k.fleteBarco].some(x => !(x >= 0)) || !(k.qtyCaja > 0);
  if (faltan || !(trm > 0)) {
    el.innerHTML = !(trm > 0) ? 'No hay tasa del dólar a pesos todavía. Revisa Configuración › Datos › Tasas de cambio.' : 'Completa los datos para ver el cálculo.';
    return;
  }
  const r = epCalcular(k, trm);
  const fila = (t, val) => `<tr><td>${t}</td><td class="cot-num">${val}</td></tr>`;
  el.innerHTML = `<table class="ep-calc"><tbody>
    ${fila('Dólar a pesos', epCop(trm))}
    ${fila('CBM por caja', epNum(r.cbmCaja))}
    ${fila('Peso por caja', epNum(r.pesoCaja) + ' kg')}
    ${fila('Precio China (USD, informativo)', epNum(r.precioChinaUsd, 2))}
    ${fila('Costo del producto (COP, unidad)', epCop(r.costoProdCop))}
    ${fila('Flete marítimo (COP, unidad)', epCop(r.fleteMarUnidad))}
    ${fila('Flete aéreo (COP, unidad)', r.fleteAereoUnidad == null ? 'SIN PESO' : epCop(r.fleteAereoUnidad))}
    ${fila('IVA (informativo)', epCop(r.iva))}
    ${fila('Arancel (informativo)', epCop(r.arancel))}
    ${fila('Costo de importación barco (producto + flete)', epCop(r.costoImportBarco))}
    ${fila('<strong>Costo marítimo por unidad (nacionalizado)</strong>', '<strong>' + epCop(r.precioMaritimo) + '</strong>')}
    ${fila('<strong>Costo aéreo por unidad (producto + flete)</strong>', '<strong>' + (r.precioAereo == null ? 'SIN PESO' : epCop(r.precioAereo)) + '</strong>')}
  </tbody></table>
  <p>El precio del PDF es el de la modalidad elegida, sin impuestos.</p>`;
}

function epAbrirCostos(id) {
  const c = epLista.find(x => x.id === id);
  if (!c || c.anulada) return;
  epCostosId = id;
  document.getElementById('epFormCostos').reset();
  epError('epErrorCostos', '');
  document.getElementById('epCostosTitulo').textContent = `Costos · No. ${c.consecutivo} · ${c.cliente}`;
  const k = c.costosEp;
  epRefrescarFletes(k && k.fleteId);
  if (k) {
    const s = (i, v) => { document.getElementById(i).value = v ?? ''; };
    s('epQtyCaja', k.qtyCaja); s('epAlto', k.alto); s('epAncho', k.ancho); s('epLargo', k.largo); s('epPesoUnit', k.pesoCaja != null ? k.pesoCaja : (k.pesoUnit > 0 && k.qtyCaja > 0 ? k.pesoUnit * k.qtyCaja / 1000 : ''));
    s('epPrecioUsd', k.precioUsd); s('epTipo', k.tipo); s('epGasto', k.gasto); s('epModalidad', k.modalidad);
    document.getElementById('epStock').checked = !!k.stock;
  } else {
    document.getElementById('epGasto').value = EP_K.GASTO_DEFECTO.NORMAL;
  }
  epActualizarResumen();
  document.getElementById('epModalCostos').hidden = false;
  document.getElementById('epQtyCaja').focus();
}

async function epGuardarCostos(e) {
  e.preventDefault();
  const c = epLista.find(x => x.id === epCostosId);
  if (!c || c.anulada) return;
  const k = epLeerCostos();
  const trm = typeof TASAS !== 'undefined' ? TASAS.usdCop : null;
  if (!(k.qtyCaja >= 1) || !Number.isInteger(k.qtyCaja)) return epError('epErrorCostos', 'Las unidades por caja deben ser un número entero mayor que cero.');
  if (![k.alto, k.ancho, k.largo, k.pesoCaja, k.precioUsd].every(x => x >= 0)) return epError('epErrorCostos', 'Completa alto, ancho, largo, peso por caja y precio con valores válidos.');
  if (!(k.fleteBarco >= 0)) return epError('epErrorCostos', 'Elige un precio de flete marítimo. Si la lista está vacía, agrégalo en Configuración › Datos › Fletes.');
  if (!(trm > 0)) return epError('epErrorCostos', 'No hay tasa del dólar a pesos. Revisa Configuración › Datos › Tasas de cambio.');
  const r = epCalcular(k, trm);
  const precioUnidad = epPrecioUnidad(r, k.modalidad);
  if (precioUnidad == null) return epError('epErrorCostos', 'La modalidad aérea necesita el peso por caja (SIN PESO). Escríbelo o elige marítima.');
  const costosEp = { ...k, ...r, precioUnidad, usdCop: trm, calculadoPor: sesionActual ? sesionActual.nombre : '' };
  Object.keys(costosEp).forEach(key => { if (typeof costosEp[key] === 'number' && !Number.isFinite(costosEp[key])) costosEp[key] = null; });
  const btn = document.getElementById('epBtnGuardarCostos');
  btn.disabled = true;
  try {
    await cotRef().doc(epCostosId).update({ costosEp, estado: 'cotizada', editadoEn: firebase.firestore.FieldValue.serverTimestamp() });
    epCerrar(document.getElementById('epModalCostos'));
  } catch (err) {
    console.error(err);
    epError('epErrorCostos', 'No se pudieron guardar los costos. Revisa tu conexión y las reglas de Firestore.');
  } finally { btn.disabled = false; }
}

/* ---------- PDF (plantilla EasyPack) ---------- */
let epFondoDataUrl = null;
async function epCargarFondo() {
  if (epFondoDataUrl) return epFondoDataUrl;
  const blob = await (await fetch('easypack-fondo.png')).blob();
  epFondoDataUrl = await new Promise((ok, fallo) => { const l = new FileReader(); l.onload = () => ok(l.result); l.onerror = () => fallo(l.error); l.readAsDataURL(blob); });
  return epFondoDataUrl;
}

async function epGenerarPDF(ids) {
  const lista = ids.map(id => epLista.find(x => x.id === id)).filter(c => c && c.costosEp && !c.anulada).sort((a, b) => a.consecutivo - b.consecutivo);
  if (!lista.length) { alert('Selecciona cotizaciones con costos y no anuladas.'); return; }
  if (lista.some(c => c.cliente !== lista[0].cliente || (c.empresa || '') !== (lista[0].empresa || '') || (c.nit || '') !== (lista[0].nit || ''))) {
    alert('Para sacar un solo PDF, las cotizaciones deben ser del mismo cliente, empresa y NIT.'); return;
  }
  if (!window.jspdf) { alert('No se cargó la librería de PDF. Revisa tu conexión y recarga la página.'); return; }
  let fondo;
  try { fondo = await epCargarFondo(); } catch (err) { console.error(err); alert('No se pudo cargar la plantilla de EasyPack (easypack-fondo.png).'); return; }

  const doc = new window.jspdf.jsPDF({ unit: 'pt', format: 'letter' });
  const X = 68, DER = 545, TOPE = 690, LH = 15;
  let y = 0;
  const pagina = primera => { if (!primera) doc.addPage(); doc.addImage(fondo, 'PNG', 0, 0, 612, 792); y = primera ? 142 : 150; };
  const texto = (t, opc = {}) => {
    doc.setFont('helvetica', opc.negrita ? 'bold' : 'normal'); doc.setFontSize(10.5); doc.setTextColor(0);
    doc.splitTextToSize(String(t), DER - X).forEach(l => { if (y > TOPE) pagina(false); doc.text(l, X, y); y += LH; });
  };
  const espacio = (n = 1) => { y += LH * n; };

  pagina(true);
  const primera = lista[0];
  doc.setFont('helvetica', 'normal'); doc.setFontSize(10.5); doc.setTextColor(0);
  doc.text(epFechaHoy(), X, y);
  y += LH * 2;

  if (primera.empresa) {
    texto(`Señores ${primera.empresa}`);
    if (primera.nit) texto(primera.nit);
    espacio(); texto(primera.cliente);
  } else {
    texto(`Señor(a) ${primera.cliente}`);
    if (primera.nit) texto(primera.nit);
  }
  espacio(); texto('Cordial saludo,'); espacio();
  texto('De acuerdo con su solicitud, nos permitimos presentar la cotización de los productos de su interés. A continuación, encontrará el cuadro con las referencias, especificaciones y valores correspondientes.');
  espacio();
  texto('Esta propuesta ha sido elaborada teniendo en cuenta los estándares de calidad que manejamos y las condiciones comerciales vigentes. Los precios indicados no incluyen impuestos y son puestos en sus instalaciones.');
  const conStock = lista.filter(c => c.costosEp.stock).length;
  if (conStock === lista.length) { espacio(); texto(`En este momento contamos con stock disponible en ${lista.length > 1 ? 'cada uno de los productos cotizados' : 'el producto cotizado'}.`); }
  else if (conStock === 0) { espacio(); texto(`En el momento no contamos con stock disponible en ${lista.length > 1 ? 'ninguna de las referencias cotizadas' : 'la referencia cotizada'}.`); }
  espacio();

  /* Una tabla por cotización (cada una con su cantidad en el título del precio) + validez, modalidad y tiempo de entrega debajo */
  const COLS = [{ x: 89, w: 205 }, { x: 294, w: 105 }, { x: 399, w: 95 }];     // Producto | Unidad de empaque | Precio
  lista.forEach(c => {
    const k = c.costosEp;
    const precioTxt = `$${epNum(k.precioUnidad, 2)} pesos unidad`;
    const cab = ['Producto', 'Unidad de empaque', `Precio para ${cotCantidadTxt(c)} u.`];
    doc.setFontSize(10.5);
    doc.setFont('helvetica', 'bold');
    const cabLineas = cab.map((t, i) => doc.splitTextToSize(t, COLS[i].w - 8));
    doc.setFont('helvetica', 'normal');
    const celdas = [doc.splitTextToSize(c.nombreProducto, COLS[0].w - 8), doc.splitTextToSize(`${epNum(k.qtyCaja, 0)} unidades por caja`, COLS[1].w - 8), doc.splitTextToSize(precioTxt, COLS[2].w - 8)];
    const hCab = Math.max(...cabLineas.map(l => l.length)) * LH + 6;
    const hFila = Math.max(...celdas.map(l => l.length)) * LH + 6;
    if (y + hCab + hFila + LH * 3 > TOPE) pagina(false);
    const y0 = y - 11;
    doc.setLineWidth(0.5); doc.setDrawColor(0);
    COLS.forEach((col, i) => {
      doc.rect(col.x, y0, col.w, hCab); doc.rect(col.x, y0 + hCab, col.w, hFila);
      doc.setFont('helvetica', 'bold'); doc.setTextColor(0, 176, 240);
      cabLineas[i].forEach((l, j) => doc.text(l, col.x + col.w / 2, y0 + 13 + j * LH, { align: 'center' }));
      doc.setFont('helvetica', 'normal'); doc.setTextColor(0);
      celdas[i].forEach((l, j) => doc.text(l, col.x + 4, y0 + hCab + 13 + j * LH));
    });
    y = y0 + hCab + hFila + LH + 4;
    texto(`Validez de la oferta: ${EP_VALIDEZ_DIAS} días  ·  Modalidad: ${EP_MODALIDAD[k.modalidad]}  ·  Tiempo de entrega estimado (${EP_MODALIDAD[k.modalidad].toLowerCase()}): ${EP_ENTREGA_DIAS[k.modalidad]} días`);
    espacio(0.5);
  });

  espacio(0.5);
  texto('Quedamos atentos a cualquier inquietud adicional o a los ajustes que requiera la presente cotización. Será un gusto acompañarlos en su proceso de compra.');
  if (y + LH * 8 > TOPE) pagina(false);
  espacio(); texto('Cordialmente'); espacio();
  texto(EP_FIRMA.nombre); texto(EP_FIRMA.cargo); texto(EP_FIRMA.correo);

  doc.save(`cotizacion-easypack-${lista.map(c => c.consecutivo).join('-')}.pdf`);
}

/* ---------- Eventos ---------- */
document.addEventListener('DOMContentLoaded', () => {
  if (!epConstruirUI()) return;
  const $ = id => document.getElementById(id);
  document.querySelectorAll('[data-ep-cerrar]').forEach(el => el.addEventListener('click', () => epCerrar(el.closest('.modal'))));
  $('epBtnNueva').addEventListener('click', epAbrirNueva);
  $('epFormNueva').addEventListener('submit', epGuardarNueva);
  $('epFormCostos').addEventListener('submit', epGuardarCostos);
  $('epArchivos').addEventListener('change', epManejarArchivos);
  $('epPrevias').addEventListener('click', e => {
    const b = e.target.closest('[data-ep-quitar]');
    if (b) { epImagenes.splice(Number(b.dataset.epQuitar), 1); epPintarPrevias(); }
  });
  ['epBuscar', 'epFiltroEstado'].forEach(id => $(id).addEventListener(id === 'epBuscar' ? 'input' : 'change', () => { epPagina = 1; epPintar(); }));
  $('epBtnPdfSel').addEventListener('click', () => epGenerarPDF([...epSel]));
  $('epBtnLimpiarSel').addEventListener('click', () => { epSel.clear(); epPintar(); });
  $('epBtnRecalcular').addEventListener('click', epRecalcularTodas);

  $('epFormCostos').addEventListener('input', epActualizarResumen);
  $('epFormCostos').addEventListener('change', epActualizarResumen);
  $('epTipo').addEventListener('change', () => { $('epGasto').value = EP_K.GASTO_DEFECTO[$('epTipo').value]; epActualizarResumen(); });
  document.addEventListener('datos-actualizados', () => { if (!$('epModalCostos').hidden) { epRefrescarFletes($('epFlete').value); epActualizarResumen(); } });
  document.addEventListener('tasas-actualizadas', () => { if (!$('epModalCostos').hidden) epActualizarResumen(); });

  $('epGrid').addEventListener('change', e => {
    const s = e.target.closest('[data-ep-por-pagina]');
    if (s) { epPorPagina = Number(s.value) || 10; epPagina = 1; epPintar(); return; }
    const cb = e.target.closest('[data-ep-sel]');
    if (cb) { if (cb.checked) epSel.add(cb.dataset.epSel); else epSel.delete(cb.dataset.epSel); epActualizarAcciones(); }
  });
  $('epGrid').addEventListener('click', e => {
    const pg = e.target.closest('[data-ep-pag]');
    if (pg) { if (!pg.disabled) { epPagina += pg.dataset.epPag === 'next' ? 1 : -1; epPintar(); } return; }
    const g = a => { const b = e.target.closest(`[data-ep-${a}]`); return b ? b.dataset['ep' + a[0].toUpperCase() + a.slice(1)] : null; };
    let id;
    if (e.target.closest('[data-ep-sel]')) return;
    if ((id = g('costos'))) epAbrirCostos(id);
    else if ((id = g('pdf'))) epGenerarPDF([id]);
    else if ((id = g('editar'))) epAbrirEditar(id);
    else if ((id = g('anular'))) epAnular(id);
    else if ((id = g('copia'))) epCopiar(id);
    else if ((id = g('finalizar'))) epFinalizar(id);
    else if ((id = g('ver'))) epAbrirDetalle(id);
  });
  $('epGrid').addEventListener('keydown', e => {
    const tr = e.target.closest('tr[data-ep-ver]');
    if (tr && e.target === tr && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); epAbrirDetalle(tr.dataset.epVer); }
  });
  $('pan-ep').addEventListener('click', e => {
    if (e.target.closest('[data-pan-volver]')) { panVolver(); return; }
    const g = a => { const b = e.target.closest(`[data-ep-${a}]`); return b ? b.dataset['ep' + a[0].toUpperCase() + a.slice(1)] : null; };
    let id;
    if ((id = g('costos'))) epAbrirCostos(id);
    else if ((id = g('pdf'))) epGenerarPDF([id]);
    else if ((id = g('editar'))) epAbrirEditar(id);
    else if ((id = g('anular'))) epAnular(id);
    else if ((id = g('copia'))) epCopiar(id);
    else if ((id = g('finalizar'))) epFinalizar(id);
  });

  // Al cerrar sesión se cierran también las ventanas de EasyPack
  const cerrarOriginal = window.cotCerrarTodo;
  window.cotCerrarTodo = function () { epCerrarTodo(); if (typeof cerrarOriginal === 'function') cerrarOriginal.apply(this, arguments); };
});
