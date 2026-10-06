/* ===== Módulo Cotizaciones =====
   Colección Firestore: cotizaciones
   {
     cliente, descripcion, cantidad, link, imagenes: [dataURL JPEG comprimido],
     estado: 'pendiente' | 'cotizada' | 'finalizada',
     costos: { valorUnidad, unidadesPorCaja, aumentoRmb, cbmCaja },
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
    <div class="cot-barra"><h2 class="cot-titulo">Cotizaciones</h2></div>
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
      <label class="login-campo"><span>Valor por unidad del producto (RMB)</span><input type="number" id="cotValorUnidad" min="0" step="0.01" required></label>
      <label class="login-campo"><span>Cantidad de unidades por caja</span><input type="number" id="cotUnidadesCaja" min="1" step="1" required></label>
      <label class="login-campo"><span>Aumento de RMB por unidad</span><input type="number" id="cotAumento" min="0" step="0.01" required></label>
      <label class="login-campo"><span>CBM por caja</span><input type="number" id="cotCbm" min="0" step="0.0001" required></label>
      <div class="modal-acciones">
        <button type="button" class="btn secundario" data-cot-cerrar>Cancelar</button>
        <button type="submit" class="btn" id="cotBtnGuardarCostos">Guardar</button>
      </div>
    </form>
  </div>`;
  document.getElementById('appShell').appendChild(modales);
  return true;
}

/* ---------- Abrir / cerrar modales ---------- */
const cotAbrir = id => { document.getElementById(id).hidden = false; };
function cotCerrarTodo() {
  document.querySelectorAll('#cotModalNueva, #cotModalCostos').forEach(m => { m.hidden = true; });
}
function cotError(id, msg) {
  const el = document.getElementById(id);
  el.textContent = msg || ''; el.hidden = !msg;
}

/* ---------- Imágenes: leer, comprimir, previsualizar ---------- */
function cotComprimir(archivo) {
  return new Promise((ok, fallo) => {
    const lector = new FileReader();
    lector.onerror = () => fallo(lector.error);
    lector.onload = () => {
      const img = new Image();
      img.onerror = () => fallo(new Error('Imagen no válida'));
      img.onload = () => {
        const k = Math.min(1, COT_MAX_LADO_PX / Math.max(img.width, img.height));
        const c = document.createElement('canvas');
        c.width = Math.round(img.width * k); c.height = Math.round(img.height * k);
        const ctx = c.getContext('2d');
        ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, c.width, c.height); // fondo blanco para PNG con transparencia
        ctx.drawImage(img, 0, 0, c.width, c.height);
        ok(c.toDataURL('image/jpeg', 0.72));
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
  cotAbrir('cotModalNueva');
  document.getElementById('cotCliente').focus();
}

async function cotGuardarNueva(e) {
  e.preventDefault();
  const cliente = document.getElementById('cotCliente').value.trim();
  const descripcion = document.getElementById('cotDescripcion').value.trim();
  const cantidad = parseInt(document.getElementById('cotCantidad').value, 10);
  const link = document.getElementById('cotLink').value.trim();

  if (!cliente || !descripcion) return cotError('cotErrorNueva', 'Escribe el nombre del cliente y la descripción.');
  if (!(cantidad > 0)) return cotError('cotErrorNueva', 'La cantidad debe ser mayor que cero.');
  if (link && !cotLinkSeguro(link)) return cotError('cotErrorNueva', 'El link debe empezar por http:// o https://');
  if (cotImagenesForm.join('').length > COT_MAX_CHARS_TOTAL) return cotError('cotErrorNueva', 'Las imágenes pesan demasiado. Quita alguna.');

  const btn = document.getElementById('cotBtnGuardar');
  btn.disabled = true;
  try {
    await cotRef().add({
      cliente, descripcion, cantidad, link, imagenes: cotImagenesForm,
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
function cotAbrirCostos(id) {
  const c = cotLista.find(x => x.id === id);
  if (!c) return;
  cotEditandoId = id;
  document.getElementById('cotFormCostos').reset();
  cotError('cotErrorCostos', '');
  document.getElementById('cotCostosTitulo').textContent = 'Costos · ' + c.cliente;
  if (c.costos) {
    document.getElementById('cotValorUnidad').value = c.costos.valorUnidad;
    document.getElementById('cotUnidadesCaja').value = c.costos.unidadesPorCaja;
    document.getElementById('cotAumento').value = c.costos.aumentoRmb;
    document.getElementById('cotCbm').value = c.costos.cbmCaja;
  }
  cotAbrir('cotModalCostos');
  document.getElementById('cotValorUnidad').focus();
}

async function cotGuardarCostos(e) {
  e.preventDefault();
  const costos = {
    valorUnidad: parseFloat(document.getElementById('cotValorUnidad').value),
    unidadesPorCaja: parseInt(document.getElementById('cotUnidadesCaja').value, 10),
    aumentoRmb: parseFloat(document.getElementById('cotAumento').value),
    cbmCaja: parseFloat(document.getElementById('cotCbm').value)
  };
  if (!(costos.valorUnidad >= 0) || !(costos.unidadesPorCaja > 0) || !(costos.aumentoRmb >= 0) || !(costos.cbmCaja >= 0))
    return cotError('cotErrorCostos', 'Completa los cuatro campos con valores válidos.');

  const btn = document.getElementById('cotBtnGuardarCostos');
  btn.disabled = true;
  try {
    await cotRef().doc(cotEditandoId).update({ costos, estado: 'cotizada' });
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

function cotCalculos(c) {
  const k = c.costos;
  if (!k) return null;
  const precioUnidad = k.valorUnidad + k.aumentoRmb;
  const cajas = Math.ceil(c.cantidad / k.unidadesPorCaja);
  return { precioUnidad, totalRmb: precioUnidad * c.cantidad, cajas, cbmTotal: cajas * k.cbmCaja };
}

/* ---------- Tarjetas ---------- */
function cotPintar() {
  const grid = document.getElementById('cotGrid');
  if (!cotLista.length) {
    grid.innerHTML = '<p class="lista-vacia">Aún no hay cotizaciones. Toca el botón + para crear la primera.</p>';
    return;
  }
  grid.innerHTML = cotLista.map(c => {
    const est = cotEstadoDe(c);
    const link = cotLinkSeguro(c.link);
    const k = c.costos;
    const costos = k
      ? `<p class="cot-costos">${cotNum(k.valorUnidad)} RMB/u · +${cotNum(k.aumentoRmb)} RMB · ${cotNum(k.unidadesPorCaja, 0)} u/caja · ${cotNum(k.cbmCaja, 4)} CBM/caja</p>` : '';

    let botones = `<button type="button" class="btn secundario btn-chico" data-cot-pdf="${c.id}" data-tipo="cotizar">A cotizar - Generar PDF</button>`;
    if (est === 'pendiente') {
      botones += `<button type="button" class="btn btn-chico" data-cot-continuar="${c.id}">Continuar</button>`;
    } else {
      botones += `<button type="button" class="btn secundario btn-chico" data-cot-pdf="${c.id}" data-tipo="cotizacion">Cotización - Generar PDF</button>`;
      botones += est === 'cotizada'
        ? `<button type="button" class="btn btn-chico" data-cot-orden="${c.id}">Generar orden de compra</button>
           <button type="button" class="cot-enlace" data-cot-continuar="${c.id}">Editar costos</button>`
        : `<button type="button" class="btn btn-chico" data-cot-orden="${c.id}">Orden de compra - Generar PDF</button>`;
    }

    return `<article class="cot-card cot-${est}">
      <header><h3>${cotEsc(c.cliente)}</h3><span class="cot-estado">${COT_ESTADOS[est]}</span></header>
      <p class="cot-desc">${cotEsc(c.descripcion)}</p>
      <p class="cot-dato"><strong>Cantidad:</strong> ${cotNum(c.cantidad, 0)}</p>
      ${link ? `<p class="cot-dato"><a href="${cotEsc(link)}" target="_blank" rel="noopener noreferrer">Ver link del producto</a></p>` : ''}
      ${(c.imagenes || []).length ? `<div class="cot-miniaturas">${c.imagenes.map((s, i) => `<img src="${s}" alt="Referencia ${i + 1} de ${cotEsc(c.cliente)}">`).join('')}</div>` : ''}
      ${costos}
      <footer>${botones}
        <button type="button" class="cot-eliminar" data-cot-eliminar="${c.id}" aria-label="Eliminar cotización de ${cotEsc(c.cliente)}">Eliminar</button>
      </footer>
    </article>`;
  }).join('');
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

  const calc = cotCalculos(c);
  if (calc && (tipo === 'cotizacion' || tipo === 'orden')) {
    const k = c.costos;
    escribir('Costos', 10, 'bold');
    escribir(`Valor por unidad: ${cotNum(k.valorUnidad)} RMB`, 12);
    escribir(`Aumento por unidad: ${cotNum(k.aumentoRmb)} RMB`, 12);
    escribir(`Precio por unidad: ${cotNum(calc.precioUnidad)} RMB`, 12, 'bold');
    escribir(`Total: ${cotNum(calc.totalRmb)} RMB`, 12, 'bold'); y += 2;
    escribir(`Unidades por caja: ${cotNum(k.unidadesPorCaja, 0)}`, 12);
    escribir(`Cajas: ${cotNum(calc.cajas, 0)}`, 12);
    escribir(`CBM por caja: ${cotNum(k.cbmCaja, 4)}`, 12);
    escribir(`CBM total: ${cotNum(calc.cbmTotal, 4)}`, 12); y += 4;
  }

  const imgs = c.imagenes || [];
  if (imgs.length) {
    escribir('Imágenes de referencia', 10, 'bold'); y += 2;
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

/* ---------- Escucha en tiempo real ---------- */
function initEscuchaCotizaciones() {
  if (!document.getElementById('cotGrid')) return;
  detenerEscuchaCotizaciones();
  cotDesuscribir = cotRef().orderBy('creadoEn', 'desc').onSnapshot(snap => {
    cotLista = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    cotPintar();
  }, err => {
    console.error('Error al leer cotizaciones:', err);
    document.getElementById('cotGrid').innerHTML = '<p class="lista-vacia">No se pudieron cargar las cotizaciones. Revisa las reglas de Firestore.</p>';
  });
}
function detenerEscuchaCotizaciones() {
  if (cotDesuscribir) { cotDesuscribir(); cotDesuscribir = null; }
  cotLista = [];
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

  document.getElementById('cotGrid').addEventListener('click', e => {
    const pdf = e.target.closest('[data-cot-pdf]');
    const ord = e.target.closest('[data-cot-orden]');
    const cont = e.target.closest('[data-cot-continuar]');
    const del = e.target.closest('[data-cot-eliminar]');
    if (pdf) cotGenerarPDF(pdf.dataset.cotPdf, pdf.dataset.tipo);
    else if (ord) cotGenerarOrden(ord.dataset.cotOrden);
    else if (cont) cotAbrirCostos(cont.dataset.cotContinuar);
    else if (del) cotEliminar(del.dataset.cotEliminar);
  });
});
