/* ===== Navegación, permisos por rol, logo y tasas de cambio ===== */

/* ---------- 1. Módulos (enlaces y secciones generados desde APP_CONFIG.MODULOS) ---------- */
function crearModulosDinamicos() {
  const nav = document.getElementById('navModulos');
  const enlaceConfig = nav.querySelector('[data-mod="config"]');
  const contenedor = document.querySelector('.modulos');
  APP_CONFIG.MODULOS.forEach(m => {
    const a = document.createElement('a');
    a.href = `#${m.id}`; a.dataset.mod = m.id; a.textContent = m.nombre;
    nav.insertBefore(a, enlaceConfig);
    const s = document.createElement('section');
    s.id = m.id; s.hidden = true;
    s.innerHTML = `<div class="panel"><div class="cabecera"><h2 style="border:none;background:none;padding:0">${m.nombre}</h2></div><div class="cuerpo"><p class="lista-vacia">Módulo en construcción.</p></div></div>`;
    contenedor.appendChild(s);
  });
}

function mostrarModulo(id) {
  document.querySelectorAll('.modulos > section').forEach(s => { s.hidden = s.id !== id; });
  document.querySelectorAll(APP_CONFIG.SELECTORES.enlacesNav).forEach(a => {
    if (a.dataset.mod) a.toggleAttribute('aria-current', a.dataset.mod === id);
    if (a.dataset.mod === id) a.setAttribute('aria-current', 'page');
  });
  cerrarMenu();
}

function initNavegacion() {
  document.querySelectorAll(APP_CONFIG.SELECTORES.enlacesNav).forEach(a => {
    if (!a.dataset.mod) return;
    a.addEventListener('click', e => { e.preventDefault(); mostrarModulo(a.dataset.mod); });
  });
}

/* ---------- 2. Permisos por rol (los llama auth.js al conocer el perfil) ---------- */
function modulosPermitidos(perfil) {
  const R = APP_CONFIG.ROLES;
  if (perfil.rol === R.SUPER || perfil.rol === R.ADMIN) return ['config', ...APP_CONFIG.MODULOS.map(m => m.id)];
  return [...(perfil.modulos || [])];
}

function aplicarPermisosPorRol(perfil) {
  const permitidos = modulosPermitidos(perfil);
  document.querySelectorAll(APP_CONFIG.SELECTORES.enlacesNav).forEach(a => {
    if (a.dataset.mod) a.style.display = permitidos.includes(a.dataset.mod) ? '' : 'none';
  });
  const primero = APP_CONFIG.MODULOS.find(m => permitidos.includes(m.id));
  mostrarModulo(primero ? primero.id : (permitidos.includes('config') ? 'config' : ''));
  if (permitidos.includes('config')) { initEscuchaUsuarios(); cargarLogoGuardado(); }
  if (['cotizaciones', 'ordenes-compra', 'adma-company', 'groupack'].some(m => permitidos.includes(m))) initEscuchaCotizaciones();
  initEscuchaTasas();
  initEscuchaDatos();
}

function alCerrarSesion() {
  detenerEscuchaUsuarios();
  detenerEscuchaCotizaciones();
  if (desuscribirTasas) { desuscribirTasas(); desuscribirTasas = null; }
  if (desuscribirDatos) { desuscribirDatos(); desuscribirDatos = null; }
}

/* ---------- 3. Menú móvil y pestañas ---------- */
function abrirMenu() {
  document.body.classList.add('menu-abierto');
  document.getElementById('overlay').hidden = false;
  document.getElementById('btnMenu').setAttribute('aria-expanded', 'true');
}
function cerrarMenu() {
  document.body.classList.remove('menu-abierto');
  const o = document.getElementById('overlay'), b = document.getElementById('btnMenu');
  if (o) o.hidden = true;
  if (b) b.setAttribute('aria-expanded', 'false');
}
function initMenuMovil() {
  document.getElementById('btnMenu').addEventListener('click', () =>
    document.body.classList.contains('menu-abierto') ? cerrarMenu() : abrirMenu());
  document.getElementById('overlay').addEventListener('click', cerrarMenu);
}

function initPestanas() {
  document.querySelectorAll(APP_CONFIG.SELECTORES.tablist).forEach(tablist => {
    const tabs = [...tablist.querySelectorAll(APP_CONFIG.SELECTORES.tab)];
    const paneles = tabs.map(t => document.getElementById(t.getAttribute('aria-controls')));
    const activar = i => {
      tabs.forEach((t, j) => { t.setAttribute('aria-selected', j === i); t.tabIndex = j === i ? 0 : -1; });
      paneles.forEach((p, j) => { if (p) p.hidden = j !== i; });
    };
    tabs.forEach((t, i) => t.addEventListener('click', () => activar(i)));
  });
}

/* ---------- 4. Logo (guardado como base64 en configuracion/empresa) ---------- */
function pintarLogo(src) {
  const img = document.getElementById('logoActualImg');
  img.src = src; img.style.display = 'block';
  document.getElementById('logoVacio').style.display = 'none';
  const lateral = document.getElementById('logoSidebar');
  lateral.src = src; lateral.hidden = false;
  document.querySelector('.sidebar .logo div').style.display = 'none';
}

async function cargarLogoGuardado() {
  try {
    const d = await db.collection(APP_CONFIG.COLECCIONES.configuracion).doc('empresa').get();
    if (d.exists && d.data().logoBase64) pintarLogo(d.data().logoBase64);
  } catch (err) { console.error('No se pudo cargar el logo:', err); }
}

function leerArchivoComoBase64(archivo) {
  return new Promise((ok, fallo) => {
    const l = new FileReader();
    l.onload = () => ok(l.result); l.onerror = () => fallo(l.error);
    l.readAsDataURL(archivo);
  });
}

async function manejarSeleccionLogo(e) {
  const archivo = e.target.files[0];
  if (!archivo) return;
  if (!APP_CONFIG.LOGO_TIPOS_PERMITIDOS.includes(archivo.type)) { alert('El logo debe ser PNG o SVG.'); e.target.value = ''; return; }
  if (archivo.size > APP_CONFIG.LOGO_MAX_BYTES) { alert('El logo no debe superar 700 KB.'); e.target.value = ''; return; }
  try {
    const base64 = await leerArchivoComoBase64(archivo);
    await db.collection(APP_CONFIG.COLECCIONES.configuracion).doc('empresa').set({
      logoBase64: base64, actualizadoEn: firebase.firestore.FieldValue.serverTimestamp()
    }, { merge: true });
    pintarLogo(base64);
  } catch (err) {
    console.error('Error al guardar el logo:', err);
    alert('No se pudo guardar el logo. Inténtalo de nuevo.');
  } finally { e.target.value = ''; }
}

/* ---------- 5. Tasas de cambio ----------
   Documento Firestore: configuracion/tasas
   { usdCop: {valor, modo:'auto'|'manual', fecha:'YYYY-MM-DD', fuente}, rmbUsd: {...} }
   Variable global para usar en el resto de la app:  TASAS.usdCop, TASAS.rmbUsd, TASAS.rmbCop
   (también obtenerTasas() y el evento 'tasas-actualizadas' en document). */
let TASAS = { usdCop: null, usdRmb: null, rmbCop: null, fecha: null };
let desuscribirTasas = null;
const refTasas = () => db.collection(APP_CONFIG.COLECCIONES.configuracion).doc('tasas');
const hoyBogota = () => new Date().toLocaleDateString('en-CA', { timeZone: APP_CONFIG.TASAS.ZONA });
function obtenerTasas() { return { ...TASAS }; }

async function consultarTasasOnline() {
  try {
    const j = await (await fetch(APP_CONFIG.TASAS.API_1)).json();
    if (j.result === 'success' && j.rates.COP && j.rates.CNY)
      return { usdCop: j.rates.COP, usdRmb: j.rates.CNY, fuente: 'open.er-api.com' };
  } catch (e) { console.warn('API 1 falló, probando la 2:', e); }
  const j = await (await fetch(APP_CONFIG.TASAS.API_2)).json();
  return { usdCop: j.usd.cop, usdRmb: j.usd.cny, fuente: 'currency-api' };
}

// Actualiza solo las tasas en modo automático que no se hayan actualizado hoy (o todas si forzar=true).
async function actualizarTasasAuto(forzar = false) {
  const snap = await refTasas().get();
  const d = snap.exists ? snap.data() : {};
  const hoy = hoyBogota();
  const pendientes = ['usdCop', 'usdRmb'].filter(k => {
    const t = d[k];
    return !(t && t.modo === 'manual') && (forzar || !t || t.fecha !== hoy);
  });
  if (!pendientes.length) return;
  const n = await consultarTasasOnline();
  const upd = {};
  pendientes.forEach(k => {
  upd[k] = { valor: Number(n[k].toFixed(k === 'usdCop' ? 2 : 4)), modo: 'auto', fecha: hoy, fuente: n.fuente };  });
  await refTasas().set(upd, { merge: true });
}

async function guardarTasaManual(k, valor) {
  await refTasas().set({ [k]: {
    valor, modo: 'manual', fecha: hoyBogota(), fuente: 'Manual · ' + (sesionActual ? sesionActual.nombre : '')
  } }, { merge: true });
}

const fmtTasa = (k, v) => v == null ? '—' : (k === 'usdCop'
  ? '$ ' + v.toLocaleString('es-CO', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
  : v.toLocaleString('es-CO', { minimumFractionDigits: 4, maximumFractionDigits: 6 }));
const metaTasa = t => t ? `${t.modo === 'manual' ? 'Manual' : 'Automática'} · ${t.fecha} · ${t.fuente}` : 'Sin datos todavía';

function pintarTasas(d) {
  TASAS.usdCop = d.usdCop ? d.usdCop.valor : null;
TASAS.usdRmb = d.usdRmb ? d.usdRmb.valor : null;
TASAS.rmbCop = TASAS.usdCop && TASAS.usdRmb ? Number((TASAS.usdCop / TASAS.usdRmb).toFixed(2)) : null;
TASAS.fecha = d.usdCop ? d.usdCop.fecha : null;

['usdCop', 'usdRmb'].forEach(k => {
    const card = document.querySelector(`.tasa-card[data-tasa="${k}"]`);
    card.querySelector('[data-valor]').textContent = fmtTasa(k, TASAS[k]);
    card.querySelector('[data-meta]').textContent = metaTasa(d[k]);
  });
  document.dispatchEvent(new CustomEvent('tasas-actualizadas', { detail: obtenerTasas() }));
}

function initEscuchaTasas() {
  if (desuscribirTasas) desuscribirTasas();
  desuscribirTasas = refTasas().onSnapshot(s => pintarTasas(s.exists ? s.data() : {}),
    err => console.error('Error al leer tasas:', err));
  actualizarTasasAuto().catch(err => console.error('No se pudieron actualizar las tasas:', err));
}

function initTasasUI() {
  document.getElementById('btnActualizarTasas').addEventListener('click', async e => {
    e.target.disabled = true;
    try { await actualizarTasasAuto(true); } catch (err) { alert('No se pudieron consultar las tasas. Revisa tu conexión.'); }
    e.target.disabled = false;
  });
  document.querySelectorAll('.tasa-card[data-tasa]').forEach(card => {
    const k = card.dataset.tasa;
    card.querySelector('[data-guardar]').addEventListener('click', async () => {
      const v = parseFloat(card.querySelector('[data-input]').value);
      if (!(v > 0)) { alert('Escribe un valor mayor que cero.'); return; }
      try { await guardarTasaManual(k, v); card.querySelector('[data-input]').value = ''; }
      catch (err) { console.error(err); alert('No se pudo guardar el valor.'); }
    });
    card.querySelector('[data-auto]').addEventListener('click', async () => {
      try {
        await refTasas().set({ [k]: { modo: 'auto', fecha: '', fuente: '', valor: TASAS[k] || 0 } }, { merge: true });
        await actualizarTasasAuto(true);
      } catch (err) { console.error(err); alert('No se pudo volver al modo automático.'); }
    });
  });
}

/* ---------- 6. Datos: fletes y comisión ADMA ----------
   Documento Firestore: configuracion/datos
   { fletes: [{ id, valor }], comisionAdma: 5 }
   Global: DATOS.fletes, DATOS.comisionAdma (y evento 'datos-actualizados' en document). */
const COMISION_ADMA_DEFECTO = 5;
let DATOS = { fletes: [], comisionAdma: COMISION_ADMA_DEFECTO };
let desuscribirDatos = null;
const refDatos = () => db.collection(APP_CONFIG.COLECCIONES.configuracion).doc('datos');

function pintarDatos(d) {
  DATOS.fletes = Array.isArray(d.fletes) ? d.fletes : [];
  DATOS.comisionAdma = typeof d.comisionAdma === 'number' ? d.comisionAdma : COMISION_ADMA_DEFECTO;

  const lista = document.getElementById('listaFletes');
  lista.replaceChildren();
  if (!DATOS.fletes.length) {
    const p = document.createElement('p');
    p.className = 'lista-vacia'; p.textContent = 'Aún no hay fletes.';
    lista.appendChild(p);
  }
  DATOS.fletes.forEach(f => {
    const fila = document.createElement('div');
    fila.setAttribute('role', 'listitem');
    fila.style.cssText = 'display:flex;justify-content:space-between;align-items:center;gap:var(--s3);padding:var(--s2) 0;border-bottom:1px solid var(--border)';
    const txt = document.createElement('span');
    txt.textContent = Number(f.valor).toLocaleString('es-CO', { maximumFractionDigits: 2 });
    const del = document.createElement('button');
    del.type = 'button'; del.className = 'btn secundario btn-chico'; del.textContent = 'Eliminar';
    del.setAttribute('aria-label', `Eliminar flete ${f.valor}`);
    del.addEventListener('click', () => eliminarFlete(f));
    fila.append(txt, del);
    lista.appendChild(fila);
  });

  document.getElementById('comisionValor').textContent =
    DATOS.comisionAdma.toLocaleString('es-CO', { maximumFractionDigits: 2 }) + ' %';
  document.dispatchEvent(new CustomEvent('datos-actualizados', { detail: { ...DATOS } }));
}

function initEscuchaDatos() {
  if (desuscribirDatos) desuscribirDatos();
  desuscribirDatos = refDatos().onSnapshot(s => pintarDatos(s.exists ? s.data() : {}),
    err => console.error('Error al leer datos (fletes/comisión):', err));
}

async function agregarFlete() {
  const valorEl = document.getElementById('fleteValor');
  const valor = parseFloat(valorEl.value);
  if (!(valor >= 0)) { alert('Escribe un valor válido para el flete.'); return; }
  const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  try {
    await refDatos().set({ fletes: firebase.firestore.FieldValue.arrayUnion({ id, valor }) }, { merge: true });
    valorEl.value = '';
  } catch (err) { console.error(err); alert('No se pudo guardar el flete.'); }
}

async function eliminarFlete(f) {
  if (!confirm(`¿Eliminar el flete ${f.valor}? Las cotizaciones que ya lo usan conservan su valor.`)) return;
  try { await refDatos().update({ fletes: firebase.firestore.FieldValue.arrayRemove(f) }); }
  catch (err) { console.error(err); alert('No se pudo eliminar el flete.'); }
}

async function guardarComision() {
  const input = document.getElementById('comisionInput');
  const v = parseFloat(input.value);
  if (!(v >= 0 && v <= 100)) { alert('Escribe un porcentaje entre 0 y 100.'); return; }
  try { await refDatos().set({ comisionAdma: v }, { merge: true }); input.value = ''; }
  catch (err) { console.error(err); alert('No se pudo guardar la comisión.'); }
}

/* Acordeón de Configuración > Datos: todo colapsado; al abrir una fila se cierran las demás. */
function initAcordeonDatos() {
  const btns = [...document.querySelectorAll('#panel-datos .acordeon-btn')];
  const fijar = (b, abierto) => {
    b.setAttribute('aria-expanded', String(abierto));
    document.getElementById(b.getAttribute('aria-controls')).hidden = !abierto;
  };
  btns.forEach(b => b.addEventListener('click', () => {
    const abrir = b.getAttribute('aria-expanded') !== 'true';
    btns.forEach(o => fijar(o, false));
    fijar(b, abrir);
  }));
}

function initDatosUI() {
  initAcordeonDatos();
  document.getElementById('btnAgregarFlete').addEventListener('click', agregarFlete);
  document.getElementById('btnGuardarComision').addEventListener('click', guardarComision);
}

/* ---------- Inicialización ---------- */
document.addEventListener('DOMContentLoaded', () => {
  crearModulosDinamicos();
  initNavegacion();
  initMenuMovil();
  initPestanas();
  initTasasUI();
  initDatosUI();
  document.getElementById('logo-file').addEventListener('change', manejarSeleccionLogo);
});
