/* ==============================================================
   MaintenancePRO · usuarios.js
   ------------------------------------------------------------
   CRUD de usuarios del sistema (Configuración > pestaña Usuarios).

   Nota importante sobre Firebase:
   Crear una cuenta nueva con `auth.createUserWithEmailAndPassword()`
   en la app PRINCIPAL cerraría la sesión del administrador y lo
   dejaría logueado como el usuario recién creado (el SDK de
   cliente solo puede manejar una sesión a la vez). Para evitarlo,
   creamos las cuentas nuevas usando una SEGUNDA instancia de
   Firebase ("app secundaria"), completamente aislada de la
   sesión principal, y la cerramos apenas terminamos.

   Limitación conocida de este enfoque (sin backend/Cloud
   Functions): no es posible borrar la cuenta de Authentication
   de otra persona desde el cliente. Por eso, al "eliminar" un
   usuario, lo que hacemos es borrar su perfil (`usuarios/{uid}`)
   y su mapeo de acceso (`usuario_email/{usuario}`) — con eso ya
   NO puede iniciar sesión, aunque la cuenta técnica siga
   existiendo en Firebase Authentication. Si quieres purgarla del
   todo, tendrías que borrarla manualmente desde la consola
   (Authentication > Users).

   Depende de `db`, `auth`, `firebase` y `APP_CONFIG.firebaseConfig`
   (definidos en config.js), así que config.js debe cargarse antes.
   ============================================================== */

/* ---------- Referencias al DOM ---------- */
let listaUsuariosEl, modalUsuario, formUsuario, modalUsuarioTitulo, modalUsuarioError,
    usuarioUidEdicionInput, campoNombre, campoUsuario, campoRol,
    camposPasswordDiv, campoPassword, campoPasswordConfirmar,
    btnAgregarUsuario, btnGuardarUsuario, campoModulosWrap, campoModulosLista;

function cachearElementosUsuarios() {
  listaUsuariosEl = document.getElementById('listaUsuarios');
  modalUsuario = document.getElementById('modalUsuario');
  formUsuario = document.getElementById('formUsuario');
  modalUsuarioTitulo = document.getElementById('modalUsuarioTitulo');
  modalUsuarioError = document.getElementById('modalUsuarioError');
  usuarioUidEdicionInput = document.getElementById('usuarioUidEdicion');
  campoNombre = document.getElementById('campoNombre');
  campoUsuario = document.getElementById('campoUsuario');
  campoRol = document.getElementById('campoRol');
  camposPasswordDiv = document.getElementById('camposPassword');
  campoPassword = document.getElementById('campoPassword');
  campoPasswordConfirmar = document.getElementById('campoPasswordConfirmar');
  btnAgregarUsuario = document.getElementById('btnAgregarUsuario');
  btnGuardarUsuario = document.getElementById('btnGuardarUsuario');
  campoModulosWrap = document.getElementById('campoModulosWrap');
  campoModulosLista = document.getElementById('campoModulosLista');
}

/* ==============================================================
   App secundaria de Firebase (solo para crear cuentas nuevas
   sin afectar la sesión principal del administrador)
   ============================================================== */
function obtenerAppSecundaria() {
  let appSecundaria = firebase.apps.find(a => a.name === 'secundaria');
  if (!appSecundaria) {
    appSecundaria = firebase.initializeApp(firebaseConfig, 'secundaria');
  }
  return appSecundaria;
}

/* ==============================================================
   Utilidades
   ============================================================== */
function normalizarNombreUsuario(usuario) {
  return usuario.trim().toLowerCase();
}

function inicialesUsuario(nombre) {
  return nombre.split(' ').filter(Boolean).slice(0, 2)
    .map(p => p[0].toUpperCase()).join('') || '--';
}

function correoInternoDesdeUsuario(usuario) {
  return `${usuario}@${APP_CONFIG.DOMINIO_CORREO_INTERNO}`;
}

function mostrarErrorModal(mensaje) {
  modalUsuarioError.textContent = mensaje;
  modalUsuarioError.hidden = false;
}

function ocultarErrorModal() {
  modalUsuarioError.hidden = true;
  modalUsuarioError.textContent = '';
}

/* ==============================================================
   Permisos y módulos asignables
   ============================================================== */
// Super usuario gestiona a todos (menos la cuenta principal); Administrador solo a los de rol "Usuario".
function puedeGestionar(datos) {
  const R = APP_CONFIG.ROLES;
  if (!sesionActual || (datos.usuario || '').toLowerCase() === APP_CONFIG.SUPER_USUARIO_EMAIL) return false;
  if (sesionActual.rol === R.SUPER) return true;
  return sesionActual.rol === R.ADMIN && datos.rol === R.USUARIO;
}

function prepararRolesYModulos(rol, modulos) {
  const R = APP_CONFIG.ROLES;
  const permitidos = sesionActual && sesionActual.rol === R.SUPER ? [R.SUPER, R.ADMIN, R.USUARIO] : [R.USUARIO];
  campoRol.innerHTML = '';
  permitidos.forEach(r => {
    const o = document.createElement('option');
    o.value = o.textContent = r;
    campoRol.appendChild(o);
  });
  campoRol.value = permitidos.includes(rol) ? rol : R.USUARIO;

  campoModulosLista.innerHTML = '';
  APP_CONFIG.MODULOS.forEach(m => {
    const label = document.createElement('label');
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.value = m.id;
    cb.checked = modulos.includes(m.id);
    label.append(cb, document.createTextNode(m.nombre));
    campoModulosLista.appendChild(label);
  });
  actualizarVisibilidadModulos();
}

function actualizarVisibilidadModulos() {
  campoModulosWrap.hidden = campoRol.value !== APP_CONFIG.ROLES.USUARIO;
}

function modulosSeleccionados() {
  return [...campoModulosLista.querySelectorAll('input:checked')].map(i => i.value);
}

/* ==============================================================
   Renderizar la lista de usuarios
   ============================================================== */
function crearTarjetaUsuario(uid, datos) {
  const div = document.createElement('div');
  div.className = 'persona';
  div.setAttribute('role', 'listitem');
  div.dataset.uid = uid;

  const avatar = document.createElement('div');
  avatar.className = 'p-avatar';
  avatar.setAttribute('aria-hidden', 'true');
  avatar.textContent = inicialesUsuario(datos.nombre || datos.usuario || '');

  const h3 = document.createElement('h3');
  h3.textContent = datos.nombre || '(sin nombre)';

  const spanUsuario = document.createElement('span');
  spanUsuario.className = 'p-dato';
  spanUsuario.textContent = (datos.usuario || '').includes('@') ? datos.usuario : `@${datos.usuario || ''}`;

  const spanRol = document.createElement('span');
  spanRol.className = 'rol';
  spanRol.textContent = datos.rol || '';

  const acciones = document.createElement('div');
  acciones.className = 'doc-acciones';

  const btnEditar = document.createElement('button');
  btnEditar.className = 'btn secundario btn-chico';
  btnEditar.type = 'button';
  btnEditar.textContent = 'Editar';
  btnEditar.addEventListener('click', () => abrirModalEdicion(uid, datos));

  const btnEliminar = document.createElement('button');
  btnEliminar.className = 'btn peligro btn-chico';
  btnEliminar.type = 'button';
  btnEliminar.textContent = 'Eliminar';
  btnEliminar.addEventListener('click', () => eliminarUsuario(uid, datos));

  if (puedeGestionar(datos)) {
    acciones.appendChild(btnEditar);
    acciones.appendChild(btnEliminar);
  }

  div.appendChild(avatar);
  div.appendChild(h3);
  div.appendChild(spanUsuario);
  div.appendChild(spanRol);
  div.appendChild(acciones);
  return div;
}

function renderizarListaUsuarios(snapshot) {
  listaUsuariosEl.innerHTML = '';

  if (snapshot.empty) {
    const vacio = document.createElement('p');
    vacio.className = 'lista-vacia';
    vacio.textContent = 'Todavía no hay usuarios registrados.';
    listaUsuariosEl.appendChild(vacio);
    return;
  }

  snapshot.forEach(doc => {
    listaUsuariosEl.appendChild(crearTarjetaUsuario(doc.id, doc.data()));
  });
}

/* ---------- Escucha en tiempo real de la colección "usuarios" ---------- */
let desuscribirUsuarios = null;
function detenerEscuchaUsuarios() {
  if (desuscribirUsuarios) { desuscribirUsuarios(); desuscribirUsuarios = null; }
}
function initEscuchaUsuarios() {
  detenerEscuchaUsuarios();
  desuscribirUsuarios = db.collection(APP_CONFIG.COLECCIONES.usuarios)
    .orderBy('nombre')
    .onSnapshot(
      renderizarListaUsuarios,
      err => {
        console.error('Error al cargar usuarios:', err);
        listaUsuariosEl.innerHTML = '<p class="lista-vacia">No se pudo cargar la lista de usuarios.</p>';
      }
    );
}

/* ==============================================================
   Abrir / cerrar el modal
   ============================================================== */
function abrirModalCreacion() {
  formUsuario.reset();
  ocultarErrorModal();
  usuarioUidEdicionInput.value = '';
  modalUsuarioTitulo.textContent = 'Agregar usuario';
  prepararRolesYModulos('', []);
  campoUsuario.disabled = false;
  camposPasswordDiv.hidden = false;
  campoPassword.required = true;
  campoPasswordConfirmar.required = true;
  modalUsuario.hidden = false;
  campoNombre.focus();
}

function abrirModalEdicion(uid, datos) {
  formUsuario.reset();
  ocultarErrorModal();
  usuarioUidEdicionInput.value = uid;
  modalUsuarioTitulo.textContent = 'Editar usuario';
  campoNombre.value = datos.nombre || '';
  campoUsuario.value = datos.usuario || '';
  campoUsuario.disabled = true; // no permitimos cambiar el usuario de acceso desde aquí
  prepararRolesYModulos(datos.rol, datos.modulos || []);
  camposPasswordDiv.hidden = true; // la contraseña no se edita desde este formulario
  campoPassword.required = false;
  campoPasswordConfirmar.required = false;
  modalUsuario.hidden = false;
  campoNombre.focus();
}

function cerrarModalUsuario() {
  modalUsuario.hidden = true;
  formUsuario.reset();
  ocultarErrorModal();
}

/* ==============================================================
   Guardar (crear o editar) usuario
   ============================================================== */
async function manejarSubmitUsuario(e) {
  e.preventDefault();
  ocultarErrorModal();

  const uidEdicion = usuarioUidEdicionInput.value;
  const nombre = campoNombre.value.trim();
  const rol = campoRol.value;
  const modulos = rol === APP_CONFIG.ROLES.USUARIO ? modulosSeleccionados() : [];

  if (!nombre) {
    mostrarErrorModal('Escribe el nombre completo.');
    return;
  }

  btnGuardarUsuario.disabled = true;
  const textoOriginalBoton = btnGuardarUsuario.textContent;
  btnGuardarUsuario.textContent = 'Guardando…';

  try {
    if (uidEdicion) {
      await guardarEdicionUsuario(uidEdicion, { nombre, rol, modulos });
    } else {
      await crearUsuarioNuevo({ nombre, rol, modulos });
    }
    cerrarModalUsuario();
  } catch (err) {
    console.error('Error al guardar usuario:', err);
    mostrarErrorModal(mensajeErrorGuardarUsuario(err));
  } finally {
    btnGuardarUsuario.disabled = false;
    btnGuardarUsuario.textContent = textoOriginalBoton;
  }
}

async function guardarEdicionUsuario(uid, { nombre, rol, modulos }) {
  await db.collection(APP_CONFIG.COLECCIONES.usuarios).doc(uid).update({ nombre, rol, modulos });
}

async function crearUsuarioNuevo({ nombre, rol, modulos }) {
  const usuario = normalizarNombreUsuario(campoUsuario.value);
  const password = campoPassword.value;
  const passwordConfirmar = campoPasswordConfirmar.value;

  if (!usuario || !/^[a-z0-9._-]{3,}$/.test(usuario)) {
    throw { code: 'validacion/usuario', message: 'El usuario debe tener al menos 3 caracteres (letras, números, puntos o guiones, sin espacios).' };
  }
  if (password.length < 6) {
    throw { code: 'validacion/password', message: 'La contraseña debe tener al menos 6 caracteres.' };
  }
  if (password !== passwordConfirmar) {
    throw { code: 'validacion/password-confirmar', message: 'Las contraseñas no coinciden.' };
  }

  // 1. Verificar que el nombre de usuario no esté en uso.
  const refUsuarioEmail = db.collection(APP_CONFIG.COLECCIONES.usuarioEmail).doc(usuario);
  const existente = await refUsuarioEmail.get();
  if (existente.exists) {
    throw { code: 'validacion/usuario-existe', message: 'Ese nombre de usuario ya está en uso.' };
  }

  const email = correoInternoDesdeUsuario(usuario);
  const appSecundaria = obtenerAppSecundaria();
  const authSecundaria = appSecundaria.auth();

  // 2. Crear la cuenta en Firebase Authentication (en la app secundaria,
  //    así no se cierra la sesión del administrador en la app principal).
  const credencial = await authSecundaria.createUserWithEmailAndPassword(email, password);
  const uid = credencial.user.uid;

  // 3. Crear los documentos en Firestore. Si algo falla aquí, deshacemos
  //    la cuenta de Authentication (todavía tenemos su sesión activa en
  //    la app secundaria, así que sí podemos borrarla) para no dejarla
  //    huérfana. Al final, pase lo que pase, cerramos esa sesión secundaria.
  try {
    const batch = db.batch();
    batch.set(refUsuarioEmail, { email });
    batch.set(db.collection(APP_CONFIG.COLECCIONES.usuarios).doc(uid), {
      usuario, nombre, rol, modulos, activo: true,
      creadoEn: firebase.firestore.FieldValue.serverTimestamp()
    });
    await batch.commit();
  } catch (errFirestore) {
    console.error('Falló Firestore tras crear la cuenta de Auth, revirtiendo la cuenta creada:', errFirestore);
    await credencial.user.delete().catch(errDelete => {
      console.error('No se pudo revertir la cuenta de Auth huérfana. Bórrala manualmente desde la consola:', email, errDelete);
    });
    throw errFirestore;
  } finally {
    await authSecundaria.signOut().catch(() => {});
  }
}

function mensajeErrorGuardarUsuario(err) {
  if (err.code && err.code.startsWith('validacion/')) return err.message;
  switch (err.code) {
    case 'auth/email-already-in-use':
      return 'Ese nombre de usuario ya está en uso.';
    case 'auth/weak-password':
      return 'La contraseña es muy débil (mínimo 6 caracteres).';
    case 'auth/network-request-failed':
      return 'No hay conexión con el servidor. Revisa tu internet.';
    default:
      return 'No se pudo guardar el usuario. Inténtalo de nuevo.';
  }
}

/* ==============================================================
   Eliminar usuario
   ============================================================== */
async function eliminarUsuario(uid, datos) {
  if (sesionActual && sesionActual.uid === uid) {
    alert('No puedes eliminar tu propio usuario mientras tienes la sesión iniciada.');
    return;
  }

  const confirmado = confirm(
    `¿Eliminar a "${datos.nombre || datos.usuario}"?\n\n` +
    `Dejará de poder iniciar sesión de inmediato. ` +
    `Su cuenta técnica seguirá existiendo en Firebase Authentication ` +
    `(puedes borrarla ahí manualmente si lo necesitas).`
  );
  if (!confirmado) return;

  try {
    const batch = db.batch();
    batch.delete(db.collection(APP_CONFIG.COLECCIONES.usuarios).doc(uid));
    if (datos.usuario) {
      batch.delete(db.collection(APP_CONFIG.COLECCIONES.usuarioEmail).doc(datos.usuario));
    }
    await batch.commit();
  } catch (err) {
    console.error('Error al eliminar usuario:', err);
    alert('No se pudo eliminar el usuario. Inténtalo de nuevo.');
  }
}

/* ==============================================================
   Inicialización
   ============================================================== */
function initUsuarios() {
  cachearElementosUsuarios();

  btnAgregarUsuario.addEventListener('click', abrirModalCreacion);
  formUsuario.addEventListener('submit', manejarSubmitUsuario);
  campoRol.addEventListener('change', actualizarVisibilidadModulos);

  modalUsuario.querySelectorAll('[data-cerrar-modal]').forEach(el => {
    el.addEventListener('click', cerrarModalUsuario);
  });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && !modalUsuario.hidden) cerrarModalUsuario();
  });

}

document.addEventListener('DOMContentLoaded', initUsuarios);
