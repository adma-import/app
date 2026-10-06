
let sesionActual = null; // Perfil del usuario logueado: { uid, usuario, nombre, rol, ... }

/* ---------- Referencias al DOM (se llenan en initAuth) ---------- */
let loginScreen, appShell, formLogin, loginUsuarioInput, loginPasswordInput,
    loginError, btnLogin, btnLoginTexto, btnCerrarSesion,
    usuarioNombreEl, usuarioRolEl, usuarioAvatarEl,
    usuarioNombreMovilEl, usuarioRolMovilEl, usuarioAvatarMovilEl;

function cachearElementosAuth() {
  loginScreen = document.getElementById('loginScreen');
  appShell = document.getElementById('appShell');
  formLogin = document.getElementById('formLogin');
  loginUsuarioInput = document.getElementById('loginUsuario');
  loginPasswordInput = document.getElementById('loginPassword');
  loginError = document.getElementById('loginError');
  btnLogin = document.getElementById('btnLogin');
  btnLoginTexto = document.getElementById('btnLoginTexto');
  btnCerrarSesion = document.getElementById('btnCerrarSesion');
  usuarioNombreEl = document.getElementById('usuarioNombre');
  usuarioRolEl = document.getElementById('usuarioRol');
  usuarioAvatarEl = document.getElementById('usuarioAvatar');
  usuarioNombreMovilEl = document.getElementById('usuarioNombreMovil');
  usuarioRolMovilEl = document.getElementById('usuarioRolMovil');
  usuarioAvatarMovilEl = document.getElementById('usuarioAvatarMovil');
}

/* ==============================================================
   Utilidades
   ============================================================== */
function normalizarUsuario(usuario) {
  return usuario.trim().toLowerCase();
}

function iniciales(nombre) {
  return nombre.split(' ').filter(Boolean).slice(0, 2)
    .map(palabra => palabra[0].toUpperCase()).join('') || '--';
}

function mostrarErrorLogin(mensaje) {
  loginError.textContent = mensaje;
  loginError.hidden = false;
}

function ocultarErrorLogin() {
  loginError.hidden = true;
  loginError.textContent = '';
}

function ponerCargandoLogin(cargando) {
  btnLogin.disabled = cargando;
  btnLoginTexto.textContent = cargando ? 'Ingresando…' : 'Iniciar sesión';
}

/* ==============================================================
   Mostrar / ocultar pantallas
   ============================================================== */
function mostrarApp() {
  loginScreen.hidden = true;
  appShell.hidden = false;
}

function mostrarLogin() {
  appShell.hidden = true;
  loginScreen.hidden = false;
  formLogin.reset();
  ocultarErrorLogin();
  ponerCargandoLogin(false);
  loginUsuarioInput.focus();
}

function ocultarSplash() {
  const splash = document.getElementById('bootSplash');
  if (splash) splash.hidden = true;
}

/* ==============================================================
   Pintar los datos del usuario logueado en la topbar
   ============================================================== */
function pintarUsuarioEnUI(perfil) {
  usuarioNombreEl.textContent = perfil.nombre || perfil.usuario;
  usuarioRolEl.textContent = perfil.rol || '';
  usuarioAvatarEl.textContent = iniciales(perfil.nombre || perfil.usuario);

  usuarioNombreMovilEl.textContent = perfil.nombre || perfil.usuario;
  usuarioRolMovilEl.textContent = perfil.rol || '';
  usuarioAvatarMovilEl.textContent = iniciales(perfil.nombre || perfil.usuario);
}

/* ==============================================================
   Firestore: usuario -> correo / perfil
   ============================================================== */
async function buscarEmailPorUsuario(usuario) {
  const doc = await db.collection(APP_CONFIG.COLECCIONES.usuarioEmail).doc(usuario).get();
  if (!doc.exists) return null;
  return doc.data().email || null;
}

async function cargarPerfilUsuario(uid) {
  const doc = await db.collection(APP_CONFIG.COLECCIONES.usuarios).doc(uid).get();
  if (!doc.exists) return null;
  return { uid, ...doc.data() };
}

async function crearPerfilSuperUsuario(firebaseUser) {
  const datos = { usuario: firebaseUser.email.toLowerCase(), nombre: 'Super usuario', rol: APP_CONFIG.ROLES.SUPER, activo: true, modulos: [] };
  try {
    await db.collection(APP_CONFIG.COLECCIONES.usuarios).doc(firebaseUser.uid).set({
      ...datos, creadoEn: firebase.firestore.FieldValue.serverTimestamp()
    });
    return { uid: firebaseUser.uid, ...datos };
  } catch (err) {
    console.error('No se pudo crear el perfil del super usuario (¿publicaste las reglas?):', err);
    return null;
  }
}

/* ==============================================================
   Envío del formulario de login
   ============================================================== */
async function manejarSubmitLogin(e) {
  e.preventDefault();
  ocultarErrorLogin();

  const usuario = normalizarUsuario(loginUsuarioInput.value);
  const password = loginPasswordInput.value;

  if (!usuario || !password) {
    mostrarErrorLogin('Escribe tu usuario y tu contraseña.');
    return;
  }

  ponerCargandoLogin(true);
  try {
    const email = usuario.includes('@') ? usuario : await buscarEmailPorUsuario(usuario);
    if (!email) {
      // Mensaje genérico a propósito: no revelamos si el usuario existe.
      mostrarErrorLogin('Usuario o contraseña incorrectos.');
      return;
    }
    await auth.signInWithEmailAndPassword(email, password);
    // A partir de aquí, onAuthStateChanged (más abajo) se encarga de
    // cargar el perfil y mostrar la aplicación.
  } catch (err) {
    console.error('Error al iniciar sesión:', err);
    mostrarErrorLogin(mensajeErrorAuth(err));
  } finally {
    ponerCargandoLogin(false);
  }
}

function mensajeErrorAuth(err) {
  switch (err.code) {
    case 'auth/wrong-password':
    case 'auth/user-not-found':
    case 'auth/invalid-credential':
    case 'auth/invalid-email':
      return 'Usuario o contraseña incorrectos.';
    case 'auth/too-many-requests':
      return 'Demasiados intentos. Espera unos minutos e inténtalo de nuevo.';
    case 'auth/network-request-failed':
      return 'No hay conexión con el servidor. Revisa tu internet.';
    case 'auth/user-disabled':
      return 'Este usuario está deshabilitado. Contacta a un administrador.';
    default:
      return 'No se pudo iniciar sesión. Inténtalo de nuevo.';
  }
}

/* ==============================================================
   Cerrar sesión
   ============================================================== */
async function manejarCerrarSesion(e) {
  e.preventDefault();
  try {
    await auth.signOut();
    // onAuthStateChanged se encarga de mostrar la pantalla de login.
  } catch (err) {
    console.error('Error al cerrar sesión:', err);
  }
}

/* ==============================================================
   Escucha del estado de sesión (mantiene la sesión al recargar)
   ============================================================== */
function initEscuchaSesion() {
  auth.onAuthStateChanged(async firebaseUser => {
    if (firebaseUser) {
      const esPrincipal = (firebaseUser.email || '').toLowerCase() === APP_CONFIG.SUPER_USUARIO_EMAIL;
      let perfil = await cargarPerfilUsuario(firebaseUser.uid);
      if (!perfil && esPrincipal) perfil = await crearPerfilSuperUsuario(firebaseUser);
      if (perfil && esPrincipal) perfil.rol = APP_CONFIG.ROLES.SUPER;
      if (!perfil || perfil.activo === false) {
        // Existe una cuenta válida en Firebase Auth pero no tiene
        // perfil en Firestore (usuarios/{uid}) — no debería pasar
        // si seguiste los pasos de creación del usuario semilla.
        console.error('No se encontró el perfil en Firestore para este usuario.');
        await auth.signOut();
        mostrarLogin();
        mostrarErrorLogin('Tu cuenta no tiene un perfil configurado. Contacta a un administrador.');
        ocultarSplash();
        return;
      }
      sesionActual = perfil;
      pintarUsuarioEnUI(perfil);
      if (typeof aplicarPermisosPorRol === 'function') aplicarPermisosPorRol(perfil);
      mostrarApp();
    } else {
      sesionActual = null;
      if (typeof alCerrarSesion === 'function') alCerrarSesion();
      mostrarLogin();
    }
    ocultarSplash();
  });
}

/* ==============================================================
   Inicialización
   ============================================================== */
function initAuth() {
  cachearElementosAuth();
  formLogin.addEventListener('submit', manejarSubmitLogin);
  btnCerrarSesion.addEventListener('click', manejarCerrarSesion);
  initEscuchaSesion();
}

document.addEventListener('DOMContentLoaded', initAuth);
