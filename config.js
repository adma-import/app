/* ===== CONFIGURACIÓN · Gestión de Importaciones ADMA ===== */
// ⚠️ Pega aquí los datos de TU NUEVO proyecto de Firebase (Configuración del proyecto > Tus apps > Web)
const firebaseConfig = {
  apiKey: "AIzaSyBobTQJn_yOwCqlBbaggCE88gMqtBf1l54",
  authDomain: "adma-1242a.firebaseapp.com",
  projectId: "adma-1242a",
  storageBucket: "adma-1242a.firebasestorage.app",
  messagingSenderId: "725655838152",
  appId: "1:725655838152:web:370fd8db7ce6f3f611b4fd",
};
firebase.initializeApp(firebaseConfig);
const auth = firebase.auth();
const db = firebase.firestore();

const APP_CONFIG = {
  NOMBRE_APP: 'Gestión de Importaciones ADMA',
  SUPER_USUARIO_EMAIL: 'juliana.admaimportaciones@gmail.com', // cuenta principal (siempre Super usuario, no se puede editar ni borrar)
  ROLES: { SUPER: 'Super usuario', ADMIN: 'Administrador', USUARIO: 'Usuario' },
  COLECCIONES: { usuarioEmail: 'usuario_email', usuarios: 'usuarios', configuracion: 'configuracion' },

  /* Módulos que se le pueden asignar a un "Usuario". Inicio y Configuración no están aquí:
     Inicio lo ven todos; Configuración solo Super usuario y Administrador.
     Para agregar un módulo nuevo: añádelo aquí (y crea su <section id="id">). */
MODULOS: [
  { id: 'cotizaciones', nombre: 'Cotizaciones' },
  { id: 'importaciones', nombre: 'Importaciones' },
  { id: 'proveedores', nombre: 'Proveedores' },
  { id: 'costos', nombre: 'Costos' }
],

  LOGO_MAX_BYTES: 700 * 1024, // Firestore limita cada documento a 1 MB y base64 pesa ~33% más
  LOGO_TIPOS_PERMITIDOS: ['image/png', 'image/svg+xml'],
  DOMINIO_CORREO_INTERNO: 'adma-importaciones.app', // correo sintético para usuarios sin correo real

  TASAS: {
    ZONA: 'America/Bogota',
    API_1: 'https://open.er-api.com/v6/latest/USD',
    API_2: 'https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@latest/v1/currencies/usd.min.json'
  },
  SELECTORES: { enlacesNav: '.sidebar nav a', tablist: '[role="tablist"]', tab: '.tab', btnMenu: '#btnMenu', overlay: '#overlay' }
};
