// URL de la Web App de Apps Script (termina en /exec).
// No es un secreto por sí sola: sin la contraseña la API no devuelve nada.
self.API_URL = "https://script.google.com/macros/s/AKfycbyzYNv9dKWNPKTBxgtiLYyXQMxBeJgJ897JiYwJSANjMAydVfSBwPppG2t90T1_QZXcVg/exec";

// Firebase (solo para avisos push). Estos valores son públicos por diseño; el secreto
// (la cuenta de servicio) vive en las Propiedades del script y nunca pasa por acá.
self.FIREBASE_CONFIG = {
  apiKey: "AIzaSyDeyRiER6OjkepzmkJpgSPgyjDehuCOfDk",
  authDomain: "finanzas-af5af.firebaseapp.com",
  projectId: "finanzas-af5af",
  storageBucket: "finanzas-af5af.firebasestorage.app",
  messagingSenderId: "783120389736",
  appId: "1:783120389736:web:dcf44a02e68bbc14ebcddd"
};
self.VAPID_KEY = "BK8R8UyVm6G39mOwLc1iuam5Cg4I1jGC0RXwETHX1ihCYRAGye1cdOcWKIOlkRWcdJu2x--E39UWd5uq2ON9vdA";
