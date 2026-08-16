/*
 * Stockage en ligne avec synchronisation en temps réel (Firebase Firestore).
 *
 * - Quand un appareil modifie une donnée, tous les autres appareils connectés
 *   la reçoivent automatiquement en quelques secondes (window.storage.subscribe).
 * - Firestore garde aussi une copie locale sur l'appareil : si la connexion
 *   coupe, l'application continue de fonctionner normalement, et tout ce qui
 *   a été modifié hors-ligne est envoyé automatiquement dès que la connexion
 *   revient — sans aucune action de l'utilisateur.
 *
 * Même interface que l'ancien stockage local (get/set/delete/list), plus une
 * méthode subscribe() pour écouter les mises à jour venant d'autres appareils.
 */

import { initializeApp } from "firebase/app";
import {
  doc, getDoc, setDoc, deleteDoc, collection, getDocs,
  onSnapshot, initializeFirestore, persistentLocalCache, persistentMultipleTabManager,
} from "firebase/firestore";
import { getAuth, signInAnonymously, onAuthStateChanged } from "firebase/auth";

// Rempli automatiquement par le script de configuration (voir README-FIREBASE.md)
const firebaseConfig = window.__FIREBASE_CONFIG__ || null;

let db = null;
let authReadyPromise = Promise.resolve(null);
let configured = false;

if (firebaseConfig) {
  try {
    const app = initializeApp(firebaseConfig);
    db = initializeFirestore(app, {
      localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }),
    });
    const auth = getAuth(app);
    authReadyPromise = new Promise((resolve) => {
      onAuthStateChanged(auth, (user) => { if (user) resolve(user); });
      signInAnonymously(auth).catch((e) => {
        console.error("Connexion au serveur impossible :", e);
        resolve(null);
      });
    });
    configured = true;
  } catch (e) {
    console.error("Configuration Firebase invalide :", e);
  }
}

function refFor(key) {
  return doc(db, "kv", key);
}

const firestoreStorage = {
  async get(key) {
    if (!configured) return null;
    await authReadyPromise;
    try {
      const snap = await getDoc(refFor(key));
      if (!snap.exists()) return null;
      return { key, value: snap.data().value, shared: true };
    } catch (e) {
      console.error("Erreur de lecture :", e);
      return null;
    }
  },

  async set(key, value) {
    if (!configured) return null;
    await authReadyPromise;
    try {
      await setDoc(refFor(key), { value, updatedAt: Date.now() });
      return { key, value, shared: true };
    } catch (e) {
      console.error("Erreur d'écriture :", e);
      return null;
    }
  },

  async delete(key) {
    if (!configured) return null;
    await authReadyPromise;
    try {
      await deleteDoc(refFor(key));
      return { key, deleted: true, shared: true };
    } catch (e) {
      return null;
    }
  },

  async list(prefix = "") {
    if (!configured) return null;
    await authReadyPromise;
    try {
      const snaps = await getDocs(collection(db, "kv"));
      const keys = [];
      snaps.forEach((d) => { if (d.id.startsWith(prefix)) keys.push(d.id); });
      return { keys };
    } catch (e) {
      return null;
    }
  },

  // Écoute les changements en temps réel (venant de cet appareil OU d'un autre).
  // Retourne une fonction à appeler pour arrêter d'écouter.
  subscribe(key, callback) {
    if (!configured) return () => {};
    let unsub = () => {};
    let cancelled = false;
    authReadyPromise.then(() => {
      if (cancelled) return;
      unsub = onSnapshot(refFor(key), (snap) => {
        callback(snap.exists() ? snap.data().value : null);
      }, (err) => {
        console.error("Erreur de synchronisation :", err);
      });
    });
    return () => { cancelled = true; unsub(); };
  },

  isConfigured() {
    return configured;
  },
};

if (typeof window !== "undefined") {
  window.storage = firestoreStorage;
}

export default firestoreStorage;
