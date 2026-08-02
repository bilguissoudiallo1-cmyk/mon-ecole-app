/*
 * Stockage local, 100% hors-ligne (remplace le stockage cloud).
 * Même interface (get/set/delete/list) que window.storage utilisée dans le
 * reste de l'application, mais tout est enregistré directement sur
 * l'appareil (localStorage). Aucune connexion internet n'est nécessaire.
 * Le paramètre "shared" est ignoré ici : sans serveur, il n'y a pas de
 * notion de données partagées entre appareils — chaque appareil a sa
 * propre base. Utilisez "Sauvegarde" dans l'application pour transférer
 * les données d'un appareil à un autre.
 */

const PREFIX = "ecole_kv::";

const localStorageAPI = {
  async get(key) {
    try {
      const raw = window.localStorage.getItem(PREFIX + key);
      if (raw === null) return null;
      return { key, value: raw, shared: false };
    } catch (e) {
      return null;
    }
  },

  async set(key, value) {
    try {
      window.localStorage.setItem(PREFIX + key, value);
      return { key, value, shared: false };
    } catch (e) {
      // Espace de stockage plein ou navigation privée : on prévient l'utilisateur
      console.error("Erreur de stockage local :", e);
      return null;
    }
  },

  async delete(key) {
    try {
      window.localStorage.removeItem(PREFIX + key);
      return { key, deleted: true, shared: false };
    } catch (e) {
      return null;
    }
  },

  async list(prefix = "") {
    try {
      const keys = Object.keys(window.localStorage)
        .filter((k) => k.startsWith(PREFIX + prefix))
        .map((k) => k.slice(PREFIX.length));
      return { keys };
    } catch (e) {
      return null;
    }
  },
};

// Exposé globalement pour que le reste de l'application (App.jsx) fonctionne
// sans aucune modification : il appelle simplement window.storage.get/set/...
if (typeof window !== "undefined") {
  window.storage = localStorageAPI;
}

export default localStorageAPI;
