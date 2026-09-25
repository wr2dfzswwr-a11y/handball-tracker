// Standalone-Browser: den von der App erwarteten Speicher bereitstellen.
// Eine vorhandene Host-Implementierung bleibt erhalten.
if (!window.storage) {
  window.storage = {
    async get(key) {
      const value = localStorage.getItem(key);
      return value === null ? null : { key, value };
    },
    async set(key, value) {
      localStorage.setItem(key, value);
      return { key, value };
    },
    async delete(key) {
      localStorage.removeItem(key);
    },
  };
}
