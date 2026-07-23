/* =====================================================================
 * store.js — persistence adapter
 * ---------------------------------------------------------------------
 * POC backend: localStorage. The interface is async and schema-clean so a
 * future swap to IndexedDB (needed at ~100+ properties; localStorage caps
 * around 5 MB) or a real database/API is a one-file change:
 * implement the same four methods and replace `createStore()`.
 *
 * FUTURE AUTH HOOK: when a CRM-backed auth layer exists, property access
 * checks belong here — filter listProperties()/getProperty() by the
 * signed-in user's property permissions before returning.
 * ===================================================================== */

(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.Store = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var INDEX_KEY = "rd:index";
  var PROP_PREFIX = "rd:prop:";

  function createLocalStorageStore() {
    function readIndex() {
      try { return JSON.parse(localStorage.getItem(INDEX_KEY)) || []; }
      catch (e) { return []; }
    }
    function writeIndex(ids) {
      localStorage.setItem(INDEX_KEY, JSON.stringify(ids));
    }
    return {
      /** @returns Promise<propertySummary[]> (full objects; callers may slim) */
      listProperties: function () {
        var out = [];
        readIndex().forEach(function (id) {
          try {
            var p = JSON.parse(localStorage.getItem(PROP_PREFIX + id));
            if (p) out.push(p);
          } catch (e) { /* skip corrupt entry */ }
        });
        return Promise.resolve(out);
      },
      getProperty: function (id) {
        try {
          return Promise.resolve(JSON.parse(localStorage.getItem(PROP_PREFIX + id)));
        } catch (e) { return Promise.resolve(null); }
      },
      saveProperty: function (property) {
        try {
          localStorage.setItem(PROP_PREFIX + property.id, JSON.stringify(property));
        } catch (e) {
          // Quota exceeded — surface a friendly error (POC limitation)
          return Promise.reject(new Error(
            "Browser storage is full (localStorage ~5 MB cap). Remove a property " +
            "or migrate to the IndexedDB adapter. Original error: " + e.message));
        }
        var ids = readIndex();
        if (ids.indexOf(property.id) === -1) { ids.push(property.id); writeIndex(ids); }
        return Promise.resolve(property);
      },
      deleteProperty: function (id) {
        localStorage.removeItem(PROP_PREFIX + id);
        writeIndex(readIndex().filter(function (x) { return x !== id; }));
        return Promise.resolve();
      }
    };
  }

  return { create: createLocalStorageStore };
});
