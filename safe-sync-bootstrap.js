(() => {
  const DATA_KEY = "handball:v1";
  const CLOUD_KEY = "handball:cloud-v1";

  if (typeof window.fetch !== "function") return;

  const originalFetch = window.fetch.bind(window);
  let localData = null;
  let startupDownloadsRemaining = 0;
  let storagePatched = false;

  const deepClone = (value) => JSON.parse(JSON.stringify(value));

  function patchStorage(storage) {
    if (storagePatched || !storage || typeof storage.get !== "function") return;
    storagePatched = true;

    const originalGet = storage.get.bind(storage);

    storage.get = async function patchedStorageGet(key, ...args) {
      const result = await originalGet(key, ...args);
      try {
        if (key === DATA_KEY) {
          localData = result && result.value
            ? JSON.parse(result.value)
            : { teams: [], games: [] };
        } else if (key === CLOUD_KEY) {
          const cloud = result && result.value ? JSON.parse(result.value) : null;
          startupDownloadsRemaining =
            navigator.onLine !== false &&
            cloud &&
            cloud.apiBase &&
            Array.isArray(cloud.links)
              ? cloud.links.length
              : 0;
        }
      } catch {
        if (key === DATA_KEY) localData = { teams: [], games: [] };
        if (key === CLOUD_KEY) startupDownloadsRemaining = 0;
      }
      return result;
    };
  }

  // Patch immediately if storage already exists. If the production bundle
  // installs window.storage itself, intercept that assignment before app startup.
  if (window.storage) {
    patchStorage(window.storage);
  } else {
    try {
      let assignedStorage;
      Object.defineProperty(window, "storage", {
        configurable: true,
        enumerable: true,
        get() {
          return assignedStorage;
        },
        set(value) {
          assignedStorage = value;
          patchStorage(value);
        },
      });
    } catch {
      // Last-resort fallback for environments where the property cannot be hooked.
      const timer = setInterval(() => {
        if (window.storage) {
          clearInterval(timer);
          patchStorage(window.storage);
        }
      }, 0);
      setTimeout(() => clearInterval(timer), 5000);
    }
  }

  window.fetch = async function safeStartupFetch(input, init) {
    const url = typeof input === "string" ? input : (input && input.url) || "";
    const requestMethod =
      input && typeof input !== "string" && input.method ? input.method : "GET";
    const method = String((init && init.method) || requestMethod || "GET").toUpperCase();
    const isStartupTeamDownload =
      startupDownloadsRemaining > 0 &&
      method === "GET" &&
      /\/api\/team(?:\?|$)/.test(url);

    // Consume the slot before the request so a network failure cannot cause a
    // later manual download to be mistaken for startup synchronization.
    if (isStartupTeamDownload) startupDownloadsRemaining -= 1;

    const response = await originalFetch(input, init);
    if (!isStartupTeamDownload || !response.ok || !localData) return response;

    try {
      const body = await response.clone().json();
      const remote = body && body.data;
      if (!remote || !remote.team || !remote.team.id || !Array.isArray(remote.games)) {
        return response;
      }

      const teamId = remote.team.id;
      const localTeam = (localData.teams || []).find((team) => team.id === teamId);
      const localGames = (localData.games || []).filter((game) => game.teamId === teamId);

      // Startup sync is additive/local-first:
      // - existing local team data wins
      // - existing local games win
      // - cloud-only games are added
      // A missing cloud game is therefore never interpreted as a deletion.
      if (localTeam) remote.team = deepClone(localTeam);

      const mergedGames = new Map(
        (remote.games || []).map((game) => [game.id, game])
      );
      for (const game of localGames) {
        mergedGames.set(game.id, deepClone(game));
      }
      remote.games = Array.from(mergedGames.values());
      body.data = remote;

      const headers = new Headers(response.headers);
      headers.delete("content-length");
      headers.delete("content-encoding");

      return new Response(JSON.stringify(body), {
        status: response.status,
        statusText: response.statusText,
        headers,
      });
    } catch {
      // If the safety merge itself fails, do not break cloud access.
      return response;
    }
  };
})();
