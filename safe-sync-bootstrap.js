(() => {
  const DATA_KEY = "handball:v1";
  const CLOUD_KEY = "handball:cloud-v1";
  const storage = window.storage;

  if (!storage || typeof storage.get !== "function" || typeof window.fetch !== "function") return;

  const originalGet = storage.get.bind(storage);
  const originalFetch = window.fetch.bind(window);
  let localData = null;
  let startupDownloadsRemaining = 0;

  storage.get = async function patchedStorageGet(key, ...args) {
    const result = await originalGet(key, ...args);
    try {
      if (key === DATA_KEY) {
        localData = result && result.value ? JSON.parse(result.value) : { teams: [], games: [] };
      } else if (key === CLOUD_KEY) {
        const cloud = result && result.value ? JSON.parse(result.value) : null;
        startupDownloadsRemaining = Array.isArray(cloud && cloud.links) ? cloud.links.length : 0;
      }
    } catch {
      if (key === DATA_KEY) localData = { teams: [], games: [] };
      if (key === CLOUD_KEY) startupDownloadsRemaining = 0;
    }
    return result;
  };

  window.fetch = async function safeStartupFetch(input, init) {
    const url = typeof input === "string" ? input : (input && input.url) || "";
    const requestMethod = input && typeof input !== "string" && input.method ? input.method : "GET";
    const method = String((init && init.method) || requestMethod || "GET").toUpperCase();
    const isStartupTeamDownload =
      startupDownloadsRemaining > 0 &&
      method === "GET" &&
      /\/api\/team(?:\?|$)/.test(url);

    const response = await originalFetch(input, init);
    if (!isStartupTeamDownload) return response;

    startupDownloadsRemaining -= 1;
    if (!response.ok || !localData) return response;

    try {
      const body = await response.clone().json();
      const remote = body && body.data;
      if (!remote || !remote.team || !remote.team.id || !Array.isArray(remote.games)) return response;

      const teamId = remote.team.id;
      const localTeam = (localData.teams || []).find((team) => team.id === teamId);
      const localGames = (localData.games || []).filter((game) => game.teamId === teamId);

      // Automatic startup sync is additive: cloud may add missing data, but it
      // must never overwrite data that already exists on this device.
      if (localTeam) remote.team = JSON.parse(JSON.stringify(localTeam));
      const mergedGames = new Map((remote.games || []).map((game) => [game.id, game]));
      for (const game of localGames) mergedGames.set(game.id, JSON.parse(JSON.stringify(game)));
      remote.games = Array.from(mergedGames.values());
      body.data = remote;

      return new Response(JSON.stringify(body), {
        status: response.status,
        statusText: response.statusText,
        headers: response.headers,
      });
    } catch {
      // If the safety merge itself fails, keep the original server response.
      return response;
    }
  };
})();
