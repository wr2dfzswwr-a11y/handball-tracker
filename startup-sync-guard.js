(() => {
  const DATA_KEY = "handball:v1";
  const originalFetch = window.fetch.bind(window);

  function isStartupSync() {
    const root = document.getElementById("root");
    if (!root) return false;
    const text = root.textContent || "";
    return text.includes("Lade Daten") || text.includes("Cloud-Stand wird geladen");
  }

  function isTeamDownload(input, init) {
    const method = String(
      (init && init.method) || (typeof Request !== "undefined" && input instanceof Request ? input.method : "GET")
    ).toUpperCase();
    const url = typeof input === "string" ? input : input && input.url;
    if (method !== "GET" || !url) return false;
    try {
      const parsed = new URL(url, window.location.href);
      return parsed.pathname.endsWith("/api/team");
    } catch {
      return false;
    }
  }

  function mergeCloudPayloadWithLocal(cloudPayload, localData) {
    if (!cloudPayload || !cloudPayload.team || !cloudPayload.team.id) return cloudPayload;
    if (!localData || !Array.isArray(localData.teams) || !Array.isArray(localData.games)) return cloudPayload;

    const teamId = cloudPayload.team.id;
    const localTeam = localData.teams.find((team) => team && team.id === teamId);
    const mergedGames = new Map();

    for (const game of Array.isArray(cloudPayload.games) ? cloudPayload.games : []) {
      if (game && game.id) mergedGames.set(game.id, game);
    }
    for (const game of localData.games) {
      if (game && game.teamId === teamId && game.id) mergedGames.set(game.id, game);
    }

    return {
      ...cloudPayload,
      team: localTeam || cloudPayload.team,
      games: Array.from(mergedGames.values()),
    };
  }

  async function protectStartupResponse(response) {
    if (!response || !response.ok || !window.storage || typeof window.storage.get !== "function") return response;

    try {
      const cloudResponse = await response.clone().json();
      if (!cloudResponse || !cloudResponse.data || !cloudResponse.data.team) return response;

      const stored = await window.storage.get(DATA_KEY);
      const localData = stored && stored.value ? JSON.parse(stored.value) : null;
      const protectedData = mergeCloudPayloadWithLocal(cloudResponse.data, localData);
      if (protectedData === cloudResponse.data) return response;

      const headers = new Headers(response.headers);
      headers.delete("content-length");
      headers.delete("content-encoding");
      headers.set("content-type", "application/json");

      return new Response(JSON.stringify({ ...cloudResponse, data: protectedData }), {
        status: response.status,
        statusText: response.statusText,
        headers,
      });
    } catch {
      return response;
    }
  }

  window.fetch = async function guardedFetch(input, init) {
    const protect = isStartupSync() && isTeamDownload(input, init);
    const response = await originalFetch(input, init);
    return protect ? protectStartupResponse(response) : response;
  };

  // Exposed only for lightweight regression tests/debugging; the app itself does not call it.
  window.__handballMergeStartupPayload = mergeCloudPayloadWithLocal;
})();
