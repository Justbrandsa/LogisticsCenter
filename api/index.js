const { routeRequest } = require("../serve");

module.exports = async function handler(request, response) {
  const requestUrl = new URL(request.url || "/", "https://route-ledger.local");
  const route = String(requestUrl.searchParams.get("route") || "").replace(/^\/+/, "");

  if (route) {
    requestUrl.searchParams.delete("route");
    const query = requestUrl.searchParams.toString();
    request.url = `/api/${route}${query ? `?${query}` : ""}`;
  }

  await routeRequest(request, response);
};
