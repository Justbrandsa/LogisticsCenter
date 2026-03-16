const { routeRequest } = require("../serve");

module.exports = async function handler(request, response) {
  await routeRequest(request, response);
};
