//basic code to allow Node to run a web server listening on port 8000//

var http = require("http");
var url = require("url");
var path = require("path");
var fs = require('fs');

function start(route, handle) {
  function onRequest(request, response) {
    var pathName = url.parse(request.url).pathname;
    console.log("Request for " + pathName + " received.");

    route(handle, pathName, response, request);
  }

  var port = 8000;
  http.createServer(onRequest).listen(port);
  console.log("Server has started. Listening on port: " + port + "...");
}

exports.start = start;
