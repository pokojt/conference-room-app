//imports js files and opens working app at http://localhost:8000 after running $node index.js

var server = require("./server");
var router = require("./router");
var authHelper = require("./authHelper");

var handle = {};
handle["/"] = home;
handle["/authorize"] = authorize; //tells our router that when a GET request comes in for /authorize, invoke the authorize function

server.start(router.route, handle);

function home(response, request) {
  console.log("Request handler 'home' was called.");
  response.writeHead(200, {"Content-Type": "text/html"});
  response.write('<p>Please <a href="' + authHelper.getAuthUrl() + '">sign in</a> with your Office 365 or Outlook.com account.</p>');
  response.end();
}


var url = require("url");


function authorize(response, request) {
  console.log("Request handler 'authorize' was called.");
  
  // The authorization code is passed as a query parameter
  var url_parts = url.parse(request.url, true);
  var code = url_parts.query.code;
  console.log("Code: " + code);
  var token = authHelper.getTokenFromCode(code, tokenReceived, response);
}

function tokenReceived(response, error, token) {
  if (error) {
    console.log("Access token error: ", error.message);
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<p>ERROR: ' + error + '</p>');
    response.end();
  }
  else {
    var cookies = ['conference-room-app-token=' + token.token.access_token + ';Max-Age=3600',
                   'conference-room-app-email=' + authHelper.getEmailFromIdToken(token.token.id_token) + ';Max-Age=3600'];
    response.setHeader('Set-Cookie', cookies);
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<p>Access token saved in cookie.</p>');
    response.end();
  }
}
