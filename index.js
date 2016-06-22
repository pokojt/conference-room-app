//imports js files and opens working app at http://localhost:8000 after running $node index.js

var server = require("./server");
var router = require("./router");
var authHelper = require("./authHelper");
var outlook = require("node-outlook");

var handle = {};
handle["/"] = home;
handle["/authorize"] = authorize; //tells our router that when a GET request comes in for /authorize, invoke the authorize function
handle["/calendar"] = calendar;

server.start(router.route, handle);


//Home screen - sign in to Exchange account
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

function getValueFromCookie(valueName, cookie) {
  if (cookie.indexOf(valueName) !== -1) {
    var start = cookie.indexOf(valueName) + valueName.length + 1;
    var end = cookie.indexOf(';', start);
    end = end === -1 ? cookie.length : end;
    return cookie.substring(start, end);
  }
}

function calendar(response, request) {
  var token = getValueFromCookie('node-tutorial-token', request.headers.cookie);
  console.log("Token found in cookie: ", token);
  var email = getValueFromCookie('node-tutorial-email', request.headers.cookie);
  console.log("Email found in cookie: ", email);
  if (token) {
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<div><h1>Your Calendar</h1></div>');
    
    var queryParams = {
      '$select': 'Subject,Start,End',
      '$orderby': 'Start/DateTime desc',
      '$top': 10
    };
    
    // Set the API endpoint to use the v2.0 endpoint
    outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');
    // Set the anchor mailbox to the user's SMTP address
    outlook.base.setAnchorMailbox(email);
    //Set the preferred time zone.
    // The API will return event date/times in this time zone.
    outlook.base.setPreferredTimeZone('America/Los_Angeles');

    outlook.calendar.getEvents({token: token, odataParams: queryParams},
      function(error, result){
        if (error) {
          console.log('getEvents returned an error: ' + error);
          response.write("<p>ERROR: " + error + "</p>");
          response.end();
        }
        else if (result) {
          console.log('getEvents returned ' + result.value.length + ' events.');
          response.write('<table><tr><th>Subject</th><th>Start</th><th>End</th></tr>');
          result.value.forEach(function(event) {
            console.log('  Subject: ' + event.Subject);
            response.write('<tr><td>' + event.Subject + 
              '</td><td>' + event.Start.DateTime.toString() +
              '</td><td>' + event.End.DateTime.toString() + '</td></tr>');
          });
          
          response.write('</table>');
          response.end();
        }
      });
	}
  else {
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<p> No token found in cookie!</p>');
    response.end();
  }
}

