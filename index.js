//imports js files and opens working app at http://localhost:8000 after running $node index.js

var server = require("./server");
var router = require("./router");
var authHelper = require("./authHelper");
var outlook = require("node-outlook");
var moment = require("moment");

var handle = {};
handle["/"] = home;
handle["/authorize"] = authorize; //tells our router that when a GET request comes in for /authorize, invoke the authorize function
// handle['/mail'] = mail;
handle["/calendar"] = calendar;

server.start(router.route, handle);


// This app links users to their Azure login page, where they can login with their Office 365 or Outlook account and grant access to the app. It then redirects users back to the app and displays a list of the day's calendar events.

//Home screen - sign in to Exchange account
function home(response, request) {
  console.log("Request handler \'home'\ was called.");
  response.writeHead(200, {"Content-Type": "text/html"});
  response.write('<p>Please <a href="' + authHelper.getAuthUrl() + '">sign in</a> with your Office 365 or Outlook.com account.</p>');
  response.end();
}


var url = require("url");

function authorize(response, request) {
  console.log("Request handler \'authorize\' was called.");
  
  // The authorization code is passed as a query parameter
  var url_parts = url.parse(request.url, true);
  var code = url_parts.query.code;
  console.log("Code: " + code);
  authHelper.getTokenFromCode(code, tokenReceived, response);
}

// stores token and email in a session cookie 
function tokenReceived(response, error, token) {
  if (error) {
    console.log("Access token error: ", error.message);
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<p>ERROR: ' + error + '</p>');
    response.end();
  }
  else {
    getUserEmail(token.token.access_token, function(error, email) {
      if (error) {
        console.log("getUserEMail returned an error: " + error);
        response.writeHead(200, {"Content-Type": "text/html"});
        response.write('<p>ERROR: ' + error + '</p>');
        response.end();
      } else if (email) {
        var cookies = ['conference-room-app-token=' + token.token.access_token + ';Max-Age=4000',
                   'conference-room-app-refresh-token=' + token.token.refresh_token + ';Max-Age=4000',
                   'conference-room-app-token-expires=' + token.token.expires_at.getTime() + ';Max-Age=4000',
                   'conference-room-app-email=' + email + ';Max-Age=4000'];
        response.setHeader('Set-Cookie', cookies);
        response.writeHead(302, {"Location": "http://localhost:8000/calendar"});
        response.end();
      }
    });
  }
}


//need users email to make requests to API

function getUserEmail(token, callback) {
  // Set the API endpoint to use the v2.0 endpoint
  outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');

  // Set up oData parameters
  var queryParams = {
    '$select': 'DisplayName, EmailAddress',
  };

  outlook.base.getUser({token: token, odataParams: queryParams}, function(error, user){
    if (error) {
      callback(error, null);
    } else {
      callback(null, user.EmailAddress);
    }
  });
}




// helper function to read cookie values
function getValueFromCookie(valueName, cookie) {
  if (cookie.indexOf(valueName) !== -1) {
    var start = cookie.indexOf(valueName) + valueName.length + 1;
    var end = cookie.indexOf(';', start);
    end = end === -1 ? cookie.length : end;
    return cookie.substring(start, end);
  }
}


//  helper function that retrieves the cached token, checks if it's expired, and refreshes it if so.
function getAccessToken(request, response, callback) {
  var expiration = new Date(parseFloat(getValueFromCookie('conference-room-app-token-expires', request.headers.cookie)));

  if (Date.compare(expiration, new Date()) === -1) {
    // refresh token
    console.log('TOKEN EXPIRED, REFRESHING');
    var refresh_token = getValueFromCookie('conference-room-app-refresh-token', request.headers.cookie);
    authHelper.refreshAccessToken(refresh_token, function(error, newToken){
      if (error) {
        callback(error, null);
      } else if (newToken) {
        var cookies = ['conference-room-app-token=' + newToken.token.access_token + ';Max-Age=4000',
                       'conference-room-app-refresh-token=' + newToken.token.refresh_token + ';Max-Age=4000',
                       'conference-room-app-token-expires=' + newToken.token.expires_at.getTime() + ';Max-Age=4000'];
        response.setHeader('Set-Cookie', cookies);
        callback(null, newToken.token.access_token);
      }
    });
  } 
  else {
    // Return cached token
    var access_token = getValueFromCookie('conference-room-app-token', request.headers.cookie);
    callback(null, access_token);
  }
}

// reads token from cookie and makes call to Calendar API
function calendar(response, request) {
    var token = getValueFromCookie('conference-room-app-token', request.headers.cookie);
    console.log('Token found in cookie: ', token);
    var email = getValueFromCookie('conference-room-app-email', request.headers.cookie);
    console.log('Email found in cookie: ', email);
  if (token) {
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write("<div><h1>Today's Events</h1></div>");
    
    var queryParams = {
      '$select': 'Subject,Start,End,Attendees,BodyPreview',
      '$orderby': 'Start/DateTime desc',
      '$top': 30
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
          var today = moment().format("YYYY-MM-DD").toString();
          console.log(result);
          console.log('getEvents returned ' + result.value.length + ' events.');
          console.log("Today is " + today);
          response.write('<div class="eventsWrapper"><table><tr><th>Subject</th><th>Start</th><th>End</th><th>Attendees</th><th>Summary</th></tr>');
          
          result.value.forEach(function(event) {
            if (event.Start.DateTime.includes(today)) {
              console.log('  Subject: ' + event.Subject);
              response.write('<tr><td>' + event.Subject + 
                '</td><td>' + event.Start.DateTime.toString() +
                '</td><td>' + event.End.DateTime.toString() + '</td><td>' + event.Attendees[0][1] + '</td><td>' + event.BodyPreview + '</td></tr>');
            }
          });
          response.write('</table></div>');
          response.end();
        }
        // code here for new event? outlook.calendar.createEvent?
      });
	}
  else {
    response.writeHead(200, {"Content-Type": "text/html"});
    response.write('<p> No token found in cookie!</p>');
    response.end();
  }
}


