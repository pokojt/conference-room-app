// defines a function to generate the login URL

var credentials = {
  clientID: "7dabb801-6f5a-403e-9083-f8b31437691f", //YOUR APP ID HERE
  clientSecret: "JRWe2tzgivBGsWbvcbg4Ze3", //YOUR APP PASSWORD
  site: "https://login.microsoftonline.com/common",
  authorizationPath: "/oauth2/v2.0/authorize",
  tokenPath: "/oauth2/v2.0/token"
}
var oauth2 = require("simple-oauth2")(credentials);

var redirectUri = "http://localhost:8000/authorize";

// The scopes the app requires
var scopes = [ "openid",
               "offline_access",
               "profile",
               "https://outlook.office.com/mail.read",
               "https://outlook.office.com/calendars.readwrite" ];


function getAuthUrl() {
  var returnVal = oauth2.authCode.authorizeURL({
    redirect_uri: redirectUri,
    scope: scopes.join(" ")
  });
  console.log("Generated auth url: " + returnVal);
  return returnVal;
}

exports.getAuthUrl = getAuthUrl;


// Takes authorization code and creates the access token
function getTokenFromCode(auth_code, callback, response) {
  var token;
  oauth2.authCode.getToken({
    code: auth_code,
    redirect_uri: redirectUri,
    scope: scopes.join(" ")
    }, function (error, result) {
      if (error) {
        console.log("Access token error: ", error.message);
        callback(response, error, null);
      }
      else {
        token = oauth2.accessToken.create(result);
        console.log("Token created: ", token.token);
        callback(response, null, token);
      }
    });
}

exports.getTokenFromCode = getTokenFromCode;



function refreshAccessToken(refreshToken, callback) {
  var tokenObj = oauth2.accessToken.create({refresh_token: refreshToken});
  tokenObj.refresh(callback);
}

exports.refreshAccessToken = refreshAccessToken;