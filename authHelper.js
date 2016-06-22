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
               "profile",
               // "https://outlook.office.com/mail.read",
               "https://outlook.office.com/calendars.readwrite" ];


function getAuthUrl() {
  var returnVal = oauth2.authCode.authorizeURL({
    redirect_uri: "http://localhost:8000/authorize",
    scope: scopes.join(" ")
  });
  console.log("Generated auth url: " + returnVal);
  return returnVal;
}

exports.getAuthUrl = getAuthUrl;


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

function getEmailFromIdToken(id_token) {
  // JWT is in three parts, separated by a '.'
  var token_parts = id_token.split('.');
  
  // Token content is in the second part, in urlsafe base64
  var encoded_token = new Buffer(token_parts[1].replace("-", "_").replace("+", "/"), 'base64');
  
  var decoded_token = encoded_token.toString();
  
  var jwt = JSON.parse(decoded_token);
  
  // Email is in the preferred_username field
  return jwt.preferred_username
}
exports.getEmailFromIdToken = getEmailFromIdToken;