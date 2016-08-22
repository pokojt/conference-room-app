// code for creating events ******************************************************************** //

var outlook = require('node-outlook');
var userInfo = {email: 'taylorp@r-west.com'};

var calendarInfo = require('index.js').getValueFromCookie;

var token = getValueFromCookie('conference-room-app-token', request.headers.cookie);
var email = getValueFromCookie('conference-room-app-email', request.headers.cookie);

var newEvent = {
  "Subject": "New Test Event",
  "Body": {
    "ContentType": "HTML",
    "Content": "Trying to figure out what I'm doing"
  },
  "Start": {
    "DateTime": "2016-02-03T18:00:00",
    "TimeZone": "Eastern Standard Time"
  },
  "End": {
    "DateTime": "2016-02-03T19:00:00",
    "TimeZone": "Eastern Standard Time"
  },
  "Attendees": [
    {
      "EmailAddress": {
        "Address": "allieb@contoso.com",
        "Name": "Allie Bellew"
      },
      "Type": "Required"
    }
  ]
};

exports.newEvent = function(Subject,Body,Start,End,Atendees) {
  this.Subject = Subject;
  this.Body = Body;
  this.Start = Start;
  this.End = End;
  this.Atendees = Atendees;
}


outlook.base.setApiEndpoint('https://outlook.office.com/api/v2.0');



outlook.calendar.createEvent({token: token, event: newEvent, user: userInfo}),
  function(error, result) {
      if(error) {
        console.log('createEvent returned an error:' + error);
      }
      else if (result) {
        console.log(JSON.stringify(result, null, 2));
      }
  };

