const { readEvents } = require('./readEvents');

const path = "/Users/eliavcohen/Downloads/December-2020.xls";
const year = "2021";

const events = readEvents(path, year);

///// GCALENDAR
///////////////////////////////////
const fs = require("fs");
const readline = require("readline");
const { google } = require("googleapis");

// If modifying these scopes, delete token.json.
const SCOPES = ["https://www.googleapis.com/auth/calendar"];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = "token.json";

// Load client secrets from a local file.
fs.readFile("credentials.json", (err, content) => {
  if (err) return console.log("Error loading client secret file:", err);
  // Authorize a client with credentials, then call the Google Calendar API.
  authorize(JSON.parse(content), createEvents);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  const { client_secret, client_id, redirect_uris } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0]
  );

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getAccessToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getAccessToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: "offline",
    scope: SCOPES,
  });
  console.log("Authorize this app by visiting this url:", authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question("Enter the code from that page here: ", (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error("Error retrieving access token", err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return console.error(err);
        console.log("Token stored to", TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

const parseDates = (event) => {
	const timeZone = Intl.DateTimeFormat().resolvedOptions().timeZone;
	event.start = {
		dateTime: event.startDate.toISOString(),
		timeZone
	};
	event.end = {
		dateTime: event.endDate.toISOString(),
		timeZone
	};
	event.description = `The sleep quality was: ${event.note}/20. As I reported on the app "WORK LOG"`;
	delete event.note;
	delete event.startDate;
	delete event.endDate;
	event.summary = 'Sleep'
	return event;
}
function createEvents(auth) { 
	const calendar = google.calendar({ version: "v3", auth });
	createEventsAsync(calendar).then(() => { console.log('done') }).catch(() => { console.log('bummer') });
}

async function createEventsAsync(calendar) {
	for (eve of events.map(parseDates)) {
		try {
			await createEvent(eve, calendar);
		} catch {
			return;
		}
	}
}

function createEvent(eve, calendar) {
	return new Promise((resolve, reject) => {

		calendar.events.insert(
		{
			calendarId: "1rfo5a9nmpo6t9jscpr4ebdfis@group.calendar.google.com",
			resource: eve,
		},
		function (err, event) {
			if (err) {
			console.log("There was an error contacting the Calendar service: " + err);
			reject();
			}
			console.log("Event created: %s", event.htmlLink);
			resolve();
		}
		);
		
	});
}
