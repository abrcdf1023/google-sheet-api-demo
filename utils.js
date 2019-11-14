const readline = require('readline');
const fs = require('fs');
const {
  google
} = require('googleapis');

// If modifying these scopes, delete token.json.
const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive.metadata',
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/drive'
];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

/**
 * Create an OAuth2 client with the given credentials.
 * @param {Object} credentials The authorization client credentials.
 */
module.exports.authorize = function authorize(credentials) {
  const {
    client_secret,
    client_id,
    redirect_uris
  } = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
    client_id, client_secret, redirect_uris[0]);
  return new Promise((resolve, reject) => {
    try {
      const token = fs.readFileSync(TOKEN_PATH);
      oAuth2Client.setCredentials(JSON.parse(token));
      resolve(oAuth2Client);
    } catch (err) {
      getNewToken(oAuth2Client, resolve);
    }
  })
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given resolve with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsResolve} resolve The resolve for the authorized client.
 */
module.exports.getNewToken = function getNewToken(oAuth2Client, resolve) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error('Error while trying to retrieve access token', err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      resolve(oAuth2Client);
    });
  });
}

/**
 * Get sheet data
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
module.exports.getSheet = function getSheet(auth, params) {
  const sheets = google.sheets({
    version: 'v4',
    auth
  });
  return new Promise((resolve, reject) => {
    sheets.spreadsheets.values.get(params, (err, res) => {
      if (err) return console.log('The API returned an error: ' + err);
      resolve(res.data);
    });
  })
}

/**
 * Create a new sheet.
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
module.exports.createSheet = function createSheet(auth, resource) {
  return google.drive({
    version: 'v3',
    auth
  }).files.copy({
    fileId: '1eyDaop4XjlFHGIzD7sUJQe1kU7MVwDouAHtVeI5PFvQ',
    fields: 'id, webViewLink',
    requestBody: {
      ...resource,
      parents: ['1PbG3D7gtJ1SJkE136qA1emsIE5vKYBeT']
    }
  })
  .then(res => res.data)
  .catch((err) => {
    console.log('The API returned an error: ' + err);
  })
}