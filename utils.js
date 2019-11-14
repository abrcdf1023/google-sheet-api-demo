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

function moveFileToPublicFolder(auth, spreadsheetId) {
  const drive = google.drive({
    version: 'v3',
    auth
  })
  return new Promise((resolve, reject) => {
    drive.files.update({
      fileId: spreadsheetId,
      addParents: '1PbG3D7gtJ1SJkE136qA1emsIE5vKYBeT',
      fields: 'id, parents'
    }, function (err, res) {
      if (err) {
        reject(err)
      } else {
        resolve(res)
      }
    });
  })
}

function updateDefaultSheetTemplateLayout(auth, spreadsheetId) {
  const sheets = google.sheets({
    version: 'v4',
    auth
  });
  return sheets.spreadsheets.batchUpdate({
    spreadsheetId,
    requestBody: {
      requests: [{
        mergeCells: {
          mergeType: 'MERGE_ALL',
          range: {
            sheetId: 0,
            startRowIndex: 0,
            endRowIndex: 1,
            startColumnIndex: 0,
            endColumnIndex: 2
          },
        }
      }, {
        mergeCells: {
          mergeType: 'MERGE_COLUMNS',
          range: {
            sheetId: 0,
            startRowIndex: 1,
            endRowIndex: 8,
            startColumnIndex: 0,
            endColumnIndex: 1
          },
        }
      }, {
        mergeCells: {
          mergeType: 'MERGE_COLUMNS',
          range: {
            sheetId: 0,
            startRowIndex: 8,
            endRowIndex: 22,
            startColumnIndex: 0,
            endColumnIndex: 1
          },
        }
      }, {
        repeatCell: {
          cell: {
            userEnteredFormat: {
              horizontalAlignment: "CENTER",
              verticalAlignment: "MIDDLE",
              numberFormat: {
                type: "NUMBER",
                pattern: "#,###"
              }
            }
          },
          range: {
            sheetId: 0,
            startRowIndex: 0,
            endRowIndex: 22,
            startColumnIndex: 0,
            endColumnIndex: 9,
          },
          fields: "userEnteredFormat"
        }
      }]
    }
  })
}

function updateDefaultSheetTemplate(auth, spreadsheetId) {
  const sheets = google.sheets({
    version: 'v4',
    auth
  });
  return new Promise((resolve, reject) => {
    sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: [{
          majorDimension: "ROWS",
          range: "A1:A1",
          values: [
            ["項目/時間"],
          ],
        }, {
          majorDimension: "ROWS",
          range: "A2:A2",
          values: [
            ["收入"],
          ],
        }, {
          majorDimension: "ROWS",
          range: "A9:A9",
          values: [
            ["支出"],
          ],
        }, {
          majorDimension: "ROWS",
          range: "C1:I1",
          values: [
            ["2019-01", "2019-02", "2019-03", "2019-04", "2019-05", "2019-06", "合計"],
          ],
        }, {
          majorDimension: "ROWS",
          range: "C8:I8",
          values: [
            ["=SUM(C2:C7)", "=SUM(D2:D7)", "=SUM(E2:E7)", "=SUM(F2:F7)", "=SUM(G2:G7)", "=SUM(H2:H7)", "=SUM(I2:I7)"],
          ],
        }, {
          majorDimension: "ROWS",
          range: "C22:I22",
          values: [
            ["=SUM(C9:C21)", "=SUM(D9:D21)", "=SUM(E9:E21)", "=SUM(F9:F21)", "=SUM(G9:G21)", "=SUM(H9:H21)", "=SUM(I9:I21)"],
          ],
        }, {
          majorDimension: "ROWS",
          range: "B2:B8",
          values: [
            ["工資"],
            ["非工資"],
            ["家人"],
            ["補助或津貼"],
            ["借款"],
            ["其他"],
            ["小計"],
          ],
        }, {
          majorDimension: "ROWS",
          range: "I2:I7",
          values: [
            ["=SUM(C2:H2)"],
            ["=SUM(C3:H3)"],
            ["=SUM(C4:H4)"],
            ["=SUM(C5:H5)"],
            ["=SUM(C6:H6)"],
            ["=SUM(C7:H7)"],
          ],
        }, {
          majorDimension: "ROWS",
          range: "B9:B22",
          values: [
            ["食"],
            ["衣"],
            ["住"],
            ["行"],
            ["育"],
            ["樂"],
            ["電信"],
            ["還款"],
            ["儲蓄"],
            ["醫療"],
            ["保險"],
            ["孝養"],
            ["其他"],
            ["小計"],
          ],
        }, {
          majorDimension: "ROWS",
          range: "I9:I22",
          values: [
            ["=SUM(C9:H9)"],
            ["=SUM(C10:H10)"],
            ["=SUM(C11:H11)"],
            ["=SUM(C12:H12)"],
            ["=SUM(C13:H13)"],
            ["=SUM(C14:H14)"],
            ["=SUM(C15:H15)"],
            ["=SUM(C16:H16)"],
            ["=SUM(C17:H17)"],
            ["=SUM(C18:H18)"],
            ["=SUM(C19:H19)"],
            ["=SUM(C20:H20)"],
            ["=SUM(C21:H21)"],
            ["=SUM(C22:H22)"],
          ],
        }
      ]
      }
    }, (err, result) => {
      if (err) {
        reject(err);
      } else {
        resolve(result.updatedCells)
      }
    })
  })
}
/**
 * Create a new sheet.
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
module.exports.createSheet = function createSheet(auth, resource) {
  const sheets = google.sheets({
    version: 'v4',
    auth
  });
  return sheets.spreadsheets.create({
    requestBody: {
      properties: {
        ...resource,
      }
    },
  }).then(async (res) => {
    await moveFileToPublicFolder(auth, res.data.spreadsheetId)
    await updateDefaultSheetTemplateLayout(auth, res.data.spreadsheetId)
    await updateDefaultSheetTemplate(auth, res.data.spreadsheetId)
    return res.data
  }).catch((err) => {
    console.log('The API returned an error: ' + err);
  })
}