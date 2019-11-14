const express = require('express')
const app = express()
const credentials = require('./credentials.json')
const { authorize, createSheet, getSheet } = require('./utils')

const API_VERSION = 'v1'

app.use(express.static('public'));
app.use(express.json());

app.get(`/api/${API_VERSION}/sheet`, async (req, res) => {
  req.query.spreadsheetId
  const auth = await authorize(credentials);
  const data = await getSheet(auth, {
    spreadsheetId: req.query.spreadsheetId,
    // range: 'Class Data!A2:E',
  })
  res.send(data);
})

app.post(`/api/${API_VERSION}/sheet`, async (req, res) => {
  const auth = await authorize(credentials);
  const data = await createSheet(auth, {
    name: req.body.name,
  })
  res.send({
    spreadsheetUrl: data.webViewLink,
    spreadsheetId: data.id,
  });
})

app.listen(3000)
