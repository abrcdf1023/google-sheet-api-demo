const express = require('express')
const app = express()
const credentials = require('./credentials.json')
const { authorize, createSheet } = require('./utils')

const API_VERSION = 'v1'

app.use(express.static('public'));
app.use(express.json());

app.post(`/api/${API_VERSION}/sheet`, async (req, res) => {
  const auth = await authorize(credentials);
  const sheetUrl = await createSheet(auth, {
    properties: {
      title: req.body.title,
    }
  })
  res.send({
    sheetUrl
  });
})

app.listen(3000)
