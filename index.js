const express = require('express')
const app = express()
const credentials = require('./credentials.json')
const { authorize, createSheet } = require('./utils')

const API_VERSION = 'v1'

app.use(express.static('public'));

app.post(`/api/${API_VERSION}/sheet`, () => {
  authorize(JSON.parse(credentials), (auth) => {
    createSheet(auth)
  });
})

app.listen(3000)
