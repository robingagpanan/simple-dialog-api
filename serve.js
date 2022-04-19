'use strict';

var fs = require('fs');
var https = require('https');
var homedir = require('os').homedir();
var express = require('express');

var app = express();

app.use(function(req, res, next) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  next();
});

app.use(express.static('src'))

var port = Number(process.env.PORT || 8000);

https.createServer({
  key: fs.readFileSync(homedir + '/.office-addin-dev-certs/localhost.key'),
  cert: fs.readFileSync(homedir + '/.office-addin-dev-certs/localhost.crt')
}, app)
.listen(port, function() {
  console.log('Listening on ' + port);
});
