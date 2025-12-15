require('dotenv').config();
const express = require('express');
const path = require('path');
const fs = require('fs');
const https = require('https');

const app = express();
const HOST = process.env.HOST || 'localhost';
const PORT = process.env.PORT || 443;

// SSL certificate paths (update if your certs are in a different location)
const sslOptions = {
  key: fs.readFileSync('/etc/letsencrypt/live/utility.sedl.in/privkey.pem'),
  cert: fs.readFileSync('/etc/letsencrypt/live/utility.sedl.in/fullchain.pem')
};

// Serve static files from the public directory
app.use(express.static(path.join(__dirname, 'public')));

// Fallback for SPA (optional, remove if not needed)
// app.get('*', (req, res) => {
//   res.sendFile(path.join(__dirname, 'public', 'index.html'));
// });

https.createServer(sslOptions, app).listen(PORT, HOST, () => {
  console.log(`HTTPS Server running on https://${HOST}:${PORT}`);
});
