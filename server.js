require('dotenv').config();
const express = require('express');
const path = require('path');

const app = express();
const HOST = process.env.HOST || '0.0.0.0';
const PORT = process.env.PORT || 3010;

// Serve static files from the public directory
app.use(express.static(path.join(__dirname, 'public')));

// Fallback for SPA (optional, remove if not needed)
// app.get('*', (req, res) => {
//   res.sendFile(path.join(__dirname, 'public', 'index.html'));
// });

app.listen(PORT, HOST, () => {
  let localUrl = `http://${HOST}:${PORT}`;
  let domainUrl = process.env.DOMAIN ? `https://${process.env.DOMAIN}/` : null;
  console.log(`Server running at: ${localUrl}`);
  if (domainUrl) {
    console.log(`Also accessible via domain: ${domainUrl}`);
    // NOTE: This server is running HTTP only. To serve HTTPS, you must set up SSL certificates and use the https module.
    // Example: https.createServer({ key, cert }, app).listen(443);
  }
});
