require('dotenv').config();
const express = require('express');
const path = require('path');

const app = express();
const HOST = process.env.HOST || 'localhost';
const PORT = process.env.PORT || 3010;

// Serve static files from the public directory
app.use(express.static(path.join(__dirname, 'public')));

// Fallback for SPA (optional, remove if not needed)
// app.get('*', (req, res) => {
//   res.sendFile(path.join(__dirname, 'public', 'index.html'));
// });

app.listen(PORT, HOST, () => {
  console.log(`Server running on http://${HOST}:${PORT}`);
});
