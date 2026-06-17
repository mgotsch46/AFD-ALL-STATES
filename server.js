const express = require('express');
const path = require('path');
const app = express();

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname)));

app.post('/generate', async (req, res) => {
  try {
    const generateDocs = require('./generateDocs');
    const zipBuffer = await generateDocs(req.body);
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="RenewEQ_Documents.zip"');
    res.send(zipBuffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('RenewEQ Doc Generator running on port ' + PORT));
