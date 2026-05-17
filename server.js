const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');
const multer = require('multer');
const FormData = require('form-data');
const { parseAndBuildDocx } = require('./generateDocx');
const { Resend } = require('resend');

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

app.use(cors());
app.use(express.json({ limit: '50mb' }));

// Route transcription audio via Whisper
app.post('/transcribe', upload.single('audio'), async (req, res) => {
  try {
    const formData = new FormData();
    formData.append('file', req.file.buffer, {
      filename: 'audio.webm',
      contentType: req.file.mimetype
    });
    formData.append('model', 'whisper-1');
    formData.append('language', 'fr');

    const response = await fetch('https://api.openai.com/v1/audio/transcriptions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${process.env.OPENAI_API_KEY}`,
        ...formData.getHeaders()
      },
      body: formData
    });

    const data = await response.json();
    res.json(data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Route génération CR via Claude
app.post('/generate', async (req, res) => {
  try {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5',
        max_tokens: 1500,
        messages: req.body.messages
      })
    });

    const data = await response.json();
    res.json(data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get('/', (req, res) => res.send('CR Réunions API — OK'));

// Route génération DOCX
app.post('/docx', async (req, res) => {
  try {
    const { crTexte, meta } = req.body;
    const buffer = await parseAndBuildDocx(crTexte, meta);
    const filename = `CR_${(meta.titre || 'reunion').replace(/\s+/g, '_')}_${(meta.date || '').replace(/\//g, '-')}.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Route envoi e-mail avec DOCX en pièce jointe
app.post('/send-email', async (req, res) => {
  try {
    const { crTexte, meta, destinataire } = req.body;
    const buffer = await parseAndBuildDocx(crTexte, meta);
    const filename = `CR_${(meta.titre || 'reunion').replace(/\s+/g, '_')}_${(meta.date || '').replace(/\s+/g, '_')}.docx`;

    const resend = new Resend(process.env.RESEND_API_KEY);
    await resend.emails.send({
      from: 'CR Réunions <onboarding@resend.dev>',
      to: destinataire,
      subject: `CR — ${meta.titre} — ${meta.date}`,
      html: `<p>Bonjour,</p><p>Veuillez trouver ci-joint le compte-rendu de la réunion <strong>${meta.titre}</strong> du ${meta.date}.</p><p>Ce document a été généré automatiquement par l'application CR Réunions.</p>`,
      attachments: [{
        filename,
        content: buffer.toString('base64'),
      }]
    });

    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Serveur démarré sur le port ${PORT}`));
