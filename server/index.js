require('dotenv').config();
const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');
const path = require('path');

const rsvpRoutes = require('./routes/rsvp');

const app = express();
const PORT = process.env.PORT || 3001;

// ── Middleware ──────────────────────────────────────────────────────────────
app.use(cors({
  origin: process.env.CORS_ORIGIN || '*',
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type'],
}));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// ── Routes ──────────────────────────────────────────────────────────────────
app.use('/api/rsvp', rsvpRoutes);

// Health check
app.get('/health', (req, res) => res.json({ status: 'ok', ts: new Date() }));

// ── MongoDB connection ───────────────────────────────────────────────────────
const MONGO_URI = process.env.MONGODB_URI;
if (!MONGO_URI || MONGO_URI.includes('replace_me')) {
  console.error('\n❌  MONGODB_URI is not set in server/.env');
  console.error('   Please open server/.env and replace the placeholder with your real connection string.\n');
  process.exit(1);
}

mongoose
  .connect(MONGO_URI)
  .then(() => {
    console.log('✅  MongoDB connected');
    app.listen(PORT, () => {
      console.log(`🚀  RSVP API running → http://localhost:${PORT}`);
      console.log(`    Admin list  → http://localhost:${PORT}/api/rsvp/all`);
      console.log(`    Download XL → http://localhost:${PORT}/api/rsvp/download`);
    });
  })
  .catch(err => {
    console.error('❌  MongoDB connection failed:', err.message);
    process.exit(1);
  });
