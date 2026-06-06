const express = require('express');
const RSVP = require('../models/RSVP');
const { rebuildExcel, EXCEL_PATH } = require('../utils/excelBuilder');
const { appendRSVPRow, rebuildSheet } = require('../utils/sheetsBuilder');
const path = require('path');

const router = express.Router();

// ── POST /api/rsvp  ──────────────────────────────────────────────────────────
// Submit a new RSVP.  Rejects if the device has already submitted.
router.post('/', async (req, res) => {
  try {
    const { name, guestCount, attending, deviceId } = req.body;

    // Basic validation
    if (!name || typeof attending !== 'boolean' || !deviceId) {
      return res.status(400).json({ error: 'name, attending (boolean) and deviceId are required.' });
    }

    // Check if this device already submitted
    const existing = await RSVP.findOne({ deviceId });
    if (existing) {
      return res.status(409).json({
        error: 'already_submitted',
        message: 'An RSVP from this device has already been recorded.',
        rsvp: {
          name: existing.name,
          attending: existing.attending,
          guestCount: existing.guestCount,
          submittedAt: existing.submittedAt,
        },
      });
    }

    // Save to MongoDB
    const rsvp = await RSVP.create({
      name: name.trim(),
      guestCount: Math.max(1, Math.min(20, parseInt(guestCount) || 1)),
      attending,
      deviceId,
      ipAddress: req.headers['x-forwarded-for'] || req.socket.remoteAddress || '',
      userAgent: req.headers['user-agent'] || '',
    });

    // Get total count for serial number
    const totalCount = await RSVP.countDocuments();

    // Rebuild local Excel + update Google Sheet in background
    const allRsvps = await RSVP.find({}).sort({ createdAt: 1 });
    rebuildExcel(allRsvps).catch(err => console.error('Excel rebuild error:', err));
    appendRSVPRow(rsvp, totalCount).catch(err => console.error('Google Sheets error:', err));

    return res.status(201).json({
      success: true,
      message: 'RSVP recorded successfully.',
      rsvp: {
        name: rsvp.name,
        attending: rsvp.attending,
        guestCount: rsvp.guestCount,
        submittedAt: rsvp.submittedAt,
      },
    });
  } catch (err) {
    // Mongo duplicate key (deviceId unique index)
    if (err.code === 11000) {
      return res.status(409).json({
        error: 'already_submitted',
        message: 'An RSVP from this device has already been recorded.',
      });
    }
    console.error('RSVP submit error:', err);
    return res.status(500).json({ error: 'Internal server error.' });
  }
});

// ── GET /api/rsvp  ──────────────────────────────────────────────────────────
// Check if this device already submitted (pass ?deviceId=xxx)
router.get('/check', async (req, res) => {
  const { deviceId } = req.query;
  if (!deviceId) return res.status(400).json({ error: 'deviceId query param required.' });

  const existing = await RSVP.findOne({ deviceId });
  if (existing) {
    return res.json({
      submitted: true,
      rsvp: {
        name: existing.name,
        attending: existing.attending,
        guestCount: existing.guestCount,
        submittedAt: existing.submittedAt,
      },
    });
  }
  return res.json({ submitted: false });
});

// ── GET /api/rsvp/all  ──────────────────────────────────────────────────────
// Admin: list all RSVPs (no auth needed for now — add middleware if desired)
router.get('/all', async (req, res) => {
  try {
    const rsvps = await RSVP.find({}).sort({ createdAt: -1 });
    const stats = {
      total: rsvps.length,
      attending: rsvps.filter(r => r.attending).length,
      declined: rsvps.filter(r => !r.attending).length,
      totalGuests: rsvps.filter(r => r.attending).reduce((s, r) => s + r.guestCount, 0),
    };
    return res.json({ stats, rsvps });
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: 'Internal server error.' });
  }
});

// ── GET /api/rsvp/download  ─────────────────────────────────────────────────
// Admin: download the latest Excel file
router.get('/download', async (req, res) => {
  try {
    const allRsvps = await RSVP.find({}).sort({ createdAt: 1 });
    await rebuildExcel(allRsvps);
    return res.download(EXCEL_PATH, 'Viraj_Devaki_RSVP.xlsx');
  } catch (err) {
    console.error(err);
    return res.status(500).json({ error: 'Could not generate Excel.' });
  }
});

// ── GET /api/rsvp/rebuild-sheet  ─────────────────────────────────────────────
// Admin: force-sync all MongoDB RSVPs → Google Sheet
router.get('/rebuild-sheet', async (req, res) => {
  try {
    const allRsvps = await RSVP.find({}).sort({ createdAt: 1 });
    await rebuildSheet(allRsvps);
    return res.json({ success: true, message: `Sheet rebuilt with ${allRsvps.length} entries.` });
  } catch (err) {
    console.error('Sheet rebuild error:', err.message);
    return res.status(500).json({ error: err.message });
  }
});

module.exports = router;
