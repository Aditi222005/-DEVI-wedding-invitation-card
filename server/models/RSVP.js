const mongoose = require('mongoose');

const rsvpSchema = new mongoose.Schema({
  name: {
    type: String,
    required: [true, 'Name is required'],
    trim: true,
  },
  guestCount: {
    type: Number,
    required: true,
    min: 1,
    max: 20,
    default: 1,
  },
  attending: {
    type: Boolean,
    required: true,
  },
  deviceId: {
    type: String,
    required: true,
    unique: true,          // one submission per device
    index: true,
  },
  submittedAt: {
    type: Date,
    default: Date.now,
  },
  ipAddress: {
    type: String,
    default: '',
  },
  userAgent: {
    type: String,
    default: '',
  },
}, {
  timestamps: true,
});

module.exports = mongoose.model('RSVP', rsvpSchema);
