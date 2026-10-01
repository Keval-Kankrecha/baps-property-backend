require('dotenv').config();
const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');
const propertyRoutes = require('./routes/properties');
const uploadRoutes = require('./routes/uploads');

const app = express();
app.use(cors({ origin: process.env.CORS_ORIGIN || 'http://localhost:4200' }));
app.use(express.json({ limit: '1mb' }));

app.get('/api/health', (req, res) => {
  res.json({ ok: true, db: mongoose.connection.readyState === 1 ? 'connected' : 'disconnected' });
});
app.use('/api/uploads', uploadRoutes.router);
app.use('/files', express.static(uploadRoutes.UPLOAD_DIR, { index: false, dotfiles: 'deny' }));
app.use('/api/properties', propertyRoutes);

// eslint-disable-next-line no-unused-vars
app.use((err, req, res, next) => {
  if (err.name === 'ValidationError' || err.name === 'CastError') {
    return res.status(400).json({ message: err.message });
  }
  console.error(err);
  res.status(500).json({ message: 'Server error' });
});

const PORT = process.env.PORT || 3000;
const MONGODB_URI = process.env.MONGODB_URI || 'mongodb://127.0.0.1:27017/baps_property';

mongoose
  .connect(MONGODB_URI)
  .then(() => {
    console.log('MongoDB connected');
    app.listen(PORT, () => console.log(`API running on http://localhost:${PORT}`));
  })
  .catch((err) => {
    console.error('MongoDB connection failed:', err.message);
    process.exit(1);
  });
