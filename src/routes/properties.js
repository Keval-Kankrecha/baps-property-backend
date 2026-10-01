const express = require('express');
const mongoose = require('mongoose');
const Property = require('../models/Property');

const fs = require('fs/promises');
const path = require('path');
const { UPLOAD_DIR } = require('./uploads');

const router = express.Router();

const FILE_FIELDS = ['propertyCardFiles', 'assessmentFiles', 'deedFiles'];
const removeFile = (f) => (f?.fileName ? fs.unlink(path.join(UPLOAD_DIR, f.fileName)).catch(() => {}) : null);

const escapeRegex = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

// GET /api/properties?search=&district=&taluka=&isNA=&purpose=&page=1&limit=20
router.get('/', async (req, res, next) => {
  try {
    const { search, district, taluka, isNA, purpose, comesUnder } = req.query;
    const page = Math.max(1, parseInt(req.query.page, 10) || 1);
    const limit = Math.min(200, Math.max(1, parseInt(req.query.limit, 10) || 20));

    const filter = {};
    if (district) filter.district = district;
    if (taluka) filter.taluka = taluka;
    if (isNA) filter.isNA = isNA;
    if (purpose) filter.purpose = purpose;
    if (comesUnder) filter.comesUnder = comesUnder;
    if (search) {
      const rx = new RegExp(escapeRegex(search.trim()), 'i');
      filter.$or = ['propertyId', 'village', 'taluka', 'district', 'surveyDetails', 'deedNo', 'giverNameAddress', 'remarks']
        .map((f) => ({ [f]: rx }));
    }

    const [items, total] = await Promise.all([
      Property.find(filter).sort({ serialNo: 1 }).skip((page - 1) * limit).limit(limit).lean(),
      Property.countDocuments(filter),
    ]);
    res.json({ items, total, page, limit });
  } catch (err) {
    next(err);
  }
});

// Distinct values for filter dropdowns
router.get('/filters', async (req, res, next) => {
  try {
    const [districts, talukas, purposes] = await Promise.all([
      Property.distinct('district'),
      Property.distinct('taluka'),
      Property.distinct('purpose'),
    ]);
    const clean = (a) => a.filter(Boolean).sort((x, y) => x.localeCompare(y));
    res.json({ districts: clean(districts), talukas: clean(talukas), purposes: clean(purposes) });
  } catch (err) {
    next(err);
  }
});

router.param('id', (req, res, next, id) => {
  if (!mongoose.isValidObjectId(id)) return res.status(404).json({ message: 'Property not found' });
  next();
});

router.get('/:id', async (req, res, next) => {
  try {
    const item = await Property.findById(req.params.id).lean();
    if (!item) return res.status(404).json({ message: 'Property not found' });
    res.json(item);
  } catch (err) {
    next(err);
  }
});

router.post('/', async (req, res, next) => {
  try {
    const { _id, createdAt, updatedAt, ...data } = req.body;
    const item = await Property.create(data);
    res.status(201).json(item);
  } catch (err) {
    next(err);
  }
});

router.put('/:id', async (req, res, next) => {
  try {
    const { _id, createdAt, updatedAt, ...data } = req.body;
    const old = await Property.findById(req.params.id).lean();
    if (!old) return res.status(404).json({ message: 'Property not found' });
    const item = await Property.findByIdAndUpdate(req.params.id, data, { new: true, runValidators: true });
    // delete files that were replaced or removed
    for (const f of FILE_FIELDS) {
      const keep = new Set((item[f] || []).map((x) => x.fileName));
      for (const file of old[f] || []) if (!keep.has(file.fileName)) await removeFile(file);
    }
    res.json(item);
  } catch (err) {
    next(err);
  }
});

router.delete('/:id', async (req, res, next) => {
  try {
    const item = await Property.findByIdAndDelete(req.params.id);
    if (!item) return res.status(404).json({ message: 'Property not found' });
    for (const f of FILE_FIELDS) for (const file of item[f] || []) await removeFile(file);
    res.status(204).end();
  } catch (err) {
    next(err);
  }
});

module.exports = router;
