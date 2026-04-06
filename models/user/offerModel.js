const mongoose = require('mongoose')

const offerSchema = new mongoose.Schema({
  name:        { type: String, required: true },
  type:        { type: String, enum: ['product', 'category'], required: true },
  discountPercent: { type: Number, required: true, min: 1, max: 99 },
  targetId:    { type: mongoose.Schema.Types.ObjectId, required: true },
  // productId if type='product', categoryId if type='category'
  startDate:   { type: Date, required: true },
  endDate:     { type: Date, required: true },
  isActive:    { type: Boolean, default: true }
}, { timestamps: true })

module.exports = mongoose.model('Offer', offerSchema)