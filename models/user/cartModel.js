const mongoose = require('mongoose')

const cartItemSchema = new mongoose.Schema({
  productId: {
    type:     mongoose.Schema.Types.ObjectId,
    ref:      'Product',
    required: true
  },
  shade:    { type: String, default: '' },   // selected variant shade
  quantity: { type: Number, required: true, min: 1, default: 1 },
  // snapshot price at time of adding (so price changes don't silently affect cart)
  priceAtAdd: { type: Number, required: true }
}, { _id: true })

const cartSchema = new mongoose.Schema(
  {
    userId: {
      type:     mongoose.Schema.Types.ObjectId,
      ref:      'User',
      required: true,
      unique:   true          // one cart per user
    },
    items: [cartItemSchema]
  },
  { timestamps: true }
)

// virtual: total item count
cartSchema.virtual('totalItems').get(function () {
  return this.items.reduce((sum, i) => sum + i.quantity, 0)
})

module.exports = mongoose.model('Cart', cartSchema)