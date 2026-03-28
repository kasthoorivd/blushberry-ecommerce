const mongoose = require('mongoose')
const { Schema } = mongoose

const orderItemSchema = new Schema({
  productId: { type: Schema.Types.ObjectId, ref: 'Product', required: true },
  productName: { type: String, required: true },
  productImage: { type: String, default: '' },
  shade: { type: String, default: '' },
  quantity: { type: Number, required: true, min: 1 },
  priceAtOrder: { type: Number, required: true },  // original price
  salePrice:    { type: Number, default: 0 },       // sale price if any
  discount:     { type: Number, default: 0 },       // discount amount per item
}, { _id: true })

const orderSchema = new Schema(
  {
    userId: { type: Schema.Types.ObjectId, ref: 'User', required: true },

    orderId: {
      type: String,
      unique: true,
      // generated before save
    },

    items: [orderItemSchema],

    // address snapshot (so edits to address don't affect old orders)
    shippingAddress: {
      name:            { type: String, required: true },
      address:         { type: String, required: true },
      city:            { type: String, default: '' },
      state:           { type: String, required: true },
      country:         { type: String, required: true },
      pincode:         { type: String, required: true },
      mobile:          { type: String, required: true },
      email: { type: String, default: '' },
      landmark:        { type: String, default: '' },
      addressType:     { type: String, default: 'Home' },
    },

    // pricing breakdown
    subtotal:      { type: Number, required: true },   // sum of (priceAtOrder * qty)
    totalDiscount: { type: Number, default: 0 },       // coupon + item discounts
    shippingCharge:{ type: Number, default: 0 },
    tax:           { type: Number, default: 0 },
    finalAmount:   { type: Number, required: true },   // what user actually pays

    couponCode:    { type: String, default: null },
    couponDiscount:{ type: Number, default: 0 },

    paymentMethod: {
      type: String,
      enum: ['COD', 'Online', 'Wallet'],
      default: 'COD'
    },

    paymentStatus: {
      type: String,
      enum: ['Pending', 'Paid', 'Failed', 'Refunded'],
      default: 'Pending'
    },

    orderStatus: {
      type: String,
      enum: ['Placed', 'Processing', 'Shipped', 'Delivered', 'Cancelled', 'Returned'],
      default: 'Placed'
    },

    // per-item status for partial cancellations
    itemStatuses: [{
  itemId:       { type: Schema.Types.ObjectId },
  status:       { type: String, enum: ['Active', 'Cancelled', 'Returned'], default: 'Active' },
  cancelReason: { type: String, default: '' },
  returnReason: { type: String, default: '' },
  returnStatus: { type: String, enum: ['None', 'Requested', 'Approved', 'Rejected', 'Completed'], default: 'None' },
}],
    // inside orderSchema, add after itemStatuses:
cancelReason:  { type: String, default: '' },
returnReason:  { type: String, default: '' },
returnStatus:  {
  type: String,
  enum: ['None', 'Requested', 'Approved', 'Rejected', 'Completed'],
  default: 'None'
},
returnedAt:    { type: Date, default: null },

    deliveredAt:  { type: Date, default: null },
    cancelledAt:  { type: Date, default: null },
  },
  { timestamps: true }
)

orderSchema.pre('save', async function () {
  if (!this.orderId) {
    const timestamp = Date.now().toString(36).toUpperCase()
    const random    = Math.random().toString(36).substring(2, 6).toUpperCase()
    this.orderId    = `BB-${timestamp}-${random}`
  }
})

module.exports = mongoose.model('Order', orderSchema)