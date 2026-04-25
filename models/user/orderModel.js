const mongoose = require('mongoose')
const { Schema } = mongoose

const orderItemSchema = new Schema({
  productId:    { type: Schema.Types.ObjectId, ref: 'Product', required: true },
  productName:  { type: String, required: true },
  productImage: { type: String, default: '' },
  shade:        { type: String, default: '' },
  quantity:     { type: Number, required: true, min: 1 },
  priceAtOrder: { type: Number, required: true },
  salePrice:    { type: Number, default: 0 },
  discount:     { type: Number, default: 0 },
}, { _id: true })

const orderSchema = new Schema(
  {
    userId: { type: Schema.Types.ObjectId, ref: 'User', required: true },

    orderId: {
      type: String,
      unique: true,
    },

    items: [orderItemSchema],

    shippingAddress: {
      name:        { type: String, required: true },
      address:     { type: String, required: true },
      city:        { type: String, default: '' },
      state:       { type: String, required: true },
      country:     { type: String, required: true },
      pincode:     { type: String, required: true },
      mobile:      { type: String, required: true },
      email:       { type: String, default: '' },
      landmark:    { type: String, default: '' },
      addressType: { type: String, default: 'Home' },
    },

    subtotal:      { type: Number, required: true },
    totalDiscount: { type: Number, default: 0 },
    shippingCharge:{ type: Number, default: 0 },
    tax:           { type: Number, default: 0 },
    finalAmount:   { type: Number, required: true },

    couponCode:    { type: String,  default: null  },
    couponDiscount:{ type: Number,  default: 0     },
    couponVoided:  { type: Boolean, default: false },   // ← ADDED HERE (inside fields)

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

    razorpayOrderId:   { type: String, default: null },
    razorpayPaymentId: { type: String, default: null },

    itemStatuses: [{
      itemId:       { type: Schema.Types.ObjectId },
      status:       { type: String, enum: ['Active', 'Cancelled', 'Returned'], default: 'Active' },
      cancelReason: { type: String, default: '' },
      returnReason: { type: String, default: '' },
      returnStatus: { type: String, enum: ['None', 'Requested', 'Approved', 'Rejected', 'Completed'], default: 'None' },
    }],

    cancelReason: { type: String, default: '' },
    returnReason: { type: String, default: '' },
    returnStatus: {
      type: String,
      enum: ['None', 'Requested', 'Approved', 'Rejected', 'Completed'],
      default: 'None'
    },
    returnedAt:  { type: Date, default: null },
    deliveredAt: { type: Date, default: null },
    cancelledAt: { type: Date, default: null },
  },                  // ← fields object closes HERE — couponVoided is safely inside ↑
  { timestamps: true }  // ← schema options object comes second
)

orderSchema.pre('save', async function () {
  if (!this.orderId) {
    const timestamp = Date.now().toString(36).toUpperCase()
    const random    = Math.random().toString(36).substring(2, 6).toUpperCase()
    this.orderId    = `BB-${timestamp}-${random}`
  }
})

module.exports = mongoose.model('Order', orderSchema)