const mongoose = require('mongoose');
const { Schema } = mongoose;

const addressSchema = new Schema(
  {
    user: {
      type: Schema.Types.ObjectId,
      ref: 'User',
      required: true
    },

    name: {
      type: String,
      required: true,
      trim: true
    },

    country: {
      type: String,
      required: true,
      trim: true
    },

    state: {
      type: String,
      required: true,
      trim: true
    },

    address: {
      type: String,
      required: true,
      trim: true
    },

    city: {
      type: String,
      required: false,
      trim: true,
      default: null
    },

    pincode: {
      type: String,
      required: true,
      trim: true
    },

    addressType: {
      type: String,
      enum: ['Home', 'Work'],
      default: 'Home'
    },

    landmark: {
      type: String,
      trim: true,
      default: null
    },

    mobile: {
      type: String,
      required: true,
      trim: true
    },

    email: {
      type: String,
      required: true,
      trim: true,
      lowercase: true
    },

    alternateNumber: {
      type: String,
      trim: true,
      default: null
    },

    isDefault: {
      type: Boolean,
      default: false
    }
  },
  {
    timestamps: true
  }
);

module.exports = mongoose.model('Address', addressSchema);