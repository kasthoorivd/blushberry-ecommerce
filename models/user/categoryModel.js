const mongoose = require('mongoose');
const {Schema}=mongoose;

const categorySchema = new Schema(
  {
    name: {
      type: String,
      required: true,
      unique: true,
      trim: true,
    },
    description: {
      type: String,
      required: true,
    },
    offer: {
      type: Number, 
      default: 0,
    },
    stock: {
      type: Number,
      default: 0,
    },
    
    isListed: {
      type: Boolean,
      default: true,
    },
    isDeleted: {
      type: Boolean,
      default:false,
    },
  },

  {
    timestamps:true,
  }
);

module.exports = mongoose.model('Category', categorySchema);