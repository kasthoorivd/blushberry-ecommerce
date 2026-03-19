const mongoose = require('mongoose')


const couponSchema = new mongoose.Schema({
    code: {
        type: String,
        required: true,
        unique:true,
        uppercase:true,
        trim:true
    },
    discountAmount : {
        type: Number,
        required: true
    },
    minOrderAmount:{
        type:Number,
        default:0
    },
    maxUses:{
        type:Number,
        default:null
    },
    usedBy:[{
        type:mongoose.Schema.Types.ObjectId,
        ref:'User'
    }],
    isActive:{
        type:Boolean,
        default:true
    },
    expiresAt: {
        type:Date,
        default:null
    }
},{timestamps:true})

module.exports = mongoose.model('Coupon',couponSchema)