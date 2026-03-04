const mongoose = require("mongoose");

const otpSchema = new mongoose.Schema({
    email: {
        type: String,
        required: true
    },
    otp: {
        type: String,
        required: true
    },
    createdAt: {
        type: Date,
        default: Date.now,
        expires: 60
    },
     purpose: { type: String, 
        enum: ['signup', 'forgot'],
         required: true }
});

module.exports = mongoose.model("Otp", otpSchema);