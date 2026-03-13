const mongoose = require('mongoose')

const userSchema = new mongoose.Schema({
   
    fullName :{
        type:String,
        required:true
    },

    email:{
        type:String,
        required:true,
        unique:true,
        lowercase:true

    },

    password :{
        type:String,
        required:false
    },

    googleId:{
        type:String,
        default:null,
    },
     profilePhoto:{
        type:String
     },
     phoneNumber:{
      type:Number,
      type: String,
      required: false,
      unique: false,
   
      default: null,
     },
     
    authProviders:{
         type:[String], //['manual','google']
         default:[]
    },
    isAdmin:{
        type:Boolean,
        default:false
    },

    isBlocked:{
        type:Boolean,
        default:false
    },

    isVerified :{
        type:Boolean ,
        default :false
    },

    createdAt :{
        type:Date,
        default : Date.now
    }
})

module.exports = mongoose.model('User',userSchema) 