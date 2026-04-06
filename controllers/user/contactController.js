const User = require('../../models/user/userModel')

const loadContact = async(req,res)=>{
   res.render('user/contact')
}

module.exports ={
    loadContact
}