const User = require('../../models/user/userModel')
const bcrypt = require('bcrypt')
const { loadLogin } = require('../user/userAuthController')


const loadAdminLogin = (req,res) =>{
     res.render('admin/adminLogin.ejs')
}


const loadDashboard = (req,res) =>{
    res.render('admin/dashboard.ejs')
}

const adminLogin = async (req,res) =>{
    try {
        const {email,password} = req.body;

        const user = await User.findOne({email});

        if(!user){
            return res.json({
                success:false,
                message:"admin not found"
            })
        }

        if(!user.isAdmin) {
            return res.json({
                success:false,
                message:'you are not authorised as Admin'
            })
        }
        const isMatch = await bcyrpt.compare(password,user.password)
      if (!isMatch) {
        return res.json({
            success:false,
            message:'Incorrect password'
        })
      }

      req.session.admin = user._id;


      res.json({
        success:true
      })
    } catch (error) {
        console.log("Admin login error",error)
        res.status(500).json({
            success:false,
            message:"server error"
        })
    }
}
module.exports = {
   loadAdminLogin,
   loadDashboard,
   adminLogin
}