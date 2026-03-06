const User = require('../../models/user/userModel')
const Address = require('../../models/user/addressModel')

 
const loadProfile = async (req,res) =>{
  try {
    const userId = req.session.user._id;
   const user = await User.findById(userId).lean();
    const address  = await Address.findOne({user:userId,isDefault:true});
      const success = req.query.success === 'true' ? 'Profile updated successfully!' : null;
      return res.render('user/userProfile',{ user, address: address || null, success,errors:null })
    
  } catch (error) {
    res.status(500).json({error:'page not loading'})
  }
}



const updateProfile = async (req, res) => {

        try {
          console.log("SESSION:", req.session.user);
console.log("BODY:", req.body);
   const findUser = await User.findById(req.session.user._id).lean();

   if (!findUser) {
     return res.render('user/userProfile', { error: "User not found", user: findUser });
   }

   const { fullName, email, phoneNumber } = req.body;
   const errors = {};
   
       const nameRegex = /^[A-Za-z\s]{3,}$/;
       if (!fullName) {
           errors.fullName = "Full name is required";
       } else if (!nameRegex.test(fullName)) {
           errors.fullName = "Enter a valid full name (only letters, min 3 chars)";
       }

       const phoneRegex = /^\d{10}$/;
       if (!phoneNumber) {
           errors.phoneNumber = "Phone number is required";
       } else if (!phoneRegex.test(phoneNumber)) {
           errors.phoneNumber = "Enter a valid 10-digit phone number";
       }

       const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
       if (!email) {
           errors.email = "Email is required";
       } else if (!emailRegex.test(email)) {
           errors.email = "Enter a valid email address";
       } else {
           const existingUser = await User.findOne({ email });
           if (existingUser && email != req.session.user.email) {
               errors.email = "Email already registered";
           }
       }
      
       if (Object.keys(errors).length > 0) {
           return res.render('user/userProfile', {
               errors,
               user:findUser,
               success:null
           });
       }

        const updateData = { fullName, phoneNumber };
    if (req.file) {
   updateData.profilePhoto = '/uploads/profiles/' + req.file.filename;
    }
  const updatedUser = await User.findByIdAndUpdate(
     req.session.user._id,
     updateData,
     { new: true }
   ).lean();

   
   req.session.user = {
     _id: updatedUser._id,
     fullName: updatedUser.fullName,   // ← key fix
     email: updatedUser.email,
     isBlocked: updatedUser.isBlocked
   
   };

   
   res.redirect('/profile');

 } catch (error) {
   console.log(error);
   res.status(500).json({ error: "cannot update" });
 }
}


         

        
module.exports = {
    loadProfile,
    updateProfile
}