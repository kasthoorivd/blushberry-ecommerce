const User = require('../../models/user/userModel')
const Otp = require('../../models/user/otpModel')
const bcrypt = require('bcrypt')
const generateOtp = require('../../utils/generateOtp');
const sendEmail = require("../../utils/sendEmail");
const passport = require('../../config/passport');

const loadHomePage = (req, res) => {

   res.render("user/homePage.ejs");
}

const loadSignUp = (req, res) => {
   try {
      res.render('user/userSignup.ejs')
   } catch (error) {
      console.log(`error from signUpPage ${error}`)
   }
};


const signup = async (req, res) => {
   try {

      const { fullName, email, password, confirmPassword } = req.body;
      console.log(req.body);

      if (password !== confirmPassword) {
         return res.json({
            success: false,
            message: 'Passwords do not match'
         });
      }

      const existingUser = await User.findOne({ email });

      if (existingUser) {
         return res.json({
            success: false,
            message: 'User already exists'
         });
      }

      //  Generate OTP
      const otp = generateOtp();
   
      //  Create expiry time
      const expiresAt = Date.now() + 60 * 1000;

      await Otp.deleteMany({ email });

      await Otp.create({
         email,
         otp,
         expiresAt
      });

      req.session.tempUser = {
         fullName,
         email,
         password
      };

      req.session.otpExpiresAt = expiresAt;

      await sendEmail(email, otp);

      return res.json({
         success: true,
         redirectUrl: "/otp",
         expiresAt
      });

   } catch (error) {
      console.log(error);
      return res.status(500).json({
         success: false,
         message: 'Server error'
      });
   }
};




const verifyOtp = async (req, res) => {
   try {
      const { otp } = req.body;
      const tempUser = req.session.tempUser;

      if (!tempUser) {
         return res.json({
            success: false,
            message: "Session expired. Please signup again."
         });
      }

      const otpRecord = await Otp.findOne({ email: tempUser.email });

      if (!otpRecord) {
         return res.json({
            success: false,
            message: "OTP expired ,please resend otp "
         });
      }

      if (otpRecord.otp.toString() !== otp.toString()) {
         return res.json({
            success: false,
            message: "Invalid Otp"
         });
      }

      const hashedPassword = await bcrypt.hash(tempUser.password, 10);

      console.log("Before creating user");

      const newUser = await User.create({
         fullName: tempUser.fullName,
         email: tempUser.email,
         password: hashedPassword,
         isVerified: true
      });

      console.log("User Saved Successfully:", newUser);
      await Otp.deleteOne({ email: tempUser.email });

req.session.user = {
   id: newUser._id,
   name: newUser.fullName,
   email: newUser.email
};

req.session.tempUser = null;

return res.json({ success: true });

   

   } catch (error) {
      console.log(" VERIFY OTP ERROR FULL:");
      console.log(error);   // VERY IMPORTANT
      return res.status(500).json({
         success: false,
         message: error.message
      });
   }
}



const resendOtp = async (req, res) => {
   try {

      const tempUser = req.session.tempUser;

      if (!tempUser) {
         return res.status(400).json({
            success: false,
            message: "Session expired. Please signup again."
         });
      }

      const email = tempUser.email;

      await Otp.deleteMany({ email });

      const newOtp = generateOtp();

      const expiresAt = Date.now() + 60 * 1000;

      await Otp.create({
         email,
         otp: newOtp,
         expiresAt
      });

      req.session.otpExpiresAt = expiresAt;

      await sendEmail(email, newOtp);

      return res.json({
         success: true,
         expiresAt
      });

   } catch (error) {
      console.log("RESEND OTP ERROR:", error);
      return res.status(500).json({
         success: false,
         message: "Something went wrong"
      });
   }
};



const loadLogin = (req, res) => {
   try {
      res.render('user/loginPage.ejs')
   } catch (error) {
      console.log(`error from loginPage ${error}`)
   }
}



// Login function
const login = (req, res, next) => {
    passport.authenticate("local", (err, user, info) => {
        if (err) return next(err);

        if (!user) {
            return res.json({
                success: false,
                message: info.message
            });
        }

        req.login(user, (err) => {
            if (err) return next(err);

            return res.json({ success: true });
        });
    })(req, res, next);
};



const logout = (req, res) => {
   req.session.destroy((err) => {
      if (err) {
         console.log(err)
      }
      res.clearCookie('connect.sid')
      res.redirect('/login')
   })
}


const loadOtpPage = (req, res) => {
   res.render('user/otpPage.ejs')
}

const forgotPassword = (req,res) =>{
   res.render('user/forgotPassword.ejs')
}
module.exports = {
   loadHomePage,
   loadSignUp,
   loadLogin,
   signup,
   login,
   logout,
   verifyOtp,
   resendOtp,
   loadOtpPage,
   forgotPassword

};