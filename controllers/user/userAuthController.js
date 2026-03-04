
const User = require('../../models/user/userModel');
const Otp = require('../../models/user/otpModel');
const bcrypt = require('bcrypt');
const generateOtp = require('../../utils/generateOtp');
const sendEmail = require("../../utils/sendEmail");
const passport = require('../../config/passport');

// Load home page
// const loadHomePage = (req, res) => {
//    res.render("user/homePage.ejs");
// };
const loadHomePage = async (req, res) => {
  try {
    let user = null;

    if (req.session.user) {
      // Fetch latest user info from DB
      user = await User.findById(req.session.user._id);
    }

    // Pass user to homePage.ejs
    res.render('user/homePage', { user }); // <-- This is crucial
  } catch (error) {
    console.log(error.message);
    res.render('user/homePage', { user: null });
  }
};

// Load signup page
const loadSignUp = (req, res) => {
   try {
      res.render('user/userSignup.ejs');
   } catch (error) {
      console.log(`Error loading signup page: ${error}`);
   }
};

// Signup
const signup = async (req, res) => {
   try {
      const { fullName, email, password, confirmPassword } = req.body;
      if (password !== confirmPassword) {
         return res.json({ success: false, message: 'Passwords do not match' });
      }

      const existingUser = await User.findOne({ email });
      if (existingUser) {
         return res.json({ success: false, message: 'User already exists' });
      }

      // Generate OTP
      const otp = generateOtp();
      const expiresAt = Date.now() + 60 * 1000;

      await Otp.deleteMany({ email }); // remove any previous OTP
      await Otp.create({ email, otp, expiresAt, purpose: 'signup' });

      req.session.tempUser = { fullName, email, password };

      // req.session.otpExpiresAt = expiresAt;

      await sendEmail(email, otp);
      req.session.otpType="signup"
      
      return res.json({
         success: true,
         redirectUrl: "/otp",
         expiresAt,
         purpose: 'signup',
      });

   } catch (error) {
      console.error(error);
      return res.status(500).json({ success: false, message: 'Server error' });
   }
};

// Verify OTP (signup or forgot password)
// Verify OTP (signup or forgot password)
const verifyOtp = async (req, res) => {
   try {
      const { otp } = req.body;

      // Determine flow: signup or forgot password
      const flow = req.session.tempUser ? 'signup' : 'forgot'
      const email = flow === 'signup' ? req.session.tempUser.email : req.session.forgotEmail;

      if (!email) {
         return res.json({ success: false, message: "Session expired. Please start again." });
      }

      // Fetch OTP from DB
      const otpRecord = await Otp.findOne({ email, purpose: flow });

      if (!otpRecord) {
         return res.json({ success: false, message: "No OTP found. Please resend OTP." });
      }

      // Check expiration
      if (otpRecord.expiresAt < Date.now()) {
         return res.json({ success: false, message: "OTP expired, please resend OTP" });
      }

      // Use the same field for both flows
      const otpToCheck = otpRecord.otp; // change code -> otp in forgot flow
      if (!otpToCheck) {
         return res.json({ success: false, message: "OTP not generated. Please resend OTP." });
      }

      if (otpToCheck.toString() !== otp.toString()) {
         return res.json({ success: false, message: "Invalid OTP" });
      }

      if (flow === 'signup') {
         // Complete signup
         const hashedPassword = await bcrypt.hash(req.session.tempUser.password, 10);
         const newUser = await User.create({
            fullName: req.session.tempUser.fullName,
            email,
            password: hashedPassword,
            isVerified: true
         });

          req.login(newUser,(err)=>{
            if(err) return next(err);

            req.session.save(async(err)=>{
               if(err)  return next(err)

                  await Otp.deleteOne({email,purpose:flow})
                  req.session.tempUser = null;

                  return res.json({success:true,redirectUrl:'/'})
            })
          })
      } else {
         // Forgot password flow
         req.session.canResetPassword = true;

         // Delete OTP
         await Otp.deleteOne({ email, purpose: flow });

         return res.json({ success: true, redirectUrl: '/reset-password' });
      }

   } catch (error) {
      console.error("VERIFY OTP ERROR:", error);
      return res.status(500).json({ success: false, message: error.message });
   }
};


const resendOtp = async (req, res) => {
   try {

      const flow = req.session.tempUser ? 'signup' : 'forgot';
      const email = flow === 'signup'
         ? req.session.tempUser?.email
         : req.session.forgotEmail;

      if (!email) {
         return res.status(400).json({
            success: false,
            message: "Session expired. Please start again."
         });
      }

      // Delete old OTP
      await Otp.deleteMany({ email, purpose: flow });

      // Generate new OTP
      const newOtp = generateOtp();
      const expiresAt = Date.now() + 60 * 1000;

      // Always use same field name: otp
      await Otp.create({
         email,
         otp: newOtp,
         expiresAt,
         purpose: flow
      });

      await sendEmail(email, newOtp);

      return res.json({
         success: true,
         expiresAt
      });

   } catch (error) {
      console.error("RESEND OTP ERROR:", error);
      return res.status(500).json({
         success: false,
         message: "Something went wrong"
      });
   }
};

// Load login page
const loadLogin = (req, res) => {
   try {
  
      
      res.render('user/loginPage.ejs');
   } catch (error) {
      console.log(`Error loading login page: ${error}`);
   }
};

// Login function
// const login = (req, res, next) => {
//    console.log('hello i');
   
//   passport.authenticate("local", (err, user, info) => {
//     if (err) return next(err);
//     if (!user) return res.json({ success: false, message: info?.message || "Invalid credentials" });

//     req.login(user, (err) => {
//       if (err) return next(err);
//       return res.json({ success: true });
//     });
//   })(req, res, next);
// };

const login = async(req,res,next) =>{
   try {
      const {email,password} = req.body;
      const user = await User.findOne({email});
      console.log('user pass word====================================',user.password);
      
      if(!user) return res.json(401).json({message:'User not found'})

         const isMatch = await bcrypt.compare(password,user.password);
         if(!isMatch) return res.status(401).json({message:'Invalid password'})
          req.logIn(user,(err)=>{
         if(err) return next(err)

            return res.json({success:true,message:'Logged in successfully'})

         })  
   } catch (error) {
      next(error)
   }
}
// Logout
const logout = (req, res,next) => {
   req.logout(function(err) {
        if (err) return next(err);   // handle error
        req.session.destroy(() => {  // destroy session in store
            res.clearCookie('connect.sid'); // remove cookie from browser
            res.redirect('/login');   // redirect to login page
        });
    });
};

// Forgot password
const LoadforgotPassword = (req, res) => {
   res.render('user/forgotPassword.ejs');
};

const forgotPassword = async (req, res) => {
   try {
      const { email } = req.body;
      const user = await User.findOne({ email });

      if (!user) return res.status(400).json({ success: false, message: "Email not registered" });

      const otpCode = generateOtp();
      const expiresAt = Date.now() + 5 * 60 * 1000;

      console.log('work')
      await Otp.findOneAndUpdate(
         { email },
         { otp: otpCode, expiresAt, purpose: 'forgot' },
         { upsert: true }
      );

      await sendEmail(email, otpCode);
      req.session.forgotEmail = email;

      return res.json({ success: true, message: "OTP sent to your email" });

   } catch (err) {
      console.error(err);
      res.status(500).json({ success: false, message: "Server error" });
   }
};

// Show OTP page for forgot password
const showForgotOtpPage = (req, res) => {
   if (!req.session.forgotEmail) return res.redirect('/forgot-password');
   res.render('user/otpPage.ejs', { isForgotPassword: true });
};

// Show reset password page
const showResetPage = (req, res) => {
   if (!req.session.canResetPassword) return res.redirect('/forgot-password');
   res.render('user/resetPassword.ejs');
};

// Reset password
const resetPassword = async (req, res) => {
   try {
      if (!req.session.canResetPassword) {
         return res.status(400).json({ success: false, message: "Not authorized" });
      }
            console.log('jkhj')
      const email = req.session.forgotEmail;
      const { password, confirmPassword } = req.body;
   console.log(req.body)
      if (password !== confirmPassword) {
         return res.status(400).json({ success: false, message: "Passwords do not match" });
      }

      const hashedPassword = await bcrypt.hash(password, 10);
      await User.findOneAndUpdate({ email }, { password: hashedPassword });

      req.session.canResetPassword = false;
      req.session.forgotEmail = null;

      return res.json({ success: true, message: "Password reset successful" });

   } catch (err) {
      console.error(err);
      res.status(500).json({ success: false, message: "Server error" });
   }
};

// Load OTP page (generic)
// const loadOtpPage = (req, res) => {
//    res.render('user/otpPage.ejs');
// };

const loadOtpPage = async (req, res) => {
   try {
      const email = req.session.tempUser?.email || req.session.forgotEmail;

      if (!email) return res.redirect('/signup');

      const otpRecord = await Otp.findOne({ email }).sort({ createdAt: -1 });

      if (!otpRecord) return res.redirect('/signup');

      const expireTime = new Date(otpRecord.createdAt).getTime() + 60 * 1000 ;

      res.render('user/otpPage.ejs', { expiresAt: expireTime });

   } catch (error) {
      console.error(error);
      res.redirect('/signup');
   }
};

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
   LoadforgotPassword,
   forgotPassword,
   showResetPage,
   resetPassword,
   showForgotOtpPage
};