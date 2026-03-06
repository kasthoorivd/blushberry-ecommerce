const User = require('../../models/user/userModel');
const Otp = require('../../models/user/otpModel');
const bcrypt = require('bcrypt');
const generateOtp = require('../../utils/generateOtp');
const sendEmail = require("../../utils/sendEmail");
const passport = require('../../config/passport');
const { isBlocked } = require('../../middleware/authMiddleware');




const loadHomePage = async (req, res) => {
  try {
    let user = null;

    const userId = req.session.user?._id || req.user?._id;

    if (userId) {
      user = await User.findById(userId); // always fetch fresh from DB
      console.log('profilePhoto:', user?.profilePhoto);
    }

    res.render('user/homePage', { user });
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

    await Otp.deleteMany({ email });
    await Otp.create({ email, otp, expiresAt, purpose: 'signup' });

    req.session.tempUser = { fullName, email, password };

    await sendEmail(email, otp);

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
const verifyOtp = async (req, res) => {
  try {
    const { otp } = req.body;

    const flow = req.session.tempUser ? 'signup' : 'forgot';
    const email = flow === 'signup' ? req.session.tempUser.email : req.session.forgotEmail;

    if (!email) {
      return res.json({ success: false, message: "Session expired. Please start again." });
    }

    const otpRecord = await Otp.findOne({ email, purpose: flow });

    if (!otpRecord) {
      return res.json({ success: false, message: "No OTP found. Please resend OTP." });
    }

    if (otpRecord.expiresAt < Date.now()) {
      return res.json({ success: false, message: "OTP expired, please resend OTP" });
    }

    const otpToCheck = otpRecord.otp;
    if (!otpToCheck) {
      return res.json({ success: false, message: "OTP not generated. Please resend OTP." });
    }

    if (otpToCheck.toString() !== otp.toString()) {
      return res.json({ success: false, message: "Invalid OTP" });
    }

    if (flow === 'signup') {
      const hashedPassword = await bcrypt.hash(req.session.tempUser.password, 10);
      const newUser = await User.create({
        fullName: req.session.tempUser.fullName,
        email,
        password: hashedPassword,
        isVerified: true
      });

      req.login(newUser, (err) => {
        if (err) return next(err);

        req.session.save(async (err) => {
          if (err) return next(err);

          await Otp.deleteOne({ email, purpose: flow });
          req.session.tempUser = null;

          // Also set session.user so navbar works consistently
          req.session.user = {
            _id: newUser._id,
            email: newUser.email,
            isBlocked: newUser.isBlocked
          };

          return res.json({ success: true, redirectUrl: '/' });
        });
      });

    } else {
      // Forgot password flow
      req.session.canResetPassword = true;
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

    await Otp.deleteMany({ email, purpose: flow });

    const newOtp = generateOtp();
    const expiresAt = Date.now() + 60 * 1000;

    await Otp.create({ email, otp: newOtp, expiresAt, purpose: flow });

    await sendEmail(email, newOtp);

    return res.json({ success: true, expiresAt });

  } catch (error) {
    console.error("RESEND OTP ERROR:", error);
    return res.status(500).json({ success: false, message: "Something went wrong" });
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


const login = async (req, res) => {
  try {
    const { email, password } = req.body;

    if (!email || !password) {
      return res.json({ success: false, message: "Email and password required" });
    } 

    const user = await User.findOne({ email });

    if (!user) {
      return res.json({ success: false, message: "User not found" });
    }

    // Google-only accounts have no password — block plain login attempt
    if (!user.password) {
      return res.json({
        success: false,
        message: "This account uses Google Sign-In. Please login with Google."
      });
    }

    const isMatch = await bcrypt.compare(password, user.password);

    if (!isMatch) {
      return res.json({ success: false, message: "Invalid password" });
    }

    req.session.user = {
      _id: user._id,
      email: user.email,
      isBlocked: user.isBlocked
    };

    return res.json({ success: true, message: "Login successful" });

  } catch (error) {
    console.log(error);
    return res.json({ success: false, message: "Server error" });
  }
};


// Forgot password
const LoadforgotPassword = (req, res) => {
  res.render('user/forgotPassword.ejs');
};

const forgotPassword = async (req, res) => {
  try {
    const { email } = req.body;
    const user = await User.findOne({ email });

    if (!user) {
      return res.status(400).json({ success: false, message: "Email not registered" });
    }

    const otpCode = generateOtp();
    const expiresAt = Date.now() + 5 * 60 * 1000;

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

    const email = req.session.forgotEmail;
    const { password, confirmPassword } = req.body;

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


const loadOtpPage = async (req, res) => {
  try {
    const email = req.session.tempUser?.email || req.session.forgotEmail;

    if (!email) return res.redirect('/signup');

    const otpRecord = await Otp.findOne({ email }).sort({ createdAt: -1 });

    if (!otpRecord) return res.redirect('/signup');

    const expireTime = new Date(otpRecord.createdAt).getTime() + 60 * 1000;

    res.render('user/otpPage.ejs', { expiresAt: expireTime });

  } catch (error) {
    console.error(error);
    res.redirect('/signup');
  }
};


// Logout — destroys session entirely
const logout = async (req, res) => {
  try {
    req.session.destroy(() => {
      res.clearCookie('connect.sid');
      res.redirect('/');
    });
  } catch (error) {
    console.log('error during logout:', error);
    res.status(500).send('error during logout');
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