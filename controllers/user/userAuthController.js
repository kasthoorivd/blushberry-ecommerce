const User = require('../../models/user/userModel');
const Otp = require('../../models/user/otpModel');
const Product = require('../../models/user/productModel')
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
      user = await User.findById(userId); 
      console.log('profilePhoto:', user?.profilePhoto);
    }

    // ── fetch 8 newest products for the bestsellers grid ──
    const products = await Product.find({ isDeleted: false, isListed: true })
      .populate('categoryId', 'name')
      .sort({ createdAt: -1 })
      .limit(8)
      .lean()

    products.forEach(p => {
      const prices     = p.variants.map(v => v.salePrice > 0 ? v.salePrice : v.varientPrice)
      p.displayPrice   = Math.min(...prices)
      const origPrices = p.variants.map(v => v.varientPrice)
      p.originalPrice  = Math.min(...origPrices)
      p.displayOffer   = p.offer || 0
      p.inStock        = p.variants.some(v => v.stock > 0)
    })

    res.render('user/homePage', { user, products });  
  } catch (error) {
    console.log(error.message);
    res.render('user/homePage', { user: null, products: [] }); 
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
    const expiresAt = Date.now() + 5 * 60 * 1000;

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


// Verify OTP (signup, forgot password, or email change)
const verifyOtp = async (req, res) => {
  try {
    const { otp } = req.body;

    // Determine which flow we're in
    let flow, email;

    if (req.session.tempUser) {
      flow  = 'signup';
      email = req.session.tempUser.email;
    } else if (req.session.pendingEmail) {
      flow  = 'emailChange';
      email = req.session.pendingEmail;  // OTP was sent to the NEW email
    } else {
      flow  = 'forgot';
      email = req.session.forgotEmail;
    }

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

    if (!otpRecord.otp) {
      return res.json({ success: false, message: "OTP not generated. Please resend OTP." });
    }

    if (otpRecord.otp.toString() !== otp.toString()) {
      return res.json({ success: false, message: "Invalid OTP" });
    }

    // ── Signup flow ──
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

          req.session.user = {
            _id: newUser._id,
            email: newUser.email,
            isBlocked: newUser.isBlocked,
              fullName: newUser.fullName,  
               profilePhoto: newUser.profilePhoto
          };

          return res.json({ success: true, redirectUrl: '/' });
        });
      });

    // ── Email change flow ──
    } else if (flow === 'emailChange') {
      const userId = req.session.user?._id;

      await User.findByIdAndUpdate(userId, { email });
      await Otp.deleteOne({ email, purpose: flow });

      // Update session with new email
      req.session.user.email = email;
      req.session.pendingEmail = null;

      return res.json({ success: true, redirectUrl: '/profile', message: 'Email updated successfully' });

    // ── Forgot password flow ──
    } else {
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
    // Determine flow — same logic as verifyOtp
    let flow, email;

    if (req.session.tempUser) {
      flow  = 'signup';
      email = req.session.tempUser.email;
    } else if (req.session.pendingEmail) {
      flow  = 'emailChange';
      email = req.session.pendingEmail;
    } else {
      flow  = 'forgot';
      email = req.session.forgotEmail;
    }

    if (!email) {
      return res.status(400).json({
        success: false,
        message: "Session expired. Please start again."
      });
    }

    await Otp.deleteMany({ email, purpose: flow });

    const newOtp = generateOtp();
    const expiresAt = Date.now() + 5 * 60 * 1000;

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
  const blockedMessage = req.query.blocked === 'true'
    ? 'Your account has been blocked by the admin. Please contact support.'
    : null

  res.render('user/loginPage', { blockedMessage })
}

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

    if (user.isBlocked) {
      return res.json({
        success: false,
        redirectURL: '/login',
        message: 'user is blocked, cannot login'
      });
    }

    req.session.user = {
      _id: user._id,
      email: user.email,
      isBlocked: user.isBlocked,
      fullName: user.fullName,
       profilePhoto: user.profilePhoto
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
const showForgotOtpPage = async (req, res) => {
  if (!req.session.forgotEmail) return res.redirect('/forgot-password');

  const otpRecord = await Otp.findOne({ email: req.session.forgotEmail, purpose: 'forgot' });

  const expiresAt = otpRecord
    ? (otpRecord.expiresAt instanceof Date
        ? otpRecord.expiresAt.getTime()
        : Number(otpRecord.expiresAt))
    : null;

  res.render('user/otpPage.ejs', { isForgotPassword: true, expiresAt });
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


// Load OTP page
const loadOtpPage = async (req, res) => {
  try {
    const email   = req.session.tempUser?.email 
                  || req.session.pendingEmail 
                  || req.session.forgotEmail;

    const purpose = req.session.tempUser    ? 'signup'
                  : req.session.pendingEmail ? 'emailChange'
                  : 'forgot';

    if (!email) return res.redirect('/signup');

    const otpRecord = await Otp.findOne({ email, purpose });

    if (!otpRecord) return res.redirect('/signup');

    const expiresAt = otpRecord.expiresAt instanceof Date
      ? otpRecord.expiresAt.getTime()
      : Number(otpRecord.expiresAt);

    res.render('user/otpPage.ejs', { expiresAt });

  } catch (error) {
    console.error(error);
    res.redirect('/signup');
  }
};


// Request email change — sends OTP to new email, then redirects to OTP page
const requestEmailChange = async (req, res) => {
  try {
    const { newEmail } = req.body;
    const userId = req.session.user?._id;

    if (!userId) {
      return res.json({ success: false, message: "Not logged in" });
    }

    if (!newEmail) {
      return res.json({ success: false, message: "New email is required" });
    }

    // Check new email not already taken
    const existing = await User.findOne({ email: newEmail });
    if (existing) {
      return res.json({ success: false, message: "Email already in use" });
    }

    const otp = generateOtp();
    const expiresAt = Date.now() + 5 * 60 * 1000;

    await Otp.deleteMany({ email: newEmail, purpose: 'emailChange' });
    await Otp.create({ email: newEmail, otp, expiresAt, purpose: 'emailChange' });

    // Store new email in session — verifyOtp will read this
    req.session.pendingEmail = newEmail;

    await sendEmail(newEmail, otp);

    return res.json({ success: true, redirectUrl: '/otp' });

  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, message: "Server error" });
  }
};


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
  showForgotOtpPage,
  requestEmailChange,
};