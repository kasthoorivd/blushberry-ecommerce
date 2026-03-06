const passport = require('passport')
const express = require('express')
const multer = require('multer')
const path = require('path')
const userRouter = express.Router()

const userController = require('../controllers/user/userAuthController')
const profileController = require('../controllers/user/profileController')
const {isLoggedIn, isLoggedOut} = require('../middleware/authMiddleware')



const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, 'public/uploads/profiles/'),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname)
    cb(null, `profile-${req.session.user._id}-${Date.now()}${ext}`)
  }
})

const upload = multer({
  storage,
  limits: { fileSize: 5 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowed = ['image/jpeg', 'image/png', 'image/webp']
    allowed.includes(file.mimetype) ? cb(null, true) : cb(new Error('Only JPG/PNG allowed'))
  }
}) 

// home
userRouter.get('/', userController.loadHomePage)

// signup
userRouter.get('/signup', isLoggedOut, userController.loadSignUp)
userRouter.post('/signup', isLoggedOut, userController.signup)

// login
userRouter.get('/login', isLoggedOut, userController.loadLogin)
userRouter.post('/login', isLoggedOut, userController.login)

// otp
userRouter.get('/otp', userController.loadOtpPage)
userRouter.post('/verifyOtp', userController.verifyOtp)
userRouter.post('/resendOtp', userController.resendOtp)

// google auth — FIXED: removed trailing empty () from first route
userRouter.get('/auth/google', passport.authenticate('google', { scope: ['profile', 'email'] }))

// userRouter.get('/auth/google/callback',
//   passport.authenticate('google', { failureRedirect: '/login' }),
//   (req, res) => {
//     // FIXED: sync req.user (set by Passport) into req.session.user
//     // so loadHomePage and all middleware can find the logged-in user
//     req.session.user = {
//       _id: req.user._id,
//       email: req.user.email,
//       isBlocked: req.user.isBlocked
//     };

//     req.session.save((err) => {
//       if (err) {
//         console.error('Session save error after Google login:', err);
//         return res.redirect('/login');
//       }
//       res.redirect('/');
//     });
//   }
// )
userRouter.get('/auth/google/callback',
  passport.authenticate('google', { failureRedirect: '/login' }),
  (req, res) => {
    // Sync passport user into session.user
    req.session.user = req.user; // ← add this
    res.redirect('/');
  }
);
// forgot password
userRouter.get('/forgot-password', userController.LoadforgotPassword)
userRouter.post('/api/forgot-password', userController.forgotPassword)
userRouter.get('/reset-password', userController.showResetPage)
userRouter.post('/api/auth/reset-password', userController.resetPassword)
userRouter.get('/otp-forgot-password', userController.showForgotOtpPage)


//profile
userRouter.get('/profile',profileController.loadProfile)
userRouter.post('/profile',upload.single('profilePhoto'),profileController.updateProfile)
// logout
userRouter.get('/logout', isLoggedIn, userController.logout)

module.exports = userRouter