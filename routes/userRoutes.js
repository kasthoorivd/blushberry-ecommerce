const passport = require('passport')
const express = require('express')
const multer = require('multer')
const path = require('path')
const userRouter = express.Router()
const User = require('../models/user/userModel')
const userController = require('../controllers/user/userAuthController')
const profileController = require('../controllers/user/profileController')
const addressController = require('../controllers/user/addressController')
const productController = require('../controllers/user/productController')
const productDetailController = require('../controllers/user/productDetailController')
const cartController = require('../controllers/user/cartController')
const checkoutController = require('../controllers/user/checkoutController')
const walletController = require('../controllers/user/walletController')
const contactController = require('../controllers/user/contactController')
const {isLoggedIn, isLoggedOut,isBlocked} = require('../middleware/authMiddleware')



const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, 'public/uploads/profiles/'),
  filename: (req, file, cb) => {
    const unique = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, unique + path.extname(file.originalname));
  }
});

const fileFilter = (req, file, cb) => {
  const allowedTypes = ['image/jpeg', 'image/png', 'image/webp', 'image/gif'];
  if (allowedTypes.includes(file.mimetype)) {
    cb(null, true);
  } else {
    cb(new Error('Only image files (JPG, PNG, WEBP, GIF) are allowed'), false);
  }
};

const upload = multer({
  storage,
  fileFilter,
  limits: { fileSize: 5 * 1024 * 1024 } 
});

// home
userRouter.get('/',isBlocked, userController.loadHomePage)

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

// google auth 
userRouter.get('/auth/google', (req, res, next) => {
  const state = req.query.intent || 'login'
  passport.authenticate('google', {
    scope: ['profile', 'email'],
    state
  })(req, res, next)
})


userRouter.get('/auth/google/callback',
  (req, res, next) => {
    passport.authenticate('google', (err, user, info) => {
      if (err) return next(err)

      if (!user) {
        // info.message is set in passport.js via done(null, false, { message })
        if (info?.message === 'already_registered') {
          return res.redirect('/login?googleError=already_registered')
        }
        return res.redirect('/login')
      }

      req.session.user = {
        _id:          user._id,
        email:        user.email,
        isBlocked:    user.isBlocked,
        fullName:     user.fullName,
        profilePhoto: user.profilePhoto,
      }
      res.redirect('/')
    })(req, res, next)
  }
)

// forgot password
userRouter.get('/forgot-password', userController.LoadforgotPassword)
userRouter.post('/api/forgot-password', userController.forgotPassword)
userRouter.get('/reset-password', userController.showResetPage)
userRouter.post('/api/auth/reset-password', userController.resetPassword)
userRouter.get('/otp-forgot-password', userController.showForgotOtpPage)


//profile
userRouter.get('/profile',isBlocked,profileController.loadProfile)
userRouter.post('/profile',isBlocked,upload.single('profilePhoto'),profileController.updateProfile)
userRouter.post('/profile/changePassword',isBlocked,profileController.changePassword)
userRouter.post('/profile/request-email-change',isBlocked,profileController.requestEmailChange);



// address
userRouter.get('/addresses', isBlocked,addressController.loadAddresses)
userRouter.get('/addresses/add', isBlocked,addressController.loadAddAddress)
userRouter.post('/addresses/add',isBlocked,addressController.addAddress)
userRouter.get('/addresses/edit/:id',isBlocked,addressController.loadEditAddress)
userRouter.put('/addresses/edit/:id',isBlocked, addressController.editAddress)
userRouter.delete('/addresses/delete/:id',isBlocked, addressController.deleteAddress)
userRouter.patch('/addresses/default/:id', isBlocked,addressController.setDefaultAddress)


//products
userRouter.get('/products',isBlocked,productController.loadProductListing)

//productdetail
userRouter.get('/products/:id',isBlocked,productDetailController.loadProductDetail)
userRouter.post('/products/:id/review',isBlocked,productDetailController.submitReview)
userRouter.delete('/products/:id/review',isBlocked,productDetailController.deleteReview)

//cart 
userRouter.get('/cart',isBlocked,cartController.loadCart)
userRouter.post('/cart/add',isBlocked,cartController.addToCart)
userRouter.post('/cart/update',isBlocked,cartController.updateCartItem)
userRouter.delete('/cart/remove/:itemId',isBlocked,cartController.removeFromCart)
userRouter.post('/wishlist/toggle', isBlocked,cartController.toggleWishlist)

userRouter.get('/wishlist',isBlocked,cartController.loadWishlist)


userRouter.post('/cart/apply-coupon',isBlocked,cartController.applyCoupon)
userRouter.delete('/cart/remove-coupon',isBlocked,cartController.removeCoupon)


 
// checkout
userRouter.get('/checkout',           isBlocked, checkoutController.loadCheckout)
userRouter.post('/checkout/place-order', isBlocked, checkoutController.placeOrder)
 userRouter.post('/checkout/create-razorpay-order', isBlocked, checkoutController.createRazorpayOrder)  
userRouter.post('/checkout/verify-payment',      isBlocked, checkoutController.verifyPayment)          

// order failure page
userRouter.get('/order-failure',                 isBlocked, (req, res) => {                           
  res.render('user/orderFailure', { reason: req.query.reason || null })
})

// order success
userRouter.get('/order-success/:orderId', isBlocked, checkoutController.loadOrderSuccess)
 
// order history & detail
userRouter.get('/orders',             isBlocked, checkoutController.loadOrders)
userRouter.get('/orders/:orderId',    isBlocked, checkoutController.loadOrderDetail)
 
// cancel order
userRouter.post('/orders/:orderId/cancel', isBlocked, checkoutController.cancelOrder)

//return order
userRouter.post('/orders/:orderId/return', isBlocked, checkoutController.returnOrder) 

//invoice
userRouter.get('/orders/:orderId/invoice', isBlocked, checkoutController.downloadInvoice)

//wallet
userRouter.get('/wallet', isBlocked, walletController.getWallet)

//coupon
userRouter.get('/coupons',isBlocked , profileController.loadCoupons)

userRouter.get('/contact', isBlocked , contactController.loadContact)


userRouter.get('/referral-info',  isBlocked, profileController.getReferralInfo)
 
userRouter.get('/validate-referral', async (req, res) => {
  const { code } = req.query
  if (!code) return res.json({ valid: false })
  const user = await User.findOne({ referralCode: code.trim().toUpperCase() })
  return res.json({ valid: !!user })
})

// logout
userRouter.get('/logout', isLoggedIn, userController.logout)

module.exports = userRouter