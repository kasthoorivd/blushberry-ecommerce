const passport = require('passport')
const express = require('express')
const multer = require('multer')
const path = require('path')
const userRouter = express.Router()

const userController = require('../controllers/user/userAuthController')
const profileController = require('../controllers/user/profileController')
const addressController = require('../controllers/user/addressController')
const productController = require('../controllers/user/productController')
const productDetailController = require('../controllers/user/productDetailController')
const cartController = require('../controllers/user/cartController')
const checkoutController = require('../controllers/user/checkoutController')
const {isLoggedIn, isLoggedOut,isBlocked} = require('../middleware/authMiddleware')



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
userRouter.get('/auth/google', passport.authenticate('google', { scope: ['profile', 'email'] }))


userRouter.get('/auth/google/callback',
  passport.authenticate('google', { failureRedirect: '/login' }),
  (req, res) => {

    req.session.user = req.user; 
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
userRouter.get('/profile',isBlocked,profileController.loadProfile)
userRouter.post('/profile',isBlocked,upload.single('profilePhoto'),profileController.updateProfile)
userRouter.put('/profile/changepassword',isBlocked,profileController.changePassword)
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

// logout
userRouter.get('/logout', isLoggedIn, userController.logout)

module.exports = userRouter