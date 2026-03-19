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
userRouter.get('/profile',profileController.loadProfile)
userRouter.post('/profile',upload.single('profilePhoto'),profileController.updateProfile)
userRouter.put('/profile/changepassword',profileController.changePassword)
userRouter.post('/profile/request-email-change',profileController.requestEmailChange);



// address
userRouter.get('/addresses', addressController.loadAddresses)
userRouter.get('/addresses/add', addressController.loadAddAddress)
userRouter.post('/addresses/add',addressController.addAddress)
userRouter.get('/addresses/edit/:id',addressController.loadEditAddress)
userRouter.put('/addresses/edit/:id', addressController.editAddress)
userRouter.delete('/addresses/delete/:id', addressController.deleteAddress)
userRouter.patch('/addresses/default/:id', addressController.setDefaultAddress)


//products
userRouter.get('/products',productController.loadProductListing)

//productdetail
userRouter.get('/products/:id',productDetailController.loadProductDetail)
userRouter.post('/products/:id/review',productDetailController.submitReview)
userRouter.delete('/products/:id/review',productDetailController.deleteReview)

//cart 
userRouter.get('/cart',cartController.loadCart)
userRouter.post('/cart/add',cartController.addToCart)
userRouter.post('/cart/update',cartController.updateCartItem)
userRouter.delete('/cart/remove/:itemId',cartController.removeFromCart)
userRouter.post('/wishlist/toggle', cartController.toggleWishlist)

userRouter.get('/wishlist',cartController.loadWishlist)


userRouter.post('/cart/apply-coupon',cartController.applyCoupon)
userRouter.delete('/cart/remove-coupon',cartController.removeCoupon)
// logout
userRouter.get('/logout', isLoggedIn, userController.logout)

module.exports = userRouter