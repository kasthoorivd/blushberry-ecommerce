const passport = require('passport')
const express = require('express')

const userRouter = express.Router()

const userController = require('../controllers/user/userAuthController')
const {isLoggedIn,isLoggedOut} = require('../middleware/authMiddleware')


userRouter.get('/',isLoggedIn,userController.loadHomePage)

userRouter.get('/signup', isLoggedOut,userController.loadSignUp)
userRouter.get('/login',isLoggedOut,userController.loadLogin) 

userRouter.post('/login',userController.login)
userRouter.post('/signup',userController.signup) 

userRouter.get('/otp',userController.loadOtpPage)
userRouter.post('/verifyOtp',userController.verifyOtp)
userRouter.post('/resendOtp',userController.resendOtp)


userRouter.get('/auth/google',passport.authenticate('google',{scope:['profile','email']}))

userRouter.get('/auth/google/callback',passport.authenticate('google',{
    failureRedirect:'/login'
}),
(req,res) => {
    res.redirect('/')
}

)


userRouter.get('/forgot-password',userController.forgotPassword)

// userRouter.post('/forgot-password',userController.verifyOtp)
userRouter.get('/logout',isLoggedIn,userController.logout)

module.exports = userRouter