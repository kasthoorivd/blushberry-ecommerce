// const isLoggedIn = (req, res, next) => {
//     if (req.isAuthenticated()) {
//         return next()
//     }
//     res.redirect('/login')
// }

// const isLoggedOut = (req, res, next) => {
//     if (!req.isAuthenticated()) {
//         return next()
//     }
//     res.redirect('/')
// }

// module.exports = { isLoggedIn, isLoggedOut }

const isLoggedIn = (req,res,next)=>{
    if(req.session && req.session.user ){
        return next()
    }
    return res.redirect('/login')
}

const isLoggedOut = (req,res,next) =>{
    if(!req.session || !req.session.user  ){
        return next()
    }

    return res.redirect('/');
}

module.exports = {isLoggedIn,isLoggedOut}

// const User=require('../models/user/userModel')
// const checkSession=(req,res,next)=>{
//     if(req.session.user){
//         return next()
//     }
//     res.render('/login',{error:"please login or signup"}); 
// }

// // const isBlocked=(req,res,next)=>{
// //     if(req.session.user){
// //         User.findById(req.session.user._id)
// //         .then(data => {
// //             if (data && data.isBlocked) {
// //                 return res.render('user/login',{error:"user is blocked by admin"});
// //             } else{
// //                 next()
// //             }
// //         }
// //         )
        
// //     }else{
// //         next()
// //     }
    
// // }

// const isLogin=(req,res,next)=>{
//     if(req.session.user){
//         return res.redirect('/');
//     }
//     next()
// }
// module.exports={
//     checkSession,
//     // isBlocked,
//     isLogin
// }