

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

