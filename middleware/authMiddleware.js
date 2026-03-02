const isLoggedIn = (req, res, next) => {
    if (req.isAuthenticated()) {
        return next()
    }
    res.redirect('/login')
}

const isLoggedOut = (req, res, next) => {
    if (!req.isAuthenticated()) {
        return next()
    }
    res.redirect('/')
}

module.exports = { isLoggedIn, isLoggedOut }