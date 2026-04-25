const User = require('../models/user/userModel')

const isLoggedIn = (req, res, next) => {
  if (req.session && req.session.user) return next()
  return res.redirect('/login')
}

const isLoggedOut = (req, res, next) => {
  if (!req.session || !req.session.user) return next()
  return res.redirect('/')
}

const isBlocked = async (req, res, next) => {
  try {
    if (!req.session || !req.session.user) return next()

    const user = await User.findById(req.session.user._id)
    if (!user || user.isBlocked) {
      delete req.session.user          // only removes user key, session still exists
      return res.redirect('/login?blocked=true')
    }
    next()
  } catch (err) {
    console.error('isBlocked middleware error:', err)
    next(err)
  }
}

module.exports = { isLoggedIn, isLoggedOut, isBlocked }