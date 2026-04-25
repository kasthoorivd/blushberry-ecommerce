const adminAuth = (req, res, next) => {
  if (req.session && req.session.admin) {
    return next()
  }
  res.redirect('/admin/login')
}

module.exports = adminAuth
