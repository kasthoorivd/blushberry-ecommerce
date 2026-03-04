// middleware/adminAuth.js

const adminAuth = (req, res, next) => {
  try {
    // Check if admin session exists
    if (!req.session.admin) {
      return res.redirect('/admin/login');
    }

    // Optional: Extra safety check
    if (!req.session.admin.isAdmin) {
      return res.redirect('/admin/login');
    }

    // If everything is fine, continue
    next();

  } catch (error) {
    console.error("Admin Auth Middleware Error:", error);
    return res.redirect('/admin/login');
  }
};

module.exports = adminAuth;