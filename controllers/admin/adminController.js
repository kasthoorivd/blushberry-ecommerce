const User = require('../../models/user/userModel')
const bcrypt = require('bcrypt')

const loadAdminLogin = (req, res) => {
    res.render('admin/adminLogin.ejs')
}

const loadDashboard = async (req, res) => {
   try {
    const user = req.session.user || null   
    res.render('admin/dashboard', { user })  
  } catch (error) {
    console.log(error)
  }
}

const adminLogin = async (req, res) => {
    try {
        const { email, password } = req.body;

        const user = await User.findOne({ email });

        if (!user) {
            return res.json({ success: false, message: "Admin not found" })
        }

        if (!user.isAdmin) {
            return res.json({ success: false, message: 'You are not authorised as Admin' })
        }

        const isMatch = await bcrypt.compare(password, user.password)

        if (!isMatch) {
            return res.json({ success: false, message: 'Incorrect password' })
        }

        req.session.admin = user._id;
        req.session.save((err) => {
            if (err) {
                console.log("Session save error", err);
                return res.status(500).json({ success: false, message: "Session error" })
            }
            return res.json({
                success: true,
                redirectUrl: '/admin/dashboard'
            })
        })

    } catch (error) {
        console.log("Admin login error", error)
        res.status(500).json({ success: false, message: "Server error" })
    }
}

const adminLogout = async (req, res) => {
    try {
        req.session.destroy((err) => {
            if (err) {
                console.error("adminLogout error:", err);
                return res.redirect("/admin/dashboard");
            }
            res.clearCookie("connect.sid");
            return res.redirect("/admin/login");
        });
    } catch (error) {
        console.error("adminLogout error:", error);
        return res.redirect("/admin/dashboard");
    }
};



module.exports = {
    loadAdminLogin,
    loadDashboard,
    adminLogin,
    adminLogout
}