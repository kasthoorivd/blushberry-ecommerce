const User = require('../../models/user/userModel')
const bcrypt = require('bcrypt')

const loadAdminLogin = (req, res) => {
    res.render('admin/adminLogin.ejs')
}

const adminLogin = async (req, res) => {
    try {
        const { email, password } = req.body

        const user = await User.findOne({ email })

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

    
        req.session.admin = { _id: user._id, email: user.email }

        req.session.save((err) => {
            if (err) {
                console.error("Session save error:", err)
                return res.status(500).json({ success: false, message: "Session error" })
            }
            return res.json({ success: true, redirectUrl: '/admin/dashboard' })
        })

    } catch (error) {
        console.error("Admin login error:", error)
        res.status(500).json({ success: false, message: "Server error" })
    }
}


const adminLogout = (req, res) => {
    try {
        delete req.session.admin
        res.clearCookie("admin.sid")
        return res.redirect("/admin/login")
    } catch (error) {
        console.error("adminLogout error:", error)
        return res.redirect("/admin/dashboard")
    }
}

module.exports = {
    loadAdminLogin,
    adminLogin,
    adminLogout
}