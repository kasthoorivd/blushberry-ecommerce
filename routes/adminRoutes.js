const express = require('express')
const adminRouter = express.Router()
const adminController = require('../controllers/admin/adminController')
const adminAuth = require('../middleware/adminAuth')


adminRouter.get('/login',adminController.loadAdminLogin)
adminRouter.get('/dashboard',adminAuth,adminController.loadDashboard)
adminRouter.post('/admin-login',adminController.adminLogin)
module.exports = adminRouter;