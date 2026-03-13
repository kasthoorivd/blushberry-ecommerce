const express            = require('express')
const adminRouter        = express.Router()
const adminController    = require('../controllers/admin/adminController')
const customerController = require('../controllers/admin/customerController')
const adminAuth          = require('../middleware/adminAuth')


adminRouter.get('/login',     adminController.loadAdminLogin)
adminRouter.post('/login',    adminController.adminLogin)
adminRouter.get('/dashboard', adminAuth, adminController.loadDashboard)



adminRouter.get('/customers',adminAuth, customerController.loadCustomer)
adminRouter.post('/customers/block/:id',   adminAuth, customerController.blockUser)
adminRouter.post('/customers/unblock/:id', adminAuth, customerController.unblockUser)
adminRouter.get('/customers/edit/:id',adminAuth,customerController.loadEditCustosmer)
adminRouter.put('/customers/edit/:id',adminAuth,customerController.editCustomer)
adminRouter.delete('/customers/delete/:id',adminAuth,customerController.deleteCustomer)

adminRouter.get('/logout',adminAuth,adminController.adminLogout)

module.exports = adminRouter;