const express            = require('express')
const adminRouter        = express.Router()
const adminController    = require('../controllers/admin/adminController')
const customerController = require('../controllers/admin/customerController')
const categoryController = require('../controllers/admin/categoryController')
const productController = require('../controllers/admin/productController')
const adminAuth          = require('../middleware/adminAuth')
const { uploadProductImages } = require('../config/cloudinary');

adminRouter.get('/login',     adminController.loadAdminLogin)
adminRouter.post('/login',    adminController.adminLogin)
adminRouter.get('/dashboard', adminAuth, adminController.loadDashboard)



adminRouter.get('/customers',adminAuth, customerController.loadCustomer)
adminRouter.post('/customers/block/:id',   adminAuth, customerController.blockUser)
adminRouter.post('/customers/unblock/:id', adminAuth, customerController.unblockUser)
adminRouter.get('/customers/edit/:id',adminAuth,customerController.loadEditCustosmer)
adminRouter.put('/customers/edit/:id',adminAuth,customerController.editCustomer)
adminRouter.delete('/customers/delete/:id',adminAuth,customerController.deleteCustomer)



adminRouter.get('/category',adminAuth,categoryController.loadCategory)
adminRouter.get('/addCategory',adminAuth,categoryController.loadAddCategory)
adminRouter.get('/editCategory/:id',adminAuth,categoryController.loadEditCategory)
adminRouter.post('/addCategory',adminAuth,categoryController.addCategory)
adminRouter.put('/editCategory/:id',adminAuth,categoryController.editCategory)
adminRouter.delete('/deleteCategory/:id',adminAuth,categoryController.deleteCategory)
adminRouter.post('/toggleCategoryListing/:id',adminAuth,categoryController.toggleCategoryListing)


adminRouter.get('/products',adminAuth,productController.loadProducts)
adminRouter.get('/addProduct',adminAuth,productController.loadAddProduct)
adminRouter.post('/addProduct',adminAuth,productController.addProduct)
adminRouter.get('/editProduct/:id',adminAuth,productController.loadEditProduct)
adminRouter.put('/editProduct/:id',adminAuth,productController.editProduct)
adminRouter.post('/toggleProductListing/:id',adminAuth,productController.toggleProductListing)
adminRouter.delete('/deleteProduct/:id',adminAuth,productController.deleteProduct)

adminRouter.get('/logout',adminAuth,adminController.adminLogout)

module.exports = adminRouter;