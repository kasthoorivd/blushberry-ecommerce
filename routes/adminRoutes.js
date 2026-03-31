const express            = require('express')
const adminRouter        = express.Router()
const adminController    = require('../controllers/admin/adminController')
const customerController = require('../controllers/admin/customerController')
const categoryController = require('../controllers/admin/categoryController')
const productController = require('../controllers/admin/productController')
const couponController = require('../controllers/admin/couponController')
const orderController = require('../controllers/admin/orderController')
const inventoryController = require('../controllers/admin/inventoryConrtoller')
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


adminRouter.get('/coupons',couponController.loadCoupons)
adminRouter.post('/coupons/create',couponController.createCoupon)
adminRouter.patch('/coupons/:id/toggle',couponController.toggleCoupon)
adminRouter.delete('/coupons/:id',couponController.deleteCoupon)


// orders
adminRouter.get('/orders',           adminAuth, orderController.loadOrders)
adminRouter.get('/orders/:id',       adminAuth, orderController.loadOrderDetail)
adminRouter.patch('/orders/:id/status', adminAuth, orderController.updateOrderStatus)
adminRouter.patch('/orders/:id/cancel',             adminAuth, orderController.cancelOrder)
adminRouter.patch('/orders/:id/items/:itemId/cancel', adminAuth, orderController.cancelItem)
adminRouter.patch('/orders/:id/return-status',adminAuth, orderController.updateReturnStatus)

// inventory
adminRouter.get('/inventory',        adminAuth, inventoryController.loadInventory)
adminRouter.patch('/inventory/:productId/variant/:shade/stock', adminAuth, inventoryController.updateVariantStock)
adminRouter.get('/inventory/:productId/history', adminAuth, inventoryController.getStockHistory)



adminRouter.get('/logout',adminAuth,adminController.adminLogout)

module.exports = adminRouter;