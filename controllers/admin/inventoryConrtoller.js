const Product = require('../../models/user/productModel')
const Category = require('../../models/user/categoryModel')

const LOW_STOCK_THRESHOLD = 10
const LIMIT = 10


const loadInventory = async (req, res) => {
  try {
    const page       = Math.max(1, parseInt(req.query.page) || 1)
    const search     = (req.query.search || '').trim()
    const categoryId = req.query.category || ''
    const stockFilter = req.query.stock   || '' 

    const query = { isDeleted: false }
    if (search)     query.name       = { $regex: search, $options: 'i' }
    if (categoryId) query.categoryId = categoryId

    
    let products = await Product.find(query)
      .populate('categoryId', 'name')
      .sort({ createdAt: -1 })
      .lean()

    
    products = products.map(p => {
      const totalStock = p.variants.reduce((s, v) => s + (v.stock || 0), 0)
      const lowVariants = p.variants.filter(v => v.stock > 0 && v.stock <= LOW_STOCK_THRESHOLD).length
      const outVariants = p.variants.filter(v => v.stock === 0).length
      return { ...p, totalStock, lowVariants, outVariants }
    })

    
    if (stockFilter === 'low') {
      products = products.filter(p => p.lowVariants > 0)
    } else if (stockFilter === 'out') {
      products = products.filter(p => p.outVariants > 0)
    } else if (stockFilter === 'ok') {
      products = products.filter(p => p.totalStock > LOW_STOCK_THRESHOLD && p.outVariants === 0)
    }

    const total      = products.length
    const totalPages = Math.ceil(total / LIMIT)
    const paginated  = products.slice((page - 1) * LIMIT, page * LIMIT)

    
    const lowStockCount = products.filter(p => p.lowVariants > 0 || p.outVariants > 0).length

    const categories = await Category.find({ isDeleted: false }).lean()

    res.render('admin/inventory', {
      products: paginated,
      currentPage: page,
      totalPages,
      total,
      search,
      categoryId,
      stockFilter,
      lowStockCount,
      LOW_STOCK_THRESHOLD,
      categories,
      user: req.session.admin || null,
    })
  } catch (err) {
    console.error('loadInventory error:', err)
    res.status(500).render('error', { message: 'Could not load inventory.' })
  }
}



const updateVariantStock = async (req, res) => {
  try {
    const { productId, shade } = req.params
    const { stock, note } = req.body
    const newStock = parseInt(stock)

    if (isNaN(newStock) || newStock < 0) {
      return res.status(400).json({ success: false, message: 'Invalid stock value.' })
    }

    const product = await Product.findById(productId)
    if (!product) return res.status(404).json({ success: false, message: 'Product not found.' })

    const variant = product.variants.find(v => v.shade === shade)
    if (!variant) return res.status(404).json({ success: false, message: 'Variant not found.' })

    const oldStock = variant.stock
    variant.stock  = newStock

    
    if (!product.stockHistory) product.stockHistory = []
    product.stockHistory.push({
      shade,
      oldStock,
      newStock,
      note:      note || '',
      updatedBy: 'admin',
      updatedAt: new Date(),
    })

    await product.save()
    return res.json({ success: true, message: 'Stock updated.', newStock })
  } catch (err) {
    console.error('updateVariantStock error:', err)
    return res.status(500).json({ success: false, message: 'Something went wrong.' })
  }
}



const getStockHistory = async (req, res) => {
  try {
    const product = await Product.findById(req.params.productId).lean()
    if (!product) return res.status(404).json({ success: false, message: 'Product not found.' })

    const history = (product.stockHistory || [])
      .sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt))
      .slice(0, 50)  

    return res.json({ success: true, history, productName: product.name })
  } catch (err) {
    console.error('getStockHistory error:', err)
    return res.status(500).json({ success: false, message: 'Something went wrong.' })
  }
}

module.exports = {
   loadInventory,
    updateVariantStock,
     getStockHistory 
    }