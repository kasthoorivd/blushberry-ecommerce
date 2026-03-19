const Product = require('../../models/user/productModel')
const Category = require('../../models/user/categoryModel')


const LIMIT = 6


const loadProductListing = async (req, res) => {
  try {
    const page        = Math.max(1, parseInt(req.query.page) || 1)
    const searchQuery = (req.query.search || '').trim()
    const categoryId  = req.query.category || ''
    const shade       = req.query.shade    || ''
    const sortBy      = req.query.sort     || 'newest'
    const minPrice    = parseFloat(req.query.minPrice) || 0
    const maxPrice    = parseFloat(req.query.maxPrice) || 999999
 
    // ── Base filter: only listed, non-deleted products ──
    const filter = { isDeleted: false, isListed: true }
 
    if (searchQuery) {
      filter.name = { $regex: searchQuery, $options: 'i' }
    }
 
    if (categoryId) {
      filter.categoryId = categoryId
    }
 
    if (shade) {
      filter['variants.shade'] = { $regex: shade, $options: 'i' }
    }
 
    if (minPrice > 0 || maxPrice < 999999) {
      filter.$or = [
        { 'variants.salePrice':    { $gte: minPrice, $lte: maxPrice } },
        { 'variants.varientPrice': { $gte: minPrice, $lte: maxPrice } }
      ]
    }
 
    // ── Sort ──
    const sortMap = {
      newest:    { createdAt: -1 },
      oldest:    { createdAt:  1 },
      priceAsc:  { 'variants.varientPrice':  1 },
      priceDesc: { 'variants.varientPrice': -1 },
      nameAsc:   { name:  1 },
      nameDesc:  { name: -1 },
    }
    const sortOption = sortMap[sortBy] || sortMap.newest
 
    const totalProducts = await Product.countDocuments(filter)
    const totalPages    = Math.ceil(totalProducts / LIMIT) || 1
    const safePage      = Math.min(page, totalPages)
 
    const products = await Product.find(filter)
      .populate('categoryId', 'name')
      .sort(sortOption)
      .skip((safePage - 1) * LIMIT)
      .limit(LIMIT)
      .lean()
 
    products.forEach(p => {
      const prices = p.variants.map(v => v.salePrice > 0 ? v.salePrice : v.varientPrice)
      p.displayPrice    = Math.min(...prices)
      p.displayMaxPrice = Math.max(...prices)
      p.displayOffer    = p.offer || 0
      const origPrices  = p.variants.map(v => v.varientPrice)
      p.originalPrice   = Math.min(...origPrices)
      p.inStock         = p.variants.some(v => v.stock > 0)
    })
 
    const categories = await Category.find({ isDeleted: false }).lean()
 
    const paginationPages = buildPaginationWindow(safePage, totalPages)
 
    const queryParams = {
      search:   searchQuery,
      category: categoryId,
      shade,
      sort:     sortBy,
      minPrice: minPrice || '',
      maxPrice: maxPrice < 999999 ? maxPrice : ''
    }
 
    res.render('user/productListing', {
      products,
      categories,
      currentPage:   safePage,
      totalPages,
      totalProducts,
      limit:         LIMIT,
      searchQuery,
      categoryId,
      shade,
      sortBy,
      minPrice:  minPrice || '',
      maxPrice:  maxPrice < 999999 ? maxPrice : '',
      paginationPages,
      queryParams,
      user: req.session.user || null
    })
 
  } catch (err) {
    console.error('loadProductListing error:', err)
    res.status(500).render('error', { message: 'Could not load products.' })
  }
}
 
function buildPaginationWindow(current, total, delta = 2) {
  const pages = []
  const left  = Math.max(1, current - delta)
  const right = Math.min(total, current + delta)
 
  if (left > 1)      { pages.push(1);     if (left > 2) pages.push('...') }
  for (let i = left; i <= right; i++) pages.push(i)
  if (right < total) { if (right < total - 1) pages.push('...'); pages.push(total) }
 
  return pages
}

module.exports = {
    loadProductListing
}