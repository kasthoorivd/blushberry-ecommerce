const Product = require('../../models/user/productModel')
const Category = require('../../models/user/categoryModel')
const Offer = require('../../models/user/offerModel')
const Wishlist = require('../../models/user/wishlistModel')
const LIMIT = 6



// ─── helper: fetch all active offers and build lookup maps ───────────────────
async function getActiveOfferMaps() {
  const now = new Date()
  const activeOffers = await Offer.find({
    isActive:  true,
    startDate: { $lte: now },
    endDate:   { $gte: now }
  }).lean()

  const productOfferMap  = {}  
  const categoryOfferMap = {}   

  for (const o of activeOffers) {
    const key = String(o.targetId)
    if (o.type === 'product') {
      if (!productOfferMap[key] || o.discountPercent > productOfferMap[key]) {
        productOfferMap[key] = o.discountPercent
      }
    } else if (o.type === 'category') {
      if (!categoryOfferMap[key] || o.discountPercent > categoryOfferMap[key]) {
        categoryOfferMap[key] = o.discountPercent
      }
    }
  }

  return { productOfferMap, categoryOfferMap }
}


function applyBestOffer(p, productOfferMap, categoryOfferMap) {
  const productOffer  = productOfferMap[String(p._id)]         || 0
  const categoryOffer = categoryOfferMap[String(p.categoryId)] || 0
  const bestOffer     = Math.max(productOffer, categoryOffer)

  const basePrices  = p.variants.map(v => v.varientPrice)
  const baseMin     = Math.min(...basePrices)

  p.originalPrice  = baseMin
  p.displayOffer   = bestOffer

  if (bestOffer > 0) {
    p.displayPrice = parseFloat((baseMin * (1 - bestOffer / 100)).toFixed(2))
  } else {
    
    const effectivePrices = p.variants.map(v => v.salePrice > 0 ? v.salePrice : v.varientPrice)
    p.displayPrice = Math.min(...effectivePrices)
  }

  p.inStock = p.variants.some(v => v.stock > 0)
  return p
}

const loadProductListing = async (req, res) => {
  try {
    const page        = Math.max(1, parseInt(req.query.page) || 1)
    const searchQuery = (req.query.search || '').trim()
    const categoryId  = req.query.category || ''
    const shade       = req.query.shade    || ''
    const sortBy      = req.query.sort     || 'newest'
    const minPrice    = parseFloat(req.query.minPrice) || 0
    const maxPrice    = parseFloat(req.query.maxPrice) || 999999

    const filter = { isDeleted: false, isListed: true }

    if (searchQuery) filter.name = { $regex: searchQuery, $options: 'i' }
    if (categoryId)  filter.categoryId = categoryId
    if (shade)       filter['variants.shade'] = { $regex: shade, $options: 'i' }

    if (minPrice > 0 || maxPrice < 999999) {
      filter.$or = [
        { 'variants.salePrice':    { $gte: minPrice, $lte: maxPrice } },
        { 'variants.varientPrice': { $gte: minPrice, $lte: maxPrice } }
      ]
    }

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

    // ── fetch all active offers once ──────────────────────────────────────
    const now = new Date()
    const activeOffers = await Offer.find({
      isActive:  true,
      startDate: { $lte: now },
      endDate:   { $gte: now }
    }).lean()

    const productOfferMap  = {}
    const categoryOfferMap = {}

    for (const o of activeOffers) {
      const key = String(o.targetId)
      if (o.type === 'product') {
        if (!productOfferMap[key] || o.discountPercent > productOfferMap[key])
          productOfferMap[key] = o.discountPercent
      } else if (o.type === 'category') {
        if (!categoryOfferMap[key] || o.discountPercent > categoryOfferMap[key])
          categoryOfferMap[key] = o.discountPercent
      }
    }
    // ─────────────────────────────────────────────────────────────────────

    products.forEach(p => {
      const productOffer  = productOfferMap[String(p._id)]              || 0
      const categoryOffer = categoryOfferMap[String(p.categoryId?._id || p.categoryId)] || 0
      const bestOffer     = Math.max(productOffer, categoryOffer)

      const origPrices    = p.variants.map(v => v.varientPrice)
      p.originalPrice     = Math.min(...origPrices)
      p.displayOffer      = bestOffer

      if (bestOffer > 0) {
        p.displayPrice = parseFloat((p.originalPrice * (1 - bestOffer / 100)).toFixed(2))
      } else {
        const effectivePrices = p.variants.map(v => v.salePrice > 0 ? v.salePrice : v.varientPrice)
        p.displayPrice = Math.min(...effectivePrices)
      }

      p.inStock = p.variants.some(v => v.stock > 0)
    })

    const categories      = await Category.find({ isDeleted: false }).lean()
    const paginationPages = buildPaginationWindow(safePage, totalPages)

    const queryParams = {
      search:   searchQuery,
      category: categoryId,
      shade,
      sort:     sortBy,
      minPrice: minPrice || '',
      maxPrice: maxPrice < 999999 ? maxPrice : ''
    }

    let wishlistIds = []
    if(req.session.user?._id){
      const wishlist = await Wishlist.findOne({userId: req.session.user._id})
      wishlistIds = wishlist ? wishlist.products.map(id => String(id)) : []
    }

    res.render('user/productListing', {
      products,
      wishlistIds ,
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