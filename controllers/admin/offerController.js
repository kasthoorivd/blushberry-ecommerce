const Offer    = require('../../models/user/offerModel')
const Product  = require('../../models/user/productModel')
const Category = require('../../models/user/categoryModel')

// List all offers
const loadOffers = async (req, res) => {
  try {
    const now = new Date()
    const offers = await Offer.find()
      .sort({ createdAt: -1 })
      .populate('targetId', 'name')
      .lean()

    const annotated = offers.map(o => ({
      ...o,
      status: !o.isActive
        ? 'inactive'
        : o.endDate < now
          ? 'expired'
          : o.startDate > now
            ? 'upcoming'
            : 'active'
    }))

    const success = req.query.success || null  // ← these two were missing
    const error   = req.query.error   || null  // ← 

    res.render('admin/offers', {
         offers: annotated, 
         success, 
         error :null,
         formData:{}
        })  
  } catch (err) {
    console.error('loadOffers error:', err)
    res.redirect('/admin/dashboard')
  }
}

// Add offer
const addOffer = async (req, res) => {
  try {
    const { name, type, discountPercent, targetId, startDate, endDate } = req.body
    await Offer.create({ name, type, discountPercent, targetId, startDate, endDate })
    res.redirect('/admin/offers')
  } catch (err) {
    console.error(err)
    res.redirect('/admin/offers')
  }
}

// Toggle active/inactive
const toggleOffer = async (req, res) => {
  const offer = await Offer.findById(req.params.id)
  offer.isActive = !offer.isActive
  await offer.save()
  res.redirect('/admin/offers')
}

// Delete offer
const deleteOffer = async (req, res) => {
  await Offer.findByIdAndDelete(req.params.id)
  res.redirect('/admin/offers')
}

// Load add-offer form (needs product + category lists)
const loadAddOffer = async (req, res) => {
  const [products, categories] = await Promise.all([
    Product.find({ isDeleted: false, isListed: true }).select('name').lean(),
    Category.find({ isDeleted: false }).select('name').lean()
  ])
  res.render('admin/addOffer', { products, categories })
}

module.exports = { loadOffers, addOffer, toggleOffer, deleteOffer, loadAddOffer }