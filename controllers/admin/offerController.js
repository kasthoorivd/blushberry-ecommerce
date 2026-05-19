const Offer = require('../../models/user/offerModel')
const Product = require('../../models/user/productModel')
const Category = require('../../models/user/categoryModel')
const { HttpStatus } = require('../../utils/statusCode')

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

    const success = req.query.success || null
    const error = req.query.error || null

    res.render('admin/offers', {
      offers: annotated,
      success,
      error: null,
      formData: {}
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
    console.error('addOffer error:', err)
    res.redirect('/admin/offers')
  }
}

// Toggle active/inactive
const toggleOffer = async (req, res) => {
  try {
    const offer = await Offer.findById(req.params.id)
    if (!offer) {
      return res.status(HttpStatus.NOT_FOUND).json({ success: false, message: 'Offer not found.' })
    }
    offer.isActive = !offer.isActive
    await offer.save()
    res.redirect('/admin/offers')
  } catch (err) {
    console.error('toggleOffer error:', err)
    res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({ success: false, message: 'Toggle failed.' })
  }
}

// Delete offer
const deleteOffer = async (req, res) => {
  try {
    await Offer.findByIdAndDelete(req.params.id)
    res.redirect('/admin/offers')
  } catch (err) {
    console.error('deleteOffer error:', err)
    res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({ success: false, message: 'Delete failed.' })
  }
}

// Load add-offer form (needs product + category lists)
const loadAddOffer = async (req, res) => {
  try {
    const [products, categories] = await Promise.all([
      Product.find({ isDeleted: false, isListed: true }).select('name').lean(),
      Category.find({ isDeleted: false }).select('name').lean()
    ])
    res.render('admin/addOffer', { products, categories })
  } catch (err) {
    console.error('loadAddOffer error:', err)
    res.status(HttpStatus.INTERNAL_SERVER_ERROR).render('error', { message: 'Could not load form.' })
  }
}

module.exports = { loadOffers, addOffer, toggleOffer, deleteOffer, loadAddOffer }