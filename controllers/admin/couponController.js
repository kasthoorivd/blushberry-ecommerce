const Coupon = require('../../models/user/couponModel')

const loadCoupons = async (req, res) => {
  try {
    const page  = parseInt(req.query.page) || 1
    const LIMIT = 3

    const totalCoupons = await Coupon.countDocuments()
    const coupons = await Coupon.find()
      .sort({ createdAt: -1 })
      .skip((page - 1) * LIMIT)
      .limit(LIMIT)
      .lean()

    res.render('admin/coupons', {
      coupons,
      currentPage:  page,
      totalPages:   Math.ceil(totalCoupons / LIMIT)
    })
  } catch (error) {
    console.error(error)
    res.status(500).send('Error loading coupons')
  }
}

const createCoupon = async (req, res) => {
  try {
    const {
      code, discountType, discountAmount,
      maxDiscount, minOrderAmount, maxUses, expiresAt
    } = req.body

    const existing = await Coupon.findOne({ code: code.toUpperCase() })
    if (existing) {
      return res.json({ success: false, message: 'Coupon code already exists.' })
    }

    const type     = discountType === 'percentage' ? 'percentage' : 'flat'
    const amount   = parseFloat(discountAmount)
    const minOrder = parseFloat(minOrderAmount) || 0

    if (!amount || amount <= 0) {
      return res.json({ success: false, message: 'Discount value must be greater than 0.' })
    }

    if (type === 'percentage') {
      if (amount > 90) {
        return res.json({ success: false, message: 'Percentage discount cannot exceed 90%.' })
      }
    } else {
      // flat: min order must be at least 2× discount
      if (minOrder > 0 && minOrder < amount * 2) {
        return res.json({
          success: false,
          message: `Min Order Amount must be at least 2× the discount (₹${amount * 2} or more).`
        })
      }
    }

    await Coupon.create({
      code:            code.toUpperCase(),
      discountType:    type,
      discountAmount:  amount,
      maxDiscount:     type === 'percentage' ? (parseFloat(maxDiscount) || null) : null,
      minOrderAmount:  minOrder,
      maxUses:         maxUses  || null,
      expiresAt:       expiresAt || null
    })

    return res.json({ success: true, message: 'Coupon created successfully.' })
  } catch (error) {
    console.error(error)
    res.status(500).json({ success: false, message: 'Server error.' })
  }
}

const toggleCoupon = async (req, res) => {
  try {
    const coupon = await Coupon.findById(req.params.id)
    coupon.isActive = !coupon.isActive
    await coupon.save()
    res.json({ success: true, isActive: coupon.isActive })
  } catch (error) {
    res.status(500).json({ success: false })
  }
}

const deleteCoupon = async (req, res) => {
  try {
    await Coupon.findByIdAndDelete(req.params.id)
    res.json({ success: true })
  } catch (error) {
    res.status(500).json({ success: false })
  }
}

module.exports = { loadCoupons, createCoupon, toggleCoupon, deleteCoupon }