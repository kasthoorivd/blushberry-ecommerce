const Offer = require('../models/user/offerModel')

async function getEffectivePrice(product, shade = null) {
  const now = new Date()

  // Pick the right variant — with safe fallback
  let variant = product.variants?.[0]
  if (shade && product.variants) {
    variant = product.variants.find(
      v => v.shade?.toLowerCase() === shade.toLowerCase()
    ) || product.variants[0]
  }

  // ── Safe fallback if no variant found ──
  if (!variant) {
    return {
      finalPrice:       0,
      bestDiscount:     0,
      originalPrice:    0,
      productOffer:     null,
      categoryOffer:    null,
      appliedOfferType: null
    }
  }

  const originalPrice = variant.varientPrice || 0
  const salePrice     = variant.salePrice > 0 ? variant.salePrice : null

  const [productOffer, categoryOffer] = await Promise.all([
    Offer.findOne({
      type:      'product',
      targetId:  product._id,
      isActive:  true,
      startDate: { $lte: now },
      endDate:   { $gte: now }
    }).lean(),
    Offer.findOne({
      type:      'category',
      targetId:  product.categoryId?._id || product.categoryId,
      isActive:  true,
      startDate: { $lte: now },
      endDate:   { $gte: now }
    }).lean()
  ])

  const productDiscount  = productOffer?.discountPercent  || 0
  const categoryDiscount = categoryOffer?.discountPercent || 0
  const bestOfferDiscount = Math.max(productDiscount, categoryDiscount)

  let finalPrice   = originalPrice
  let bestDiscount = 0

  if (bestOfferDiscount > 0) {
    const offerPrice = +(originalPrice * (1 - bestOfferDiscount / 100)).toFixed(2)
    if (!salePrice || offerPrice < salePrice) {
      finalPrice   = offerPrice
      bestDiscount = bestOfferDiscount
    } else {
      finalPrice   = salePrice
      bestDiscount = +((1 - salePrice / originalPrice) * 100).toFixed(1)
    }
  } else if (salePrice) {
    finalPrice   = salePrice
    bestDiscount = +((1 - salePrice / originalPrice) * 100).toFixed(1)
  }

  return {
    finalPrice,
    bestDiscount,
    originalPrice,  
    productOffer,
    categoryOffer,
    appliedOfferType: productDiscount >= categoryDiscount ? 'product' : 'category'
  }
}

module.exports = getEffectivePrice