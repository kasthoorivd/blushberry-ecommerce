const Cart = require('../../models/user/cartModel')
const Wishlist = require('../../models/user/wishlistModel')
const Product = require('../../models/user/productModel')
const Coupon = require('../../models/user/couponModel')


const MAX_QTY_PER_ITEM = 5

function getVariantPrice(variant) {
  return variant.salePrice > 0 ? variant.salePrice : variant.varientPrice
}

function findVariant(product, shade) {
  if (!shade) return product.variants[0]
  return product.variants.find(
    v => v.shade.toLowerCase() === shade.toLowerCase()
  ) || product.variants[0]
}

const loadCart = async (req, res) => {
  try {
    const userId = req.session.user?._id
    if (!userId) return res.redirect('/login')

    const cart = await Cart.findOne({ userId })
      .populate({
        path: 'items.productId',
        select: 'name images variants isListed isDeleted offer categoryId'
      })
      .lean()
    let validItems = []
    let removedNames = []

    if (cart && cart.items.length) {
      for (const item of cart.items) {
        const product = item.productId
        if (!product || product.isDeleted || !product.isListed) {
          removedNames.push(product?.name || 'A product')
          continue
        }
        const variant = findVariant(product, item.shade)
        if (!variant) continue

        const currentPrice = getVariantPrice(variant)
        const inStock = variant.stock > 0


        validItems.push({
          ...item,
          product,
          variant,
          currentPrice,
          inStock,
          stockAvailable: variant.stock,

          qtyExceedsStock: item.quantity > variant.stock,
          qtyExceedsMax: item.quantity > MAX_QTY_PER_ITEM,
          effectiveQty: Math.min(item.quantity, variant.stock, MAX_QTY_PER_ITEM)

        })
      }
    }

    const subtotal = validItems
      .filter(i => i.inStock)
      .reduce((sum, i) => sum + i.currentPrice * i.effectiveQty, 0)


    const hasOutOfStock = validItems.some(i => !i.inStock)
    const hasQtyIssues = validItems.some(i => i.qtyExceedsStock || i.qtyExceedsMax)
    const canCheckout = validItems.length > 0 && !hasOutOfStock && !hasQtyIssues

    const totalItemDiscount = validItems.reduce((sum, item) => {
      const original = item.variant?.varientPrice || 0
      const current = item.currentPrice || 0
      return sum + ((original - current) * item.effectiveQty)
    }, 0)


    res.render('user/cart', {
      items: validItems,
      subtotal,
      totalItemDiscount,
      removedNames,
      hasOutOfStock,
      hasQtyIssues,
      canCheckout,
      maxQty: MAX_QTY_PER_ITEM,
      user: req.session.user || null,

    })

  } catch (error) {
    console.error('loadCart error:', error)
    res.status(500).redirect('/products')
  }
}


const addToCart = async (req, res) => {
  try {
    const userId = req.session.user?._id
    if (!userId) {
      return res.status(401).json({ success: false, message: 'Please log in.', redirect: '/login' })
    }

    const { productId, quantity = 1, shade = '' } = req.body
    const qty = parseInt(quantity) || 1

    // ── validate product ──
    const product = await Product.findOne({ _id: productId, isDeleted: false, isListed: true })
    if (!product) {
      return res.status(404).json({ success: false, message: 'Product is no longer available.' })
    }

    const variant = findVariant(product, shade)
    if (!variant) {
      return res.status(400).json({ success: false, message: 'Shade not found.' })
    }

    if (variant.stock <= 0) {
      return res.status(400).json({ success: false, message: 'This shade is out of stock.' })
    }

    // ── get or create cart ──
    let cart = await Cart.findOne({ userId })
    if (!cart) cart = new Cart({ userId, items: [] })

    // check if same product+shade already in cart
    const existingIdx = cart.items.findIndex(
      i => String(i.productId) === String(productId) &&
        i.shade.toLowerCase() === (shade || variant.shade).toLowerCase()
    )

    if (existingIdx > -1) {
      // item already exists — increment quantity
      const newQty = cart.items[existingIdx].quantity + qty

      if (newQty > MAX_QTY_PER_ITEM) {
        return res.status(400).json({
          success: false,
          message: `Maximum ${MAX_QTY_PER_ITEM} units allowed per item.`
        })
      }
      if (newQty > variant.stock) {
        return res.status(400).json({
          success: false,
          message: `Only ${variant.stock} units available in this shade.`
        })
      }
      cart.items[existingIdx].quantity = newQty
      cart.items[existingIdx].priceAtAdd = getVariantPrice(variant)
    } else {
      // new item
      if (qty > MAX_QTY_PER_ITEM) {
        return res.status(400).json({
          success: false,
          message: `Maximum ${MAX_QTY_PER_ITEM} units allowed per item.`
        })
      }
      if (qty > variant.stock) {
        return res.status(400).json({
          success: false,
          message: `Only ${variant.stock} units available in stock.`
        })
      }
      cart.items.push({
        productId: product._id,
        shade: shade || variant.shade,
        quantity: qty,
        priceAtAdd: getVariantPrice(variant)
      })
    }

    await cart.save()

    // ── remove from wishlist if present ──
    await Wishlist.findOneAndUpdate(
      { userId },
      { $pull: { products: product._id } }
    )

    const cartCount = cart.items.reduce((s, i) => s + i.quantity, 0)
    res.json({ success: true, message: 'Added to cart!', cartCount })

  } catch (error) {
    console.error('addToCart error:', error)
    res.status(500).json({ success: false, message: 'Could not add to cart.' })
  }
}

const updateCartItem = async (req, res) => {
  try {
    const userId = req.session.user?._id
    if (!userId) return res.status(401).json({ success: false, message: 'Not logged in.' })

    const { itemId, action } = req.body

    const cart = await Cart.findOne({ userId })
    if (!cart) return res.status(404).json({ success: false, message: 'Cart not found.' })

    const itemIdx = cart.items.findIndex(i => String(i._id) === String(itemId))
    if (itemIdx === -1) return res.status(404).json({ success: false, message: 'Item not found.' })

    const item = cart.items[itemIdx]
    const product = await Product.findOne({ _id: item.productId, isDeleted: false, isListed: true })

    if (!product) {
      cart.items.splice(itemIdx, 1)
      await cart.save()
      return res.json({ success: true, removed: true, message: 'Product no longer available.' })
    }

    const variant = findVariant(product, item.shade)
    if (!variant) {
      cart.items.splice(itemIdx, 1)
      await cart.save()
      return res.json({ success: true, removed: true, message: 'Shade no longer available.' })
    }

    let newQty = item.quantity

    if (action === 'increment') {
      if (item.quantity >= MAX_QTY_PER_ITEM) {
        return res.status(400).json({ success: false, message: `Maximum ${MAX_QTY_PER_ITEM} units allowed.` })
      }
      if (item.quantity >= variant.stock) {
        return res.status(400).json({ success: false, message: `Only ${variant.stock} units in stock.` })
      }
      newQty = item.quantity + 1
    } else if (action === 'decrement') {
      if (item.quantity <= 1) {
        cart.items.splice(itemIdx, 1)
        await cart.save()
        const cartCount = cart.items.reduce((s, i) => s + i.quantity, 0)
        return res.json({ success: true, removed: true, cartCount })
      }
      newQty = item.quantity - 1
    }

    cart.items[itemIdx].quantity = newQty
    cart.items[itemIdx].priceAtAdd = getVariantPrice(variant)
    await cart.save()

    const cartCount = cart.items.reduce((s, i) => s + i.quantity, 0)
    const itemTotal = getVariantPrice(variant) * newQty
    res.json({ success: true, newQty, itemTotal, cartCount })

  } catch (err) {
    console.error('updateCartItem error:', err)
    res.status(500).json({ success: false, message: 'Could not update cart.' })
  }
}

const removeFromCart = async (req, res) => {
  try {
    const userId = req.session.user?._id
    if (!userId) return res.status(401).json({ success: false, message: 'Not logged in.' })

    const cart = await Cart.findOne({ userId })
    if (!cart) return res.status(404).json({ success: false, message: 'Cart not found.' })

    cart.items = cart.items.filter(i => String(i._id) !== String(req.params.itemId))
    await cart.save()

    const cartCount = cart.items.reduce((s, i) => s + i.quantity, 0)
    res.json({ success: true, message: 'Item removed.', cartCount })

  } catch (err) {
    console.error('removeFromCart error:', err)
    res.status(500).json({ success: false, message: 'Could not remove item.' })
  }
}

const loadWishlist = async (req, res) => {
  try {
    const userId = req.session.user?._id
    if (!userId) return res.redirect('/login')

    const wishlist = await Wishlist.findOne({ userId })
      .populate({
        path: 'products',
        match: { isDeleted: false, isListed: true },
        select: 'name images variants offer createdAt',
        populate: { path: 'categoryId', select: 'name' }
      })
      .lean()

    const products = (wishlist?.products || []).map(p => {
      const prices = p.variants.map(v => v.salePrice > 0 ? v.salePrice : v.varientPrice)
      p.displayPrice = Math.min(...prices)
      const origPrices = p.variants.map(v => v.varientPrice)
      p.originalPrice = Math.min(...origPrices)
      p.displayOffer = p.offer || 0
      p.inStock = p.variants.some(v => v.stock > 0)
      return p
    })

    res.render('user/wishlist', {
      products,
      user: req.session.user || null
    })

  } catch (err) {
    console.error('loadWishlist error:', err)
    res.redirect('/products')
  }
}

const toggleWishlist = async (req, res) => {
  try {
    const userId = req.session.user?._id
    if (!userId) {
      return res.status(401).json({ success: false, message: 'Please log in', redirect: '/login' })
    }

    const { productId } = req.body
    const product = await Product.findOne({ _id: productId, isDeleted: false, isListed: true })
    if (!product) {
      return res.status(404).json({ success: false, message: 'Product not found' })

    }

    let wishlist = await Wishlist.findOne({ userId })
    if (!wishlist) wishlist = new Wishlist({ userId, products: [] })

    const idx = wishlist.products.findIndex(id => String(id) === String(productId))
    let added = false

    if (idx > -1) {
      wishlist.products.splice(idx, 1)
    } else {
      wishlist.products.push(productId)
      added = true
    }

    await wishlist.save()
    res.json({ success: true, added, message: added ? 'Added to wishlist' : 'Removed From wishlist' })
  } catch (error) {
    console.error('toggleWishlist error:', err)
    res.status(500).json({ success: false, message: 'Could not update wishlist' })
  }
}

const applyCoupon = async (req, res) => {
  try {
    const { code } = req.body
    const userId = req.session.user?._id

    if (!code) {
      return res.json({ success: false, message: 'Please enter a coupon code.' })
    }

    const cart = await Cart.findOne({ userId }).populate({
      path: 'items.productId',
      select: 'variants offer isDeleted isListed'
    }).lean()

    if (!cart || !cart.items.length) {
      return res.json({ success: false, message: 'Your cart is empty.' })
    }

    const cartTotal = cart.items.reduce((sum, item) => {
      const product = item.productId
      if (!product || product.isDeleted || !product.isListed) return sum
      const variant = product.variants?.find(v =>
        v.shade?.toLowerCase() === item.shade?.toLowerCase()
      ) || product.variants?.[0]
      if (!variant) return sum
      const price = variant.salePrice > 0 ? variant.salePrice : variant.varientPrice
      return sum + price * item.quantity
    }, 0)

    if (req.session.coupon) {
      return res.json({ success: false, message: 'A coupon is already applied. Remove it first.' })
    }

    const coupon = await Coupon.findOne({ code: code.toUpperCase(), isActive: true })
    if (!coupon) {
      return res.json({ success: false, message: 'Invalid coupon code.' })
    }

    if (coupon.expiresAt && coupon.expiresAt < new Date()) {
      return res.json({ success: false, message: 'This coupon has expired.' })
    }

    if (coupon.usedBy.some(id => id.toString() === userId.toString())) {
      return res.json({ success: false, message: 'You have already used this coupon.' })
    }

    if (coupon.maxUses && coupon.usedBy.length >= coupon.maxUses) {
      return res.json({ success: false, message: 'Coupon usage limit reached.' })
    }

    if (cartTotal < coupon.minOrderAmount) {
      return res.json({
        success: false,
        message: `Minimum order amount is ₹${coupon.minOrderAmount} to use this coupon.`
      })
    }

    // ── Calculate discount based on type ─────────────────────────────────
    let discount = 0

    if (coupon.discountType === 'percentage') {
      discount = Math.round((cartTotal * coupon.discountAmount) / 100)
      // cap at maxDiscount if set
      if (coupon.maxDiscount && discount > coupon.maxDiscount) {
        discount = coupon.maxDiscount
      }
    } else {
      // flat
      discount = coupon.discountAmount
    }

    // discount can never exceed cart total
    discount = Math.min(discount, cartTotal)

    req.session.coupon = {
      code:     coupon.code,
      discount: discount,
      couponId: coupon._id
    }

    const label = coupon.discountType === 'percentage'
      ? `${coupon.discountAmount}% off`
      : `₹${discount} off`

    return res.json({
      success:  true,
      discount: discount,
      message:  `Coupon applied! ${label}`
    })

  } catch (err) {
    console.error('applyCoupon error:', err)
    res.status(500).json({ success: false, message: 'Server error. Please try again.' })
  }
}

const removeCoupon = (req, res) => {
  req.session.coupon = null;
  res.json({ success: true })
}


module.exports = {
  loadCart,
  addToCart,
  updateCartItem,
  removeFromCart,
  toggleWishlist,
  loadWishlist,
  applyCoupon,
  removeCoupon,
}