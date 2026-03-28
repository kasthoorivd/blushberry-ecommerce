const Cart    = require('../../models/user/cartModel')
const Address = require('../../models/user/addressModel')
const Product = require('../../models/user/productModel')
const Order   = require('../../models/user/orderModel')
const PDFDocument = require('pdfkit')

const SHIPPING_THRESHOLD = 499   // free shipping above this
const SHIPPING_CHARGE    = 49
const TAX_RATE           = 0.05  // 5% GST

/* ─────────────────────────────────────────
   helper: build enriched cart lines
   ───────────────────────────────────────── */
async function buildCartLines(cartItems) {
  const lines = []

  for (const item of cartItems) {
    const product = await Product.findById(item.productId).lean()
    if (!product) continue

    // find the matching variant by shade
    const variant = product.variants?.find(v => v.shade === item.shade) || product.variants?.[0]
    if (!variant) continue

    const originalPrice = variant.varientPrice || 0
    const salePrice     = variant.salePrice     || 0

    // effective price = salePrice if set, else originalPrice
    const effectivePrice = salePrice > 0 && salePrice < originalPrice ? salePrice : originalPrice

    // product-level offer discount
    const offerDiscount = product.offer > 0
      ? Math.round(originalPrice * (product.offer / 100))
      : 0

    const priceAfterOffer  = originalPrice - offerDiscount
    const finalUnitPrice   = salePrice > 0 && salePrice < priceAfterOffer ? salePrice : priceAfterOffer
    const itemTotal        = finalUnitPrice * item.quantity
    const itemDiscount     = (originalPrice - finalUnitPrice) * item.quantity

    // pick best image: shade gallery first, then product images
    const shadeImages = variant.images?.length ? variant.images : null
    const image = shadeImages?.[0] || product.images?.[0] || ''

    lines.push({
      cartItemId:   item._id,
      productId:    product._id,
      productName:  product.name,
      shade:        item.shade,
      image,
      quantity:     item.quantity,
      originalPrice,
      offerDiscount,
      finalUnitPrice,
      itemTotal,
      itemDiscount,
      stock:        variant.stock || 0,
      offer:        product.offer || 0,
    })
  }

  return lines
}

/* ─────────────────────────────────────────
   helper: compute order totals
   ───────────────────────────────────────── */
function computeTotals(lines, couponDiscount = 0) {
  const subtotal      = lines.reduce((s, l) => s + l.originalPrice * l.quantity, 0)
  const itemDiscounts = lines.reduce((s, l) => s + l.itemDiscount, 0)
  const shippingCharge = (subtotal - itemDiscounts - couponDiscount) > SHIPPING_THRESHOLD ? 0 : SHIPPING_CHARGE
  const taxableAmount  = subtotal - itemDiscounts - couponDiscount
  const tax            = Math.round(taxableAmount * TAX_RATE)
  const finalAmount    = taxableAmount + shippingCharge + tax

  return {
    subtotal,
    itemDiscounts,
    couponDiscount,
    totalDiscount: itemDiscounts + couponDiscount,
    shippingCharge,
    tax,
    finalAmount: Math.max(finalAmount, 0),
    freeShippingThreshold: SHIPPING_THRESHOLD,
    amountForFreeShipping: Math.max(SHIPPING_THRESHOLD - (subtotal - itemDiscounts - couponDiscount), 0)
  }
}

/* ─────────────────────────────────────────
   GET /checkout
   ───────────────────────────────────────── */
const loadCheckout = async (req, res) => {
  try {
    const userId = req.session.user._id

    // fetch cart
    const cart = await Cart.findOne({ userId }).lean()
    if (!cart || !cart.items.length) {
      return res.redirect('/cart')
    }

    // fetch addresses
    const addresses = await Address.find({ user: userId }).sort({ isDefault: -1, createdAt: -1 }).lean()

    // build enriched lines
    const lines = await buildCartLines(cart.items)
    if (!lines.length) return res.redirect('/cart')

    // check stock issues
    const outOfStock = lines.filter(l => l.stock < l.quantity)

    // coupon from session
    const couponDiscount = req.session.coupon?.discount || 0
    const couponCode     = req.session.coupon?.code     || null

    const totals = computeTotals(lines, couponDiscount)

    return res.render('user/checkout', {
      addresses,
      lines,
      totals,
      couponDiscount,
      couponCode,
      outOfStock,
      user: req.session.user,
    })
  } catch (err) {
    console.error('loadCheckout error:', err)
    return res.redirect('/cart')
  }
}

/* ─────────────────────────────────────────
   POST /checkout/place-order
   ───────────────────────────────────────── */
const placeOrder = async (req, res) => {
  try {
    const userId = req.session.user._id
    const { addressId, paymentMethod = 'COD' } = req.body

    // validate address
    if (!addressId) {
      return res.status(400).json({ success: false, message: 'Please select a delivery address.' })
    }

    const address = await Address.findOne({ _id: addressId, user: userId }).lean()
    if (!address) {
      return res.status(400).json({ success: false, message: 'Invalid address selected.' })
    }

    // fetch cart
    const cart = await Cart.findOne({ userId })
    if (!cart || !cart.items.length) {
      return res.status(400).json({ success: false, message: 'Your cart is empty.' })
    }

    // build lines & validate stock
    const lines = await buildCartLines(cart.items)
    if (!lines.length) {
      return res.status(400).json({ success: false, message: 'No valid items in cart.' })
    }

    // stock check
    for (const line of lines) {
      if (line.stock < line.quantity) {
        return res.status(400).json({
          success: false,
          message: `"${line.productName} - ${line.shade}" only has ${line.stock} units left.`
        })
      }
    }

    const couponDiscount = req.session.coupon?.discount || 0
    const couponCode     = req.session.coupon?.code     || null
    const totals         = computeTotals(lines, couponDiscount)

    // build order items
    const orderItems = lines.map(l => ({
      productId:    l.productId,
      productName:  l.productName,
      productImage: l.image,
      shade:        l.shade,
      quantity:     l.quantity,
      priceAtOrder: l.originalPrice,
      salePrice:    l.finalUnitPrice,
      discount:     l.itemDiscount,
    }))

    // snapshot address
    const shippingAddress = {
      name:        address.name,
      address:     address.address,
      city:        address.city        || '',
      state:       address.state,
      country:     address.country,
      pincode:     address.pincode,
      mobile:      address.mobile,
      email:       address.email,
      landmark:    address.landmark    || '',
      addressType: address.addressType || 'Home',
    }

    // create order
    const order = new Order({
      userId,
      items:           orderItems,
      shippingAddress,
      subtotal:        totals.subtotal,
      totalDiscount:   totals.totalDiscount,
      shippingCharge:  totals.shippingCharge,
      tax:             totals.tax,
      finalAmount:     totals.finalAmount,
      couponCode,
      couponDiscount,
      paymentMethod,
      paymentStatus:   paymentMethod === 'COD' ? 'Pending' : 'Paid',
      orderStatus:     'Placed',
      itemStatuses:    orderItems.map((_, i) => ({ itemId: undefined, status: 'Active' })),
    })

    await order.save()

    // deduct stock
    for (const line of lines) {
      await Product.updateOne(
        { _id: line.productId, 'variants.shade': line.shade },
        { $inc: { 'variants.$.stock': -line.quantity } }
      )
    }

    // clear cart & coupon
    cart.items = []
    await cart.save()
    delete req.session.coupon

    return res.status(200).json({
      success:  true,
      message:  'Order placed successfully!',
      orderId:  order.orderId,
      dbId:     order._id,
    })
  } catch (err) {
    console.error('placeOrder error:', err)
    return res.status(500).json({ success: false, message: 'Something went wrong. Please try again.' })
  }
}

/* ─────────────────────────────────────────
   GET /order-success/:orderId
   ───────────────────────────────────────── */
const loadOrderSuccess = async (req, res) => {
  try {
    const order = await Order.findById(req.params.orderId).lean()
    if (!order || order.userId.toString() !== req.session.user._id.toString()) {
      return res.redirect('/')
    }
    return res.render('user/orderSuccess', { order, user: req.session.user })
  } catch (err) {
    console.error('loadOrderSuccess error:', err)
    return res.redirect('/')
  }
}


/* ─────────────────────────────────────────
   GET /orders/:orderId  (order detail)
   ───────────────────────────────────────── */
const loadOrderDetail = async (req, res) => {
  try {
    const order = await Order.findById(req.params.orderId).lean()
    if (!order || order.userId.toString() !== req.session.user._id.toString()) {
      return res.redirect('/orders')
    }
    return res.render('user/orderDetail', { order, user: req.session.user })
  } catch (err) {
    console.error('loadOrderDetail error:', err)
    return res.redirect('/orders')
  }
}

/* ─────────────────────────────────────────
   POST /orders/:orderId/cancel
   body: { reason, itemId? }
   ───────────────────────────────────────── */
const cancelOrder = async (req, res) => {
  try {
    const { reason = '', itemId } = req.body
    const order = await Order.findById(req.params.orderId)

    if (!order || order.userId.toString() !== req.session.user._id.toString()) {
      return res.status(404).json({ success: false, message: 'Order not found.' })
    }

    // ── ITEM-LEVEL CANCEL ──
    if (itemId) {
      const item = order.items.id(itemId)
      if (!item) return res.status(404).json({ success: false, message: 'Item not found.' })

      const itemStatus = order.itemStatuses.find(s => s.itemId?.toString() === itemId)
      if (itemStatus && itemStatus.status !== 'Active') {
        return res.status(400).json({ success: false, message: 'Item already cancelled or returned.' })
      }

      // update item status
      if (itemStatus) {
        itemStatus.status       = 'Cancelled'
        itemStatus.cancelReason = reason
      } else {
        order.itemStatuses.push({ itemId, status: 'Cancelled', cancelReason: reason })
      }

      // restore stock for this item
      await Product.updateOne(
        { _id: item.productId, 'variants.shade': item.shade },
        { $inc: { 'variants.$.stock': item.quantity } }
      )

      // check if all items cancelled → cancel whole order
      const activeItems = order.items.filter(i => {
        const s = order.itemStatuses.find(s => s.itemId?.toString() === i._id.toString())
        return !s || s.status === 'Active'
      })
      if (activeItems.length === 0) {
        order.orderStatus  = 'Cancelled'
        order.cancelReason = 'All items cancelled'
        order.cancelledAt  = new Date()
      }

      await order.save()
      return res.json({ success: true, message: 'Item cancelled successfully.' })
    }

    // ── WHOLE ORDER CANCEL ──
    if (!['Placed', 'Processing'].includes(order.orderStatus)) {
      return res.status(400).json({ success: false, message: 'This order cannot be cancelled.' })
    }

    order.orderStatus  = 'Cancelled'
    order.cancelReason = reason
    order.cancelledAt  = new Date()

    // mark all active items as cancelled
    order.items.forEach(item => {
      const existing = order.itemStatuses.find(s => s.itemId?.toString() === item._id.toString())
      if (existing) {
        if (existing.status === 'Active') { existing.status = 'Cancelled'; existing.cancelReason = reason }
      } else {
        order.itemStatuses.push({ itemId: item._id, status: 'Cancelled', cancelReason: reason })
      }
    })

    await order.save()

    // restore stock for all items
    for (const item of order.items) {
      await Product.updateOne(
        { _id: item.productId, 'variants.shade': item.shade },
        { $inc: { 'variants.$.stock': item.quantity } }
      )
    }

    return res.json({ success: true, message: 'Order cancelled successfully.' })
  } catch (err) {
    console.error('cancelOrder error:', err)
    return res.status(500).json({ success: false, message: 'Something went wrong.' })
  }
}

/* ─────────────────────────────────────────
   POST /orders/:orderId/return
   body: { reason, itemId? }
   ───────────────────────────────────────── */
const returnOrder = async (req, res) => {
  try {
    const { reason, itemId } = req.body

    if (!reason || !reason.trim()) {
      return res.status(400).json({ success: false, message: 'Return reason is required.' })
    }

    const order = await Order.findById(req.params.orderId)
    if (!order || order.userId.toString() !== req.session.user._id.toString()) {
      return res.status(404).json({ success: false, message: 'Order not found.' })
    }

    if (order.orderStatus !== 'Delivered') {
      return res.status(400).json({ success: false, message: 'Only delivered orders can be returned.' })
    }

    // ── ITEM-LEVEL RETURN ──
    if (itemId) {
      const item = order.items.id(itemId)
      if (!item) return res.status(404).json({ success: false, message: 'Item not found.' })

      const itemStatus = order.itemStatuses.find(s => s.itemId?.toString() === itemId)
      if (itemStatus && itemStatus.status !== 'Active') {
        return res.status(400).json({ success: false, message: 'Item already cancelled or returned.' })
      }

      if (itemStatus) {
        itemStatus.status       = 'Returned'
        itemStatus.returnReason = reason
        itemStatus.returnStatus = 'Requested'
      } else {
        order.itemStatuses.push({ itemId, status: 'Returned', returnReason: reason, returnStatus: 'Requested' })
      }

      await order.save()
      return res.json({ success: true, message: 'Return requested for item.' })
    }

    // ── WHOLE ORDER RETURN ──
    order.returnStatus = 'Requested'
    order.returnReason = reason
    order.returnedAt   = new Date()

    order.items.forEach(item => {
      const existing = order.itemStatuses.find(s => s.itemId?.toString() === item._id.toString())
      if (existing) {
        if (existing.status === 'Active') { existing.status = 'Returned'; existing.returnReason = reason; existing.returnStatus = 'Requested' }
      } else {
        order.itemStatuses.push({ itemId: item._id, status: 'Returned', returnReason: reason, returnStatus: 'Requested' })
      }
    })

    await order.save()
    return res.json({ success: true, message: 'Return request submitted successfully.' })
  } catch (err) {
    console.error('returnOrder error:', err)
    return res.status(500).json({ success: false, message: 'Something went wrong.' })
  }
}

/* ─────────────────────────────────────────
   GET /orders  — with search
   ───────────────────────────────────────── */
const loadOrders = async (req, res) => {
  try {
    const userId = req.session.user._id
    const search = (req.query.search?.trim() || '').replace(/^#/,'')

    let query = { userId }

    if (search) {
      query.$or = [
        { orderId:     { $regex: search, $options: 'i' } },
        { orderStatus: { $regex: search, $options: 'i' } },
        { 'items.productName': { $regex: search, $options: 'i' } },
      ]
    }

    const orders = await Order.find(query).sort({ createdAt: -1 }).lean()
    return res.render('user/orders', { orders, search, user: req.session.user })
  } catch (err) {
    console.error('loadOrders error:', err)
    return res.redirect('/')
  }
}

const downloadInvoice = async (req, res) => {
  try {
    const order = await Order.findById(req.params.orderId).lean()
    if (!order || order.userId.toString() !== req.session.user._id.toString()) {
      return res.redirect('/orders')
    }

    const doc = new PDFDocument({ margin: 50, size: 'A4' })

    res.setHeader('Content-Type', 'application/pdf')
    res.setHeader('Content-Disposition', `attachment; filename="invoice-${order.orderId}.pdf"`)
    doc.pipe(res)

    // ── COLORS & HELPERS ──
    const PINK    = '#c2185b'
    const DARK    = '#1a0a10'
    const MUTED   = '#8a6d74'
    const LIGHT   = '#fef0f4'
    const BORDER  = '#f0dde4'
    const SUCCESS = '#22b573'
    const W       = 495  // usable width (595 - 50 - 50)

    const rupee = n => `Rs. ${Number(n || 0).toLocaleString('en-IN')}`
    const fmt   = d => new Date(d).toLocaleDateString('en-IN', { day: 'numeric', month: 'long', year: 'numeric' })

    // ── HEADER BAND ──
    doc.rect(50, 45, W, 70).fill(LIGHT)

    doc.fontSize(26).font('Helvetica-Bold').fillColor(PINK)
       .text('Blushberry', 60, 60)

    doc.fontSize(8).font('Helvetica').fillColor(MUTED)
       .text('Beauty & Cosmetics', 60, 90)

    doc.fontSize(9).font('Helvetica-Bold').fillColor(DARK)
       .text('TAX INVOICE', 410, 60, { width: 125, align: 'right' })
    doc.fontSize(8).font('Helvetica').fillColor(MUTED)
       .text(`# ${order.orderId}`, 410, 75, { width: 125, align: 'right' })
       .text(`Date: ${fmt(order.createdAt)}`, 410, 88, { width: 125, align: 'right' })

    // ── DIVIDER ──
    let y = 130
    doc.moveTo(50, y).lineTo(545, y).strokeColor(BORDER).lineWidth(1).stroke()
    y += 15

    // ── BILLING & SHIPPING INFO ──
    doc.fontSize(7).font('Helvetica-Bold').fillColor(MUTED)
       .text('BILL TO / SHIP TO', 50, y)

    y += 12
    const addr = order.shippingAddress
    doc.fontSize(9).font('Helvetica-Bold').fillColor(DARK)
       .text(addr.name, 50, y)
    y += 13
    doc.fontSize(8).font('Helvetica').fillColor(MUTED)
    const addrLine = [addr.address, addr.city, addr.state, addr.country].filter(Boolean).join(', ')
    doc.text(addrLine, 50, y, { width: 250 })
    y += 12
    doc.text(`PIN: ${addr.pincode}`, 50, y)
    y += 12
    doc.text(`Mobile: ${addr.mobile}`, 50, y)

    // payment info — right side
    doc.fontSize(7).font('Helvetica-Bold').fillColor(MUTED)
       .text('PAYMENT INFO', 350, y - 49)
    doc.fontSize(8).font('Helvetica').fillColor(DARK)
       .text(`Method: ${order.paymentMethod === 'COD' ? 'Cash on Delivery' : order.paymentMethod}`, 350, y - 36)
       .text(`Status: ${order.paymentStatus}`, 350, y - 23)
       .text(`Order Status: ${order.orderStatus}`, 350, y - 10)

    y += 20
    doc.moveTo(50, y).lineTo(545, y).strokeColor(BORDER).lineWidth(0.5).stroke()
    y += 15

    // ── TABLE HEADER ──
    doc.rect(50, y, W, 22).fill(PINK)
    doc.fontSize(8).font('Helvetica-Bold').fillColor('#ffffff')
       .text('ITEM',           60,  y + 7, { width: 200 })
       .text('SHADE',         260,  y + 7, { width: 70  })
       .text('QTY',           330,  y + 7, { width: 35, align: 'center' })
       .text('UNIT PRICE',    365,  y + 7, { width: 75, align: 'right' })
       .text('TOTAL',         440,  y + 7, { width: 95, align: 'right' })
    y += 22

    // ── TABLE ROWS ──
    order.items.forEach((item, i) => {
      const rowH  = 28
      const bgCol = i % 2 === 0 ? '#ffffff' : LIGHT
      doc.rect(50, y, W, rowH).fill(bgCol)

      const itemStatus = (order.itemStatuses || []).find(
        s => s.itemId && s.itemId.toString() === item._id.toString()
      )
      const cancelled = itemStatus && itemStatus.status === 'Cancelled'

      doc.fontSize(8).font(cancelled ? 'Helvetica-Oblique' : 'Helvetica')
         .fillColor(cancelled ? MUTED : DARK)
         .text(item.productName + (cancelled ? ' (Cancelled)' : ''), 60, y + 9, { width: 195, ellipsis: true })

      doc.font('Helvetica').fillColor(MUTED)
         .text(item.shade || '—',           260, y + 9, { width: 65  })
         .text(String(item.quantity),        330, y + 9, { width: 35, align: 'center' })
         .text(rupee(item.salePrice),        365, y + 9, { width: 75, align: 'right'  })

      doc.fillColor(cancelled ? MUTED : DARK)
         .text(rupee(item.salePrice * item.quantity), 440, y + 9, { width: 95, align: 'right' })

      y += rowH
    })

    // ── TABLE BOTTOM BORDER ──
    doc.moveTo(50, y).lineTo(545, y).strokeColor(BORDER).lineWidth(0.5).stroke()
    y += 15

    // ── TOTALS BLOCK ──
    const totRow = (label, value, bold = false, color = DARK) => {
      if (bold) {
        doc.rect(50, y - 3, W, 22).fill(LIGHT)
        doc.fontSize(9).font('Helvetica-Bold').fillColor(color)
           .text(label, 50, y, { width: 390, align: 'right' })
           .text(value, 440, y, { width: 95,  align: 'right' })
        y += 22
      } else {
        doc.fontSize(8).font('Helvetica').fillColor(MUTED)
           .text(label, 50, y, { width: 390, align: 'right' })
        doc.fillColor(DARK)
           .text(value, 440, y, { width: 95, align: 'right' })
        y += 16
      }
    }

    totRow('Subtotal',                                      rupee(order.subtotal))
    if (order.totalDiscount > 0)
      totRow(`Discount`,                                    `- ${rupee(order.totalDiscount)}`, false, SUCCESS)
    if (order.couponDiscount > 0)
      totRow(`Coupon (${order.couponCode})`,                `- ${rupee(order.couponDiscount)}`, false, SUCCESS)
    totRow('Shipping',                                      order.shippingCharge === 0 ? 'FREE' : rupee(order.shippingCharge))
    totRow('GST (5%)',                                      rupee(order.tax))

    y += 4
    doc.moveTo(350, y).lineTo(545, y).strokeColor(BORDER).lineWidth(0.5).stroke()
    y += 8
    totRow('TOTAL PAID',                                    rupee(order.finalAmount), true, PINK)

    
    if (order.totalDiscount > 0) {
      y += 8
      doc.rect(50, y, W, 24).fill('rgba(34,181,115,0.08)').stroke()
      doc.rect(50, y, W, 24).strokeColor(SUCCESS).lineWidth(0.5).stroke()
      doc.fontSize(8).font('Helvetica-Bold').fillColor(SUCCESS)
         .text(`You saved ${rupee(order.totalDiscount)} on this order!`, 60, y + 8)
      y += 32
    }

    // ── FOOTER ──
    y = Math.max(y + 20, 720)
    doc.moveTo(50, y).lineTo(545, y).strokeColor(BORDER).lineWidth(0.5).stroke()
    y += 10
    doc.fontSize(7).font('Helvetica').fillColor(MUTED)
       .text('Thank you for shopping with Blushberry! For support, contact us at support@blushberry.com', 50, y, { align: 'center', width: W })
    y += 10
    doc.text('This is a computer-generated invoice and does not require a signature.', 50, y, { align: 'center', width: W })

    doc.end()
  } catch (err) {
    console.error('downloadInvoice error:', err)
    return res.redirect('/orders')
  }
}

module.exports = {
  loadCheckout,
  placeOrder,
  loadOrderSuccess,
  loadOrders,
  loadOrderDetail,
  cancelOrder,
  returnOrder,
  downloadInvoice
}