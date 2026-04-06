const Cart    = require('../../models/user/cartModel')
const User =require('../../models/user/userModel')
const Wallet = require('../../models/user/walletModel')
const Address = require('../../models/user/addressModel')
const Product = require('../../models/user/productModel')
const Order   = require('../../models/user/orderModel')
const Coupon = require('../../models/user/couponModel')
const PDFDocument = require('pdfkit')
const Razorpay = require('razorpay')
const crypto = require('crypto')

const { creditWallet, handleCancellationRefund } = require('./walletController')  // ✅ ADDED

const SHIPPING_THRESHOLD = 499
const SHIPPING_CHARGE    = 49

async function buildCartLines(cartItems) {
  const lines = []

  for (const item of cartItems) {
    const product = await Product.findById(item.productId).lean()
    if (!product) continue

    const variant = product.variants?.find(v => v.shade === item.shade) || product.variants?.[0]
    if (!variant) continue

    const originalPrice  = variant.varientPrice || 0
    const salePrice      = variant.salePrice     || 0
    const offerDiscount  = product.offer > 0
      ? Math.round(originalPrice * (product.offer / 100))
      : 0
    const priceAfterOffer  = originalPrice - offerDiscount
    const finalUnitPrice   = salePrice > 0 && salePrice < priceAfterOffer ? salePrice : priceAfterOffer
    const itemTotal        = finalUnitPrice * item.quantity
    const itemDiscount     = (originalPrice - finalUnitPrice) * item.quantity

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


function computeTotals(lines, couponDiscount = 0) {
  const subtotal       = lines.reduce((s, l) => s + l.originalPrice * l.quantity, 0)
  const itemDiscounts  = lines.reduce((s, l) => s + l.itemDiscount, 0)
  const amountAfterDiscounts = subtotal - itemDiscounts - couponDiscount
  const shippingCharge = amountAfterDiscounts > SHIPPING_THRESHOLD ? 0 : SHIPPING_CHARGE
  const finalAmount    = amountAfterDiscounts + shippingCharge

  return {
    subtotal,
    itemDiscounts,
    couponDiscount,
    totalDiscount: itemDiscounts + couponDiscount,
    shippingCharge,
    tax: 0,
    finalAmount: Math.max(finalAmount, 0),
    freeShippingThreshold: SHIPPING_THRESHOLD,
    amountForFreeShipping: Math.max(SHIPPING_THRESHOLD - amountAfterDiscounts, 0)
  }
}


const loadCheckout = async (req, res) => {
  try {
    const userId = req.session.user._id

    const cart = await Cart.findOne({ userId }).lean()
    if (!cart || !cart.items.length) return res.redirect('/cart')

    const addresses = await Address.find({ user: userId }).sort({ isDefault: -1, createdAt: -1 }).lean()
    const lines     = await buildCartLines(cart.items)
    if (!lines.length) return res.redirect('/cart')

    const outOfStock     = lines.filter(l => l.stock < l.quantity)
    const couponDiscount = req.session.coupon?.discount || 0
    const couponCode     = req.session.coupon?.code     || null
    const totals         = computeTotals(lines, couponDiscount)
   const wallet = await Wallet.findOne({userId})
   const walletBalance = wallet ? wallet.balance :  0
   
    return res.render('user/checkout', {
      addresses,
      lines,
      totals,
      couponDiscount,
      couponCode,
      outOfStock,
      walletBalance,
      user: req.session.user
    })
  } catch (err) {
    console.error('loadCheckout error:', err)
    return res.redirect('/cart')
  }
}


const placeOrder = async (req, res) => {
  try {
    const userId = req.session.user._id
    const { addressId, paymentMethod = 'COD' } = req.body
    console.log("Payment Method:", paymentMethod);

    if (!addressId) {
      return res.status(400).json({ success: false, message: 'Please select a delivery address.' })
    }

    const address = await Address.findOne({ _id: addressId, user: userId }).lean()
    if (!address) {
      return res.status(400).json({ success: false, message: 'Invalid address selected.' })
    }

    const cart = await Cart.findOne({ userId })
    if (!cart || !cart.items.length) {
      return res.status(400).json({ success: false, message: 'Your cart is empty.' })
    }

    const lines = await buildCartLines(cart.items)
    if (!lines.length) {
      return res.status(400).json({ success: false, message: 'No valid items in cart.' })
    }

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
    
     let wallet;

if (paymentMethod.toLowerCase() === 'wallet') {
  wallet = await Wallet.findOne({ userId });

  if (!wallet || wallet.balance < totals.finalAmount) {
    return res.status(400).json({
      success: false,
      message: 'Insufficient wallet balance'
    });
  }

  wallet.balance -= totals.finalAmount;

  wallet.transactions.push({
    type: 'debit',
    amount: totals.finalAmount,
    description: `Payment for order`,
    status: 'completed'
  });

  await wallet.save();
}
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

    const order = new Order({
      userId,
      items:           orderItems,
      shippingAddress,
      subtotal:        totals.subtotal,
      totalDiscount:   totals.totalDiscount,
      shippingCharge:  totals.shippingCharge,
      tax:             0,
      finalAmount:     totals.finalAmount,
      couponCode,
      couponDiscount,
      paymentMethod,
      paymentStatus:   paymentMethod === 'COD' ? 'Pending' :
     paymentMethod === 'WALLET' ? 'Paid' :
     'Paid',
      orderStatus:     'Placed',
      itemStatuses:    orderItems.map(() => ({ status: 'Active' })),
    })
    
  
    await order.save()
      if (paymentMethod.toLowerCase() === 'wallet' && wallet) {
  const lastTransaction = wallet.transactions[wallet.transactions.length - 1];
  if (lastTransaction) {
    lastTransaction.orderId = order._id;
  }
  await wallet.save();
}
    if (couponCode) {
      await Coupon.findOneAndUpdate(
        { code: couponCode },
        { $addToSet: { usedBy: userId } }
      )
    }

    for (const line of lines) {
      await Product.updateOne(
        { _id: line.productId, 'variants.shade': line.shade },
        { $inc: { 'variants.$.stock': -line.quantity } }
      )
    }

    cart.items = []
    await cart.save()
    delete req.session.coupon

    return res.status(200).json({
      success: true,
      message: 'Order placed successfully!',
      orderId: order.orderId,
      dbId:    order._id,
    })
  } catch (err) {
    console.error('placeOrder error:', err)
    return res.status(500).json({ success: false, message: 'Something went wrong. Please try again.' })
  }
}


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

const cancelOrder = async (req, res) => {
  try {
    const { reason = '', itemId } = req.body
    const order = await Order.findById(req.params.orderId)

    if (!order || order.userId.toString() !== req.session.user._id.toString()) {
      return res.status(404).json({ success: false, message: 'Order not found.' })
    }

    // ── SINGLE ITEM CANCEL ──────────────────────────────────────────────────
    if (itemId) {
      const item = order.items.id(itemId)
      if (!item) return res.status(404).json({ success: false, message: 'Item not found.' })

      const itemStatus = order.itemStatuses.find(s => s.itemId?.toString() === itemId)
      if (itemStatus && itemStatus.status !== 'Active') {
        return res.status(400).json({ success: false, message: 'Item already cancelled or returned.' })
      }

      if (itemStatus) {
        itemStatus.status       = 'Cancelled'
        itemStatus.cancelReason = reason
      } else {
        order.itemStatuses.push({ itemId, status: 'Cancelled', cancelReason: reason })
      }

      await Product.updateOne(
        { _id: item.productId, 'variants.shade': item.shade },
        { $inc: { 'variants.$.stock': item.quantity } }
      )

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

    
      const isPaid = order.paymentMethod !== 'COD' || order.paymentStatus === 'Paid'
      if (isPaid) {
        const refundAmount = item.salePrice * item.quantity
        await creditWallet(
          req.session.user._id,
          refundAmount,
          `Refund for cancelled item "${item.productName}" in order #${order.orderId}`,
          order._id
        )
      }

      return res.json({ success: true, message: 'Item cancelled successfully.' })
    }

    // ── FULL ORDER CANCEL ───────────────────────────────────────────────────
    if (!['Placed', 'Processing'].includes(order.orderStatus)) {
      return res.status(400).json({ success: false, message: 'This order cannot be cancelled.' })
    }

    order.orderStatus  = 'Cancelled'
    order.cancelReason = reason
    order.cancelledAt  = new Date()

    order.items.forEach(item => {
      const existing = order.itemStatuses.find(s => s.itemId?.toString() === item._id.toString())
      if (existing) {
        if (existing.status === 'Active') { existing.status = 'Cancelled'; existing.cancelReason = reason }
      } else {
        order.itemStatuses.push({ itemId: item._id, status: 'Cancelled', cancelReason: reason })
      }
    })

    await order.save()

    for (const item of order.items) {
      await Product.updateOne(
        { _id: item.productId, 'variants.shade': item.shade },
        { $inc: { 'variants.$.stock': item.quantity } }
      )
    }

   
    await handleCancellationRefund(order, req.session.user._id)

    return res.json({ success: true, message: 'Order cancelled successfully.' })
  } catch (err) {
    console.error('cancelOrder error:', err)
    return res.status(500).json({ success: false, message: 'Something went wrong.' })
  }
}


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


const loadOrders = async (req, res) => {
  try {
    const userId = req.session.user._id
    const search = (req.query.search?.trim() || '').replace(/^#/, '')

    let query = { userId }
    if (search) {
      query.$or = [
        { orderId:             { $regex: search, $options: 'i' } },
        { orderStatus:         { $regex: search, $options: 'i' } },
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

    const PINK    = '#c2185b'
    const DARK    = '#1a0a10'
    const MUTED   = '#8a6d74'
    const LIGHT   = '#fef0f4'
    const BORDER  = '#f0dde4'
    const SUCCESS = '#22b573'
    const W       = 495

    const rupee = n => `Rs. ${Number(n || 0).toLocaleString('en-IN')}`
    const fmt   = d => new Date(d).toLocaleDateString('en-IN', { day: 'numeric', month: 'long', year: 'numeric' })

    doc.rect(50, 45, W, 70).fill(LIGHT)
    doc.fontSize(26).font('Helvetica-Bold').fillColor(PINK).text('Blushberry', 60, 60)
    doc.fontSize(8).font('Helvetica').fillColor(MUTED).text('Beauty & Cosmetics', 60, 90)
    doc.fontSize(9).font('Helvetica-Bold').fillColor(DARK).text('INVOICE', 410, 60, { width: 125, align: 'right' })
    doc.fontSize(8).font('Helvetica').fillColor(MUTED)
       .text(`# ${order.orderId}`, 410, 75, { width: 125, align: 'right' })
       .text(`Date: ${fmt(order.createdAt)}`, 410, 88, { width: 125, align: 'right' })

    let y = 130
    doc.moveTo(50, y).lineTo(545, y).strokeColor(BORDER).lineWidth(1).stroke()
    y += 15

    doc.fontSize(7).font('Helvetica-Bold').fillColor(MUTED).text('BILL TO / SHIP TO', 50, y)
    y += 12
    const addr = order.shippingAddress
    doc.fontSize(9).font('Helvetica-Bold').fillColor(DARK).text(addr.name, 50, y)
    y += 13
    doc.fontSize(8).font('Helvetica').fillColor(MUTED)
    const addrLine = [addr.address, addr.city, addr.state, addr.country].filter(Boolean).join(', ')
    doc.text(addrLine, 50, y, { width: 250 })
    y += 12
    doc.text(`PIN: ${addr.pincode}`, 50, y)
    y += 12
    doc.text(`Mobile: ${addr.mobile}`, 50, y)

    doc.fontSize(7).font('Helvetica-Bold').fillColor(MUTED).text('PAYMENT INFO', 350, y - 49)
    doc.fontSize(8).font('Helvetica').fillColor(DARK)
       .text(`Method: ${order.paymentMethod === 'COD' ? 'Cash on Delivery' : order.paymentMethod}`, 350, y - 36)
       .text(`Status: ${order.paymentStatus}`, 350, y - 23)
       .text(`Order Status: ${order.orderStatus}`, 350, y - 10)

    y += 20
    doc.moveTo(50, y).lineTo(545, y).strokeColor(BORDER).lineWidth(0.5).stroke()
    y += 15

    doc.rect(50, y, W, 22).fill(PINK)
    doc.fontSize(8).font('Helvetica-Bold').fillColor('#ffffff')
       .text('ITEM',        60,  y + 7, { width: 200 })
       .text('SHADE',      260,  y + 7, { width: 70  })
       .text('QTY',        330,  y + 7, { width: 35, align: 'center' })
       .text('UNIT PRICE', 365,  y + 7, { width: 75, align: 'right'  })
       .text('TOTAL',      440,  y + 7, { width: 95, align: 'right'  })
    y += 22

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
         .text(item.shade || '—',                        260, y + 9, { width: 65  })
         .text(String(item.quantity),                    330, y + 9, { width: 35, align: 'center' })
         .text(rupee(item.salePrice),                    365, y + 9, { width: 75, align: 'right'  })
      doc.fillColor(cancelled ? MUTED : DARK)
         .text(rupee(item.salePrice * item.quantity),    440, y + 9, { width: 95, align: 'right'  })

      y += rowH
    })

    doc.moveTo(50, y).lineTo(545, y).strokeColor(BORDER).lineWidth(0.5).stroke()
    y += 15

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

    totRow('Subtotal', rupee(order.subtotal))
    if (order.totalDiscount > 0)
      totRow('Discount', `- ${rupee(order.totalDiscount)}`, false, SUCCESS)
    if (order.couponDiscount > 0)
      totRow(`Coupon (${order.couponCode})`, `- ${rupee(order.couponDiscount)}`, false, SUCCESS)
    totRow('Shipping', order.shippingCharge === 0 ? 'FREE' : rupee(order.shippingCharge))

    y += 4
    doc.moveTo(350, y).lineTo(545, y).strokeColor(BORDER).lineWidth(0.5).stroke()
    y += 8
    totRow('TOTAL', rupee(order.finalAmount), true, PINK)

    if (order.totalDiscount > 0) {
      y += 8
      doc.rect(50, y, W, 24).fill('rgba(34,181,115,0.08)').stroke()
      doc.rect(50, y, W, 24).strokeColor(SUCCESS).lineWidth(0.5).stroke()
      doc.fontSize(8).font('Helvetica-Bold').fillColor(SUCCESS)
         .text(`You saved ${rupee(order.totalDiscount)} on this order!`, 60, y + 8)
      y += 32
    }

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

const razorpay = new Razorpay({
  key_id:     process.env.RAZORPAY_KEY_ID,
  key_secret: process.env.RAZORPAY_KEY_SECRET,
})


const createRazorpayOrder = async (req, res) => {
  try {
    const userId = req.session.user._id
    const { addressId } = req.body

    if (!addressId) {
      return res.status(400).json({ success: false, message: 'Please select a delivery address.' })
    }

    const cart = await Cart.findOne({ userId }).lean()
    if (!cart || !cart.items.length) {
      return res.status(400).json({ success: false, message: 'Your cart is empty.' })
    }

    const lines = await buildCartLines(cart.items)
    if (!lines.length) {
      return res.status(400).json({ success: false, message: 'No valid items in cart.' })
    }

    for (const line of lines) {
      if (line.stock < line.quantity) {
        return res.status(400).json({
          success: false,
          message: `"${line.productName} - ${line.shade}" only has ${line.stock} units left.`
        })
      }
    }

    const couponDiscount = req.session.coupon?.discount || 0
    const totals         = computeTotals(lines, couponDiscount)
    const amountInPaise  = Math.round(totals.finalAmount * 100)

    const razorpayOrder = await razorpay.orders.create({
      amount:   amountInPaise,
      currency: 'INR',
      receipt:  `blushberry_${Date.now()}`,
    })

    const user = req.session.user

    return res.json({
      success:         true,
      razorpayOrderId: razorpayOrder.id,
      razorpayKeyId:   process.env.RAZORPAY_KEY_ID,
      amount:          amountInPaise,
      customerName:    user.name   || '',
      customerEmail:   user.email  || '',
      customerPhone:   user.mobile || '',
    })
  } catch (err) {
    console.error('createRazorpayOrder error:', err)
    return res.status(500).json({ success: false, message: 'Failed to initiate payment.' })
  }
}


const verifyPayment = async (req, res) => {
  try {
    const userId = req.session.user._id
    const { razorpay_order_id, razorpay_payment_id, razorpay_signature, addressId } = req.body

    const expectedSignature = crypto
      .createHmac('sha256', process.env.RAZORPAY_KEY_SECRET)
      .update(`${razorpay_order_id}|${razorpay_payment_id}`)
      .digest('hex')

    if (expectedSignature !== razorpay_signature) {
      return res.status(400).json({ success: false, message: 'Payment verification failed. Please contact support.' })
    }

    const address = await Address.findOne({ _id: addressId, user: userId }).lean()
    if (!address) {
      return res.status(400).json({ success: false, message: 'Invalid address.' })
    }

    const cart = await Cart.findOne({ userId })
    if (!cart || !cart.items.length) {
      return res.status(400).json({ success: false, message: 'Cart is empty.' })
    }

    const lines          = await buildCartLines(cart.items)
    const couponDiscount = req.session.coupon?.discount || 0
    const couponCode     = req.session.coupon?.code     || null
    const totals         = computeTotals(lines, couponDiscount)

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

    const order = new Order({
      userId,
      items:             orderItems,
      shippingAddress,
      subtotal:          totals.subtotal,
      totalDiscount:     totals.totalDiscount,
      shippingCharge:    totals.shippingCharge,
      tax:               0,
      finalAmount:       totals.finalAmount,
      couponCode,
      couponDiscount,
      paymentMethod:     'Online',
      paymentStatus:     'Paid',
      orderStatus:       'Placed',
      razorpayOrderId:   razorpay_order_id,
      razorpayPaymentId: razorpay_payment_id,
      itemStatuses:      orderItems.map(() => ({ status: 'Active' })),
    })

    await order.save()

    if (couponCode) {
      await Coupon.findOneAndUpdate(
        { code: couponCode },
        { $addToSet: { usedBy: userId } }
      )
    }

    for (const line of lines) {
      await Product.updateOne(
        { _id: line.productId, 'variants.shade': line.shade },
        { $inc: { 'variants.$.stock': -line.quantity } }
      )
    }

    cart.items = []
    await cart.save()
    delete req.session.coupon

    return res.json({ success: true, dbId: order._id })
  } catch (err) {
    console.error('verifyPayment error:', err)
    return res.status(500).json({ success: false, message: 'Something went wrong while saving your order.' })
  }
}

module.exports = {
  loadCheckout,
  placeOrder,
  createRazorpayOrder,
  verifyPayment,
  loadOrderSuccess,
  loadOrders,
  loadOrderDetail,
  cancelOrder,
  returnOrder,
  downloadInvoice,
}