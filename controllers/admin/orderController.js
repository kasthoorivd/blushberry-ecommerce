const Order   = require('../../models/user/orderModel')
const User    = require('../../models/user/userModel')
const Product = require('../../models/user/productModel')
const Coupon  = require('../../models/user/couponModel')   // ← adjust path if different
const { creditWallet } = require('../user/walletController')

const LIMIT = 5


// ─────────────────────────────────────────────
// helpers
// ─────────────────────────────────────────────

/** Returns the sum of salePrice×qty for items that are NOT cancelled */
function activeItemsTotal(order) {
  return order.items.reduce((sum, item) => {
    const s = (order.itemStatuses || []).find(
      st => st.itemId && st.itemId.toString() === item._id.toString()
    )
    const isCancelled = s && s.status === 'Cancelled'
    return isCancelled ? sum : sum + item.salePrice * item.quantity
  }, 0)
}

/**
 * Recalculates finalAmount for a COD order after an item is cancelled.
 * If a coupon was applied but the remaining active-item total no longer meets
 * the coupon's minimumOrderValue, the coupon is voided (discount set to 0).
 *
 * Formula (mirrors checkout):
 *   finalAmount = activeTotal - couponDiscount + shippingCharge + tax
 *
 * Returns { finalAmount, couponVoided }
 */
async function recalcCODAmount(order) {
  const total = activeItemsTotal(order)

  let couponDiscount = order.couponDiscount || 0
  let couponVoided   = false

  if (order.couponCode && couponDiscount > 0) {
    // Fetch the coupon to check its minimum order value
    const coupon = await Coupon.findOne({ code: order.couponCode })
    const minVal = coupon ? (coupon.minOrderAmount || 0) : 0

    if (total < minVal) {
      couponDiscount = 0
      couponVoided   = true
    }
  }

  // Recalculate tax on discounted amount (same 5% logic used at checkout)
  const taxable   = Math.max(0, total - couponDiscount)
  const tax       = Math.round(taxable * 0.05)
  const finalAmt  = taxable + tax + (order.shippingCharge || 0)

  return { finalAmount: finalAmt, couponVoided, couponDiscount, tax }
}


// ─────────────────────────────────────────────
// loadOrders
// ─────────────────────────────────────────────
const loadOrders = async (req, res) => {
  try {
    const page   = Math.max(1, parseInt(req.query.page) || 1)
    const search = (req.query.search || '').trim().replace(/^#/, '')
    const status = req.query.status || ''
    const sort   = req.query.sort   || 'newest'

    const query = {}
    if (status === 'Returned') {
      query.$or = [
        { returnStatus: { $ne: 'None' } },
        { 'itemStatuses.status': 'Returned' }
      ]
    } else if (status) {
      query.orderStatus = status
    }
    if (search) {
      query.$or = [
        { orderId: { $regex: search, $options: 'i' } },
        { 'shippingAddress.name': { $regex: search, $options: 'i' } },
        { 'shippingAddress.mobile': { $regex: search, $options: 'i' } },
        { 'items.productName': { $regex: search, $options: 'i' } },
      ]
    }

    const sortMap = {
      newest:      { createdAt: -1 },
      oldest:      { createdAt:  1 },
      amount_high: { finalAmount: -1 },
      amount_low:  { finalAmount:  1 },
    }
    const sortObj = sortMap[sort] || { createdAt: -1 }

    const total      = await Order.countDocuments(query)
    const totalPages = Math.ceil(total / LIMIT)

    const orders = await Order.find(query)
      .select('orderId createdAt shippingAddress items finalAmount paymentMethod paymentStatus orderStatus returnStatus itemStatuses cancelReason')
      .populate('userId', 'name email')
      .sort(sortObj)
      .skip((page - 1) * LIMIT)
      .limit(LIMIT)

    res.render('admin/orders', {
      orders,
      currentPage: page,
      totalPages,
      total,
      search,
      status,
      sort,
      user: req.session.admin || null,
    })
  } catch (err) {
    console.error('admin loadOrders error:', err)
    res.status(500).render('error', { message: 'Could not load orders.' })
  }
}


// ─────────────────────────────────────────────
// loadOrderDetail
// ─────────────────────────────────────────────
const loadOrderDetail = async (req, res) => {
  try {
    const order = await Order.findById(req.params.id)
      .populate('userId', 'name email phone')
      .lean()

    if (!order) return res.redirect('/admin/orders')

    res.render('admin/orderDetail', {
      order,
      user: req.session.admin || null,
    })
  } catch (err) {
    console.error('admin loadOrderDetail error:', err)
    res.redirect('/admin/orders')
  }
}


// ─────────────────────────────────────────────
// updateOrderStatus
// ─────────────────────────────────────────────
const updateOrderStatus = async (req, res) => {
  try {
    const { status } = req.body
    const allowed = ['Placed', 'Processing', 'Shipped', 'Delivered', 'Cancelled']

    if (!allowed.includes(status)) {
      return res.status(400).json({ success: false, message: 'Invalid status.' })
    }

    const order = await Order.findById(req.params.id)
    if (!order) return res.status(404).json({ success: false, message: 'Order not found.' })

    if (order.orderStatus === 'Cancelled') {
      return res.status(400).json({ success: false, message: 'Cancelled orders cannot be updated.' })
    }
    if (order.orderStatus === 'Delivered' && status !== 'Delivered') {
      return res.status(400).json({ success: false, message: 'Delivered orders cannot be changed.' })
    }

    const prev = order.orderStatus
    order.orderStatus = status

    if (status === 'Delivered') {
      order.deliveredAt  = new Date()
      order.paymentStatus = 'Paid'
    }
    if (status === 'Cancelled' && prev !== 'Cancelled') {
      order.cancelledAt  = new Date()
      order.cancelReason = order.cancelReason || 'Cancelled by admin'
      for (const item of order.items) {
        await Product.updateOne(
          { _id: item.productId, 'variants.shade': item.shade },
          { $inc: { 'variants.$.stock': item.quantity } }
        )
      }
    }

    await order.save()
    return res.json({ success: true, message: `Order status updated to ${status}.`, status })
  } catch (err) {
    console.error('updateOrderStatus error:', err)
    return res.status(500).json({ success: false, message: 'Something went wrong.' })
  }
}


// ─────────────────────────────────────────────
// cancelOrder  (whole order)
// ─────────────────────────────────────────────
const cancelOrder = async (req, res) => {
  try {
    const { reason } = req.body
    if (!reason) {
      return res.status(400).json({ success: false, message: 'Cancellation reason is required.' })
    }

    const order = await Order.findById(req.params.id)
    if (!order) return res.status(404).json({ success: false, message: 'Order not found.' })

    if (['Shipped', 'Delivered', 'Cancelled'].includes(order.orderStatus)) {
      return res.status(400).json({
        success: false,
        message: `Cannot cancel an order that is already ${order.orderStatus}.`
      })
    }

    order.orderStatus  = 'Cancelled'
    order.cancelReason = reason
    order.cancelledAt  = new Date()

    order.itemStatuses = order.itemStatuses || []
    for (const item of order.items) {
      const existing = order.itemStatuses.find(
        s => s.itemId && s.itemId.toString() === item._id.toString()
      )
      if (existing) {
        existing.status      = 'Cancelled'
        existing.cancelReason = reason
        existing.cancelledAt  = new Date()
      } else {
        order.itemStatuses.push({
          itemId:      item._id,
          status:      'Cancelled',
          cancelReason: reason,
          cancelledAt:  new Date()
        })
      }
      await Product.updateOne(
        { _id: item.productId, 'variants.shade': item.shade },
        { $inc: { 'variants.$.stock': item.quantity } }
      )
    }

    await order.save()
    return res.json({ success: true, message: 'Order cancelled successfully.' })
  } catch (error) {
    console.error('cancelOrder error', error)
    return res.status(500).json({ success: false, message: 'Something went wrong.' })
  }
}


// ─────────────────────────────────────────────
// cancelItem  (single item)
// ─────────────────────────────────────────────
const cancelItem = async (req, res) => {
  try {
    const { reason } = req.body
    if (!reason) {
      return res.status(400).json({ success: false, message: 'Cancellation reason is required.' })
    }

    const order = await Order.findById(req.params.id)
    if (!order) return res.status(404).json({ success: false, message: 'Order not found.' })

    if (['Shipped', 'Delivered', 'Cancelled'].includes(order.orderStatus)) {
      return res.status(400).json({
        success: false,
        message: `Cannot cancel items in a ${order.orderStatus} order.`
      })
    }

    // ── BLOCK: online paid orders with a coupon cannot do per-item cancel ──
    const isOnlinePaid = order.paymentMethod !== 'COD'
    const hasCoupon    = !!(order.couponCode && order.couponDiscount > 0)

    if (isOnlinePaid && hasCoupon) {
      return res.status(400).json({
        success: false,
        message: 'Per-item cancellation is not allowed for online orders with a coupon applied. Please cancel the entire order instead.'
      })
    }

    // Find the item
    const item = order.items.find(i => i._id.toString() === req.params.itemId)
    if (!item) return res.status(404).json({ success: false, message: 'Item not found in this order.' })

    // Mark item as cancelled
    order.itemStatuses = order.itemStatuses || []
    const existing = order.itemStatuses.find(
      s => s.itemId && s.itemId.toString() === req.params.itemId
    )

    if (existing) {
      if (existing.status === 'Cancelled') {
        return res.status(400).json({ success: false, message: 'Item is already cancelled.' })
      }
      existing.status       = 'Cancelled'
      existing.cancelReason = reason
      existing.cancelledAt  = new Date()
    } else {
      order.itemStatuses.push({
        itemId:      item._id,
        status:      'Cancelled',
        cancelReason: reason,
        cancelledAt:  new Date()
      })
    }

    // Restock
    await Product.updateOne(
      { _id: item.productId, 'variants.shade': item.shade },
      { $inc: { 'variants.$.stock': item.quantity } }
    )

    // ── COD + COUPON: recalculate amount ────────────────────────────────
    let couponVoidedMsg = ''
    if (order.paymentMethod === 'COD' && hasCoupon) {
      const { finalAmount, couponVoided, couponDiscount, tax } = await recalcCODAmount(order)
      order.finalAmount = finalAmount
      order.tax         = tax

      if (couponVoided) {
        order.couponDiscount = 0
        order.couponVoided   = true   // flag used by the EJS template
        couponVoidedMsg = ' Coupon was removed because remaining items no longer meet the minimum order value.'
      } else {
        order.couponDiscount = couponDiscount
      }
    }

    // ── If all items are now cancelled → cancel the whole order ──────────
    const allCancelled = order.items.every(i =>
      (order.itemStatuses || []).some(
        s => s.itemId && s.itemId.toString() === i._id.toString() && s.status === 'Cancelled'
      )
    )
    if (allCancelled) {
      order.orderStatus  = 'Cancelled'
      order.cancelReason = 'All items cancelled'
      order.cancelledAt  = new Date()
    }

    await order.save()

    const baseMsg = allCancelled
      ? 'Item cancelled. Order also marked as Cancelled since all items are cancelled.'
      : 'Item cancelled successfully.'

    return res.json({
      success: true,
      message: baseMsg + couponVoidedMsg
    })
  } catch (err) {
    console.error('cancelItem error:', err)
    return res.status(500).json({ success: false, message: 'Something went wrong.' })
  }
}


// ─────────────────────────────────────────────
// updateReturnStatus
// ─────────────────────────────────────────────
const updateReturnStatus = async (req, res) => {
  try {
    const { returnStatus, itemId } = req.body
    const allowed = ['Requested', 'Approved', 'Rejected', 'Completed']

    if (!allowed.includes(returnStatus)) {
      return res.status(400).json({ success: false, message: 'Invalid return status.' })
    }

    const order = await Order.findById(req.params.id)
    if (!order) {
      return res.status(404).json({ success: false, message: 'Order not found.' })
    }

    const isPaid = order.paymentMethod !== 'COD' || order.paymentStatus === 'Paid'

   // ── Only count items that were actually paid (not cancelled) ──
const paidItems = order.items.filter(item => {
  const s = (order.itemStatuses || []).find(
    st => st.itemId && st.itemId.toString() === item._id.toString()
  )
  return !s || s.status !== 'Cancelled'
})

const orderItemsTotal = paidItems.reduce(
  (sum, i) => sum + i.salePrice * i.quantity, 0
)

const calcItemRefund = (item) => {
  const itemBase = item.salePrice * item.quantity

  // Proportional coupon share — only from paid items total
  const couponShare = orderItemsTotal > 0
    ? Math.round((itemBase / orderItemsTotal) * (order.couponDiscount || 0))
    : 0

  // Proportional shipping share — only from paid items total  
  const shippingShare = orderItemsTotal > 0
    ? Math.round((itemBase / orderItemsTotal) * (order.shippingCharge || 0))
    : 0

  return itemBase - couponShare + shippingShare
}

    // ── SINGLE ITEM RETURN ──────────────────────────────────────────────
    if (itemId) {
      const itemStatus = order.itemStatuses.find(
        s => s.itemId && s.itemId.toString() === itemId.toString()
      )
      if (!itemStatus || itemStatus.status !== 'Returned') {
        return res.status(400).json({ success: false, message: 'No return request found for this item.' })
      }

      const wasAlreadyApproved = ['Approved', 'Completed'].includes(itemStatus.returnStatus)
      itemStatus.returnStatus  = returnStatus

      if (['Approved', 'Completed'].includes(returnStatus) && !wasAlreadyApproved) {
        const item = order.items.find(i => i._id.toString() === itemId.toString())
        if (item) {
          await Product.updateOne(
            { _id: item.productId, 'variants.shade': item.shade },
            { $inc: { 'variants.$.stock': item.quantity } }
          )
          if (isPaid) {
            const refundAmount = calcItemRefund(item)
            await creditWallet(
              order.userId,
              refundAmount,
              `Refund for returned item "${item.productName}" in order #${order.orderId}`,
              order._id
            )
          }
        }
      }
          // ── Auto-update order-level returnStatus when all item returns are done ──


const allItemReturns = order.itemStatuses.filter(s => s.status === 'Returned')

if (allItemReturns.length > 0) {
  const allCompleted = allItemReturns.every(s => s.returnStatus === 'Completed')
  const allRejected  = allItemReturns.every(s => s.returnStatus === 'Rejected')
  const anyApproved  = allItemReturns.some(s => s.returnStatus === 'Approved')
  const anyRequested = allItemReturns.some(s => s.returnStatus === 'Requested')

  if (allCompleted)      order.returnStatus = 'Completed'
  else if (allRejected)  order.returnStatus = 'Rejected'
  else if (anyApproved)  order.returnStatus = 'Approved'
  else if (anyRequested) order.returnStatus = 'Requested'
}

await order.save()
return res.json({ success: true, message: `Item return marked as ${returnStatus}.` })
    }
    // ── FULL ORDER RETURN ───────────────────────────────────────────────
    if (!order.returnStatus || order.returnStatus === 'None') {
      return res.status(400).json({ success: false, message: 'No return request found for this order.' })
    }

    const wasAlreadyApproved = ['Approved', 'Completed'].includes(order.returnStatus)
    order.returnStatus       = returnStatus

    if (['Approved', 'Completed'].includes(returnStatus) && !wasAlreadyApproved) {
      const returnedItems = order.items.filter(item => {
        const s = order.itemStatuses.find(
          st => st.itemId && st.itemId.toString() === item._id.toString()
        )
        return s && s.status === 'Returned'
      })

      for (const item of returnedItems) {
        await Product.updateOne(
          { _id: item.productId, 'variants.shade': item.shade },
          { $inc: { 'variants.$.stock': item.quantity } }
        )
      }

      if (isPaid) {
        let totalRefund
        if (returnedItems.length === order.items.length) {
          totalRefund = order.finalAmount
        } else {
          totalRefund = returnedItems.reduce((sum, item) => sum + calcItemRefund(item), 0)
        }
        if (totalRefund > 0) {
          await creditWallet(
            order.userId,
            totalRefund,
            `Refund for returned order #${order.orderId}`,
            order._id
          )
        }
      }
    }

    await order.save()
    return res.json({ success: true, message: `Return status updated to ${returnStatus}.` })
  } catch (err) {
    console.error('updateReturnStatus error:', err.message, err.stack)
    return res.status(500).json({ success: false, message: 'Something went wrong.' })
  }
}


module.exports = {
  loadOrders,
  loadOrderDetail,
  updateOrderStatus,
  cancelOrder,
  cancelItem,
  updateReturnStatus,
}