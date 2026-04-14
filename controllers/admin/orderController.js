const Order = require('../../models/user/orderModel')
const User = require('../../models/user/userModel')
const Product = require('../../models/user/productModel')
const { creditWallet } = require('../user/walletController')  

const LIMIT = 5


const loadOrders = async (req, res) => {
  try {
    const page = Math.max(1, parseInt(req.query.page) || 1)
    const search = (req.query.search || '').trim().replace(/^#/, '')
    const status = req.query.status || ''
    const sort = req.query.sort || 'newest'

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
      newest: { createdAt: -1 },
      oldest: { createdAt: 1 },
      amount_high: { finalAmount: -1 },
      amount_low: { finalAmount: 1 },
    }
    const sortObj = sortMap[sort] || { createdAt: -1 }

    const total = await Order.countDocuments(query)
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
      order.deliveredAt = new Date()
      order.paymentStatus = 'Paid'
    }
    if (status === 'Cancelled' && prev !== 'Cancelled') {
      order.cancelledAt = new Date()
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
        message: `Cannot cancel an order that is already ${order.orderStatus}.`  // ✅ fixed template literal
      })
    }

    order.orderStatus = 'Cancelled'
    order.cancelReason = reason
    order.cancelledAt = new Date()

    order.itemStatuses = order.itemStatuses || []
    for (const item of order.items) {
      const existing = order.itemStatuses.find(
        s => s.itemId && s.itemId.toString() === item._id.toString()
      )

      if (existing) {
        existing.status = 'Cancelled'
        existing.cancelReason = reason
        existing.cancelledAt = new Date()
      } else {
        order.itemStatuses.push({
          itemId: item._id,
          status: 'Cancelled',
          cancelReason: reason,
          cancelledAt: new Date()
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


const cancelItem = async (req, res) => {
  try {
    const { reason } = req.body
    if (!reason) {
      return res.status(400).json({ success: false, message: 'Cancellation reason is required.' })
    }

    const order = await Order.findById(req.params.id)
    if (!order) return res.status(404).json({ success: false, message: 'Order not found.' })

    if (['Shipped', 'Delivered', 'Cancelled'].includes(order.orderStatus)) {
      return res.status(400).json({ success: false, message: `Cannot cancel items in a ${order.orderStatus} order.` })
    }

    const item = order.items.find(i => i._id.toString() === req.params.itemId)
    if (!item) return res.status(404).json({ success: false, message: 'Item not found in this order.' })

    order.itemStatuses = order.itemStatuses || []
    const existing = order.itemStatuses.find(
      s => s.itemId && s.itemId.toString() === req.params.itemId
    )

    if (existing) {
      if (existing.status === 'Cancelled') {
        return res.status(400).json({ success: false, message: 'Item is already cancelled.' })
      }
      existing.status = 'Cancelled'
      existing.cancelReason = reason
      existing.cancelledAt = new Date()
    } else {
      order.itemStatuses.push({
        itemId: item._id,
        status: 'Cancelled',
        cancelReason: reason,
        cancelledAt: new Date()
      })
    }

    await Product.updateOne(
      { _id: item.productId, 'variants.shade': item.shade },
      { $inc: { 'variants.$.stock': item.quantity } }
    )

    const allCancelled = order.items.every(i =>
      (order.itemStatuses || []).some(
        s => s.itemId && s.itemId.toString() === i._id.toString() && s.status === 'Cancelled'
      )
    )
    if (allCancelled) {
      order.orderStatus = 'Cancelled'
      order.cancelReason = 'All items cancelled'
      order.cancelledAt = new Date()
    }

    await order.save()
    return res.json({
      success: true,
      message: allCancelled
        ? 'Item cancelled. Order also marked as Cancelled since all items are cancelled.'
        : 'Item cancelled successfully.'
    })
  } catch (err) {
    console.error('cancelItem error:', err)
    return res.status(500).json({ success: false, message: 'Something went wrong.' })
  }
}

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

    // Helper: total of all items' salePrice * qty (used for coupon proportioning)
    const orderItemsTotal = order.items.reduce(
      (sum, i) => sum + i.salePrice * i.quantity, 0
    )

    // Helper: calculate proportional refund for a single item
    const calcItemRefund = (item) => {
      const itemBase = item.salePrice * item.quantity

      const couponShare = orderItemsTotal > 0
        ? Math.round((itemBase / orderItemsTotal) * (order.couponDiscount || 0))
        : 0

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
      itemStatus.returnStatus = returnStatus

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

      await order.save()
      return res.json({ success: true, message: `Item return marked as ${returnStatus}.` })
    }

    // ── FULL ORDER RETURN ───────────────────────────────────────────────
    if (!order.returnStatus || order.returnStatus === 'None') {
      return res.status(400).json({ success: false, message: 'No return request found for this order.' })
    }

    const wasAlreadyApproved = ['Approved', 'Completed'].includes(order.returnStatus)
    order.returnStatus = returnStatus

    if (['Approved', 'Completed'].includes(returnStatus) && !wasAlreadyApproved) {
      // Collect returned items
      const returnedItems = order.items.filter(item => {
        const s = order.itemStatuses.find(
          st => st.itemId && st.itemId.toString() === item._id.toString()
        )
        return s && s.status === 'Returned'
      })

      // Restock
      for (const item of returnedItems) {
        await Product.updateOne(
          { _id: item.productId, 'variants.shade': item.shade },
          { $inc: { 'variants.$.stock': item.quantity } }
        )
      }

      if (isPaid) {
        let totalRefund

        if (returnedItems.length === order.items.length) {
          // ALL items returned → refund exactly what was paid
          totalRefund = order.finalAmount
        } else {
          // PARTIAL return → proportional refund per returned item
          totalRefund = returnedItems.reduce(
            (sum, item) => sum + calcItemRefund(item), 0
          )
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
  updateReturnStatus
}