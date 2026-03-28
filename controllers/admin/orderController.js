const Order   = require('../../models/user/orderModel')
const User    = require('../../models/user/userModel')
const Product = require('../../models/user/productModel')

const LIMIT = 5


const loadOrders = async (req, res) => {
  try {
    const page    = Math.max(1, parseInt(req.query.page) || 1)
    const search  = (req.query.search || '').trim().replace(/^#/,'')
    const status  = req.query.status  || ''
    const sort    = req.query.sort    || 'newest'

    // build query
    const query = {}
    if (status) query.orderStatus = status
    if (search) {
      query.$or = [
        { orderId:             { $regex: search, $options: 'i' } },
        { 'shippingAddress.name':   { $regex: search, $options: 'i' } },
        { 'shippingAddress.mobile': { $regex: search, $options: 'i' } },
        { 'items.productName':      { $regex: search, $options: 'i' } },
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
      .populate('userId', 'name email')
      .sort(sortObj)
      .skip((page - 1) * LIMIT)
      .limit(LIMIT)
      .lean()

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

    // prevent going backwards
    const flow = ['Placed', 'Processing', 'Shipped', 'Delivered']
    const curIdx = flow.indexOf(order.orderStatus)
    const newIdx = flow.indexOf(status)

    if (order.orderStatus === 'Cancelled') {
      return res.status(400).json({ success: false, message: 'Cancelled orders cannot be updated.' })
    }
    if (order.orderStatus === 'Delivered' && status !== 'Delivered') {
      return res.status(400).json({ success: false, message: 'Delivered orders cannot be changed.' })
    }

    const prev = order.orderStatus
    order.orderStatus = status

    if (status === 'Delivered') {
      order.deliveredAt   = new Date()
      order.paymentStatus = 'Paid'
    }
    if (status === 'Cancelled' && prev !== 'Cancelled') {
      order.cancelledAt  = new Date()
      order.cancelReason = order.cancelReason || 'Cancelled by admin'
      // restore stock
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

    if (itemId) {
      const itemStatus = order.itemStatuses.find(
        s => s.itemId && s.itemId.toString() === itemId.toString()
      )
      
      if (!itemStatus || itemStatus.status !== 'Returned') {
        return res.status(400).json({ success: false, message: 'No return request found for this item.' })
      }

      const wasAlreadyRestored = ['Approved', 'Completed'].includes(itemStatus.returnStatus)
      itemStatus.returnStatus = returnStatus

  
      if (['Approved', 'Completed'].includes(returnStatus) && !wasAlreadyRestored) {
        const item = order.items.find(i => i._id.toString() === itemId.toString())
        if (item) {
          await Product.updateOne(
            { _id: item.productId, 'variants.shade': item.shade },
            { $inc: { 'variants.$.stock': item.quantity } }
          )
        }
      }

      await order.save()
      return res.json({ success: true, message: `Item return marked as ${returnStatus}.` })
    }

    if (!order.returnStatus || order.returnStatus === 'None') {
      return res.status(400).json({ success: false, message: 'No return request found for this order.' })
    }

    const wasAlreadyRestored = ['Approved', 'Completed'].includes(order.returnStatus)
    order.returnStatus = returnStatus

    
    if (['Approved', 'Completed'].includes(returnStatus) && !wasAlreadyRestored) {
      for (const item of order.items) {
        const s = order.itemStatuses.find(
          st => st.itemId && st.itemId.toString() === item._id.toString()
        )
       
        if (s && s.status === 'Returned') {
          await Product.updateOne(
            { _id: item.productId, 'variants.shade': item.shade },
            { $inc: { 'variants.$.stock': item.quantity } }
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
   updateReturnStatus
  }