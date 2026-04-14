const Order   = require('../../models/user/orderModel')
const Product = require('../../models/user/productModel')

// ── helpers ──────────────────────────────────────────────────────────────────

function getDateRange(filter) {
  const now   = new Date()
  const start = new Date()

  switch (filter) {
    case 'today':
      start.setHours(0, 0, 0, 0)
      break
    case 'week':
      start.setDate(now.getDate() - 6)
      start.setHours(0, 0, 0, 0)
      break
    case 'month':
      start.setDate(1)
      start.setHours(0, 0, 0, 0)
      break
    case 'year':
      start.setMonth(0, 1)
      start.setHours(0, 0, 0, 0)
      break
    case 'custom': // handled by caller
      return null
    default:       // 'monthly' — last 12 months default
      start.setMonth(now.getMonth() - 11)
      start.setDate(1)
      start.setHours(0, 0, 0, 0)
  }

  return { $gte: start, $lte: now }
}

// ── main controller ───────────────────────────────────────────────────────────

const loadDashboard = async (req, res) => {
  try {
    const filter    = req.query.filter    || 'monthly'
    const chartType = req.query.chartType || 'revenue'

    // custom date range
    let dateFilter
    if (filter === 'custom' && req.query.from && req.query.to) {
      dateFilter = {
        $gte: new Date(req.query.from),
        $lte: new Date(new Date(req.query.to).setHours(23, 59, 59, 999))
      }
    } else {
      dateFilter = getDateRange(filter)
    }

    const matchStage = {
      orderStatus: { $nin: ['Cancelled', 'Returned'] },
      ...(dateFilter && { createdAt: dateFilter })
    }

    // ── summary cards ─────────────────────────────────────────────────────────
    const [summaryArr] = await Order.aggregate([
      { $match: matchStage },
      {
        $group: {
          _id:          null,
          totalRevenue: { $sum: '$finalAmount' },
          totalOrders:  { $sum: 1 },
          totalDiscount:{ $sum: '$totalDiscount' }
        }
      }
    ])

    const summary = summaryArr || { totalRevenue: 0, totalOrders: 0, totalDiscount: 0 }

    // total users (all time)
    const User = require('../../models/user/userModel')
    summary.totalUsers = await User.countDocuments({ isDeleted: false })

    // ── chart data ─
    // Group format depends on filter
    let groupId
    if (filter === 'today') {
      groupId = { hour: { $hour: '$createdAt' } }
    } else if (filter === 'week') {
      groupId = { dayOfWeek: { $dayOfWeek: '$createdAt' }, date: { $dateToString: { format: '%Y-%m-%d', date: '$createdAt' } } }
    } else if (filter === 'month') {
      groupId = { day: { $dayOfMonth: '$createdAt' } }
    } else {
      // year / monthly / custom → group by month
      groupId = { year: { $year: '$createdAt' }, month: { $month: '$createdAt' } }
    }

    const chartRaw = await Order.aggregate([
      { $match: matchStage },
      {
        $group: {
          _id:      groupId,
          revenue:  { $sum: '$finalAmount' },
          orders:   { $sum: 1 },
          discount: { $sum: '$totalDiscount' }
        }
      },
      { $sort: { '_id.year': 1, '_id.month': 1, '_id.day': 1, '_id.hour': 1, '_id.date': 1 } }
    ])

    // build readable labels
    const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
    const dayNames   = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat']

    const chartLabels  = []
    const chartRevenue = []
    const chartOrders  = []
    const chartDiscount= []

    if (filter === 'today') {
      // fill 24 hours
      const byHour = {}
      chartRaw.forEach(r => { byHour[r._id.hour] = r })
      for (let h = 0; h < 24; h++) {
        chartLabels.push(h === 0 ? '12am' : h < 12 ? `${h}am` : h === 12 ? '12pm' : `${h-12}pm`)
        chartRevenue.push((byHour[h]?.revenue  || 0).toFixed(2))
        chartOrders.push(byHour[h]?.orders   || 0)
        chartDiscount.push((byHour[h]?.discount || 0).toFixed(2))
      }
    } else if (filter === 'week') {
      chartRaw.forEach(r => {
        chartLabels.push(r._id.date)
        chartRevenue.push(r.revenue.toFixed(2))
        chartOrders.push(r.orders)
        chartDiscount.push(r.discount.toFixed(2))
      })
    } else if (filter === 'month') {
      const now      = new Date()
      const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate()
      const byDay    = {}
      chartRaw.forEach(r => { byDay[r._id.day] = r })
      for (let d = 1; d <= daysInMonth; d++) {
        chartLabels.push(d)
        chartRevenue.push((byDay[d]?.revenue  || 0).toFixed(2))
        chartOrders.push(byDay[d]?.orders   || 0)
        chartDiscount.push((byDay[d]?.discount || 0).toFixed(2))
      }
    } else {
      // monthly grouping
      chartRaw.forEach(r => {
        chartLabels.push(`${monthNames[r._id.month - 1]} ${r._id.year}`)
        chartRevenue.push(r.revenue.toFixed(2))
        chartOrders.push(r.orders)
        chartDiscount.push(r.discount.toFixed(2))
      })
    }

    //  best selling products (top 10) 
    const bestProducts = await Order.aggregate([
      { $match: matchStage },
      { $unwind: '$items' },
      {
        $group: {
          _id:        '$items.productId',
          name:       { $first: '$items.productName' },
          image:      { $first: '$items.productImage' },
          totalQty:   { $sum: '$items.quantity' },
          totalRevenue:{ $sum: { $multiply: ['$items.quantity', '$items.salePrice'] } }
        }
      },
      { $sort: { totalQty: -1 } },
      { $limit: 10}
    ])
  
    //  best selling categories (top 10) 
    const bestCategories = await Order.aggregate([
      { $match: matchStage },
      { $unwind: '$items' },
      {
        $lookup: {
          from:         'products',
          localField:   'items.productId',
          foreignField: '_id',
          as:           'prod'
        }
      },
      { $unwind: { path: '$prod', preserveNullAndEmptyArrays: true } },
      {
        $lookup: {
          from:         'categories',
          localField:   'prod.categoryId',
          foreignField: '_id',
          as:           'cat'
        }
      },
      { $unwind: { path: '$cat', preserveNullAndEmptyArrays: true } },
      {
        $group: {
          _id:         '$cat._id',
          name:        { $first: '$cat.name' },
          totalQty:    { $sum: '$items.quantity' },
          totalRevenue:{ $sum: { $multiply: ['$items.quantity', '$items.salePrice'] } }
        }
      },
      { $match: { _id: { $ne: null } } },
      { $sort: { totalQty: -1 } },
      { $limit: 10 }
    ])

    // ── best selling brands (top 10) ─────────────────────────────────────────
    // Brand = first word of product name (adjust if you have a brand field)
    const bestBrands = await Order.aggregate([
      { $match: matchStage },
      { $unwind: '$items' },
      {
        $group: {
          _id:         { $arrayElemAt: [{ $split: ['$items.productName', ' '] }, 0] },
          totalQty:    { $sum: '$items.quantity' },
          totalRevenue:{ $sum: { $multiply: ['$items.quantity', '$items.salePrice'] } }
        }
      },
      { $sort: { totalQty: -1 } },
      { $limit: 10 }
    ])

    // ── ledger: last 10 transactions ──────────────────────────────────────────
    const ledger = await Order.find(matchStage)
      .populate('userId', 'fullName email')
      .sort({ createdAt: -1 })
      .limit(5)
      .lean()

    res.render('admin/dashboard', {
      summary,
      chartLabels:   JSON.stringify(chartLabels),
      chartRevenue:  JSON.stringify(chartRevenue),
      chartOrders:   JSON.stringify(chartOrders),
      chartDiscount: JSON.stringify(chartDiscount),
      chartType,
      filter,
      fromDate: req.query.from || '',
      toDate:   req.query.to   || '',
      bestProducts,
      bestCategories,
      bestBrands,
      ledger
    })

  } catch (err) {
    console.error('loadDashboard error:', err)
    res.status(500).send('Dashboard error: ' + err.message)
  }
}

module.exports = { loadDashboard }