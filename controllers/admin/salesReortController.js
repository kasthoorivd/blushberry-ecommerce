const Order  = require('../../models/user/orderModel')
const ExcelJS = require('exceljs')
const PDFDocument = require('pdfkit')

// ─── helpers ────────────────────────────────────────────────────────────────

function getDateRange(type, from, to) {
  const now = new Date()
  let start, end

  if (type === 'daily') {
    start = new Date(now); start.setHours(0, 0, 0, 0)
    end   = new Date(now); end.setHours(23, 59, 59, 999)
  } else if (type === 'weekly') {
    const day = now.getDay()
    start = new Date(now); start.setDate(now.getDate() - day); start.setHours(0, 0, 0, 0)
    end   = new Date(now); end.setHours(23, 59, 59, 999)
  } else if (type === 'yearly') {
    start = new Date(now.getFullYear(), 0, 1)
    end   = new Date(now.getFullYear(), 11, 31, 23, 59, 59, 999)
  } else {
    // custom
    start = from ? new Date(from) : new Date(0)
    end   = to   ? new Date(new Date(to).setHours(23, 59, 59, 999)) : new Date()
  }
  return { start, end }
}

// ─── main loader ─────────────────────────────────────────────────────────────

const loadSalesReport = async (req, res) => {
  try {
    const {
      type   = 'daily',
      from   = '',
      to     = '',
      coupon = '',
      page   = 1
    } = req.query

    const LIMIT = 5
    const currentPage = Math.max(1, parseInt(page))
    const { start, end } = getDateRange(type, from, to)

    // base match
    const match = {
      createdAt:   { $gte: start, $lte: end },
      orderStatus: { $nin: ['Cancelled'] }
    }
    if (coupon) match.couponCode = coupon.toUpperCase()

    // ── summary aggregation ──────────────────────────────────────────────────
    const [summary] = await Order.aggregate([
      { $match: match },
      {
        $group: {
          _id:            null,
          totalOrders:    { $sum: 1 },
          grossRevenue:   { $sum: '$subtotal' },
          totalDiscount:  { $sum: '$totalDiscount' },
          couponDiscount: { $sum: '$couponDiscount' },
          netRevenue:     { $sum: '$finalAmount' }
        }
      }
    ])

    const stats = summary || {
      totalOrders: 0, grossRevenue: 0,
      totalDiscount: 0, couponDiscount: 0, netRevenue: 0
    }

    // ── coupon breakdown ─────────────────────────────────────────────────────
    const couponBreakdown = await Order.aggregate([
      { $match: { ...match, couponCode: { $ne: null } } },
      {
        $group: {
          _id:           '$couponCode',
          uses:          { $sum: 1 },
          totalDeducted: { $sum: '$couponDiscount' },
          netRevenue:    { $sum: '$finalAmount' }
        }
      },
      { $sort: { totalDeducted: -1 } }
    ])

    // ── chart data ───────────────────────────────────────────────────────────
    let groupId, dateFormat
    if (type === 'daily') {
      groupId = { hour: { $hour: '$createdAt' } }
    } else if (type === 'weekly') {
      groupId = { dayOfWeek: { $dayOfWeek: '$createdAt' } }
    } else {
      groupId = { month: { $month: '$createdAt' } }
    }

    const chartRaw = await Order.aggregate([
      { $match: match },
      {
        $group: {
          _id:     groupId,
          gross:   { $sum: '$subtotal' },
          net:     { $sum: '$finalAmount' },
          discount:{ $sum: '$totalDiscount' }
        }
      },
      { $sort: { '_id.hour': 1, '_id.dayOfWeek': 1, '_id.month': 1 } }
    ])

    const DAYS  = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat']
    const MONTHS= ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

    let chartLabels, chartGross, chartNet, chartDiscount
    if (type === 'daily') {
      chartLabels   = chartRaw.map(d => `${d._id.hour}:00`)
      chartGross    = chartRaw.map(d => d.gross)
      chartNet      = chartRaw.map(d => d.net)
      chartDiscount = chartRaw.map(d => d.discount)
    } else if (type === 'weekly') {
      chartLabels   = chartRaw.map(d => DAYS[(d._id.dayOfWeek - 1 + 7) % 7])
      chartGross    = chartRaw.map(d => d.gross)
      chartNet      = chartRaw.map(d => d.net)
      chartDiscount = chartRaw.map(d => d.discount)
    } else {
      chartLabels   = chartRaw.map(d => MONTHS[d._id.month - 1])
      chartGross    = chartRaw.map(d => d.gross)
      chartNet      = chartRaw.map(d => d.net)
      chartDiscount = chartRaw.map(d => d.discount)
    }

    // ── paginated order list ─────────────────────────────────────────────────
    const total  = await Order.countDocuments(match)
    const orders = await Order.find(match)
      .populate('userId', 'name email')
      .sort({ createdAt: -1 })
      .skip((currentPage - 1) * LIMIT)
      .limit(LIMIT)
      .lean()

    // ── distinct coupons for filter dropdown ─────────────────────────────────
    const allCoupons = await Order.distinct('couponCode', {
      createdAt: { $gte: new Date(Date.now() - 365 * 24 * 60 * 60 * 1000) },
      couponCode: { $ne: null }
    })

    res.render('admin/salesReport', {
      user:           req.session.admin || null,
      stats,
      orders,
      couponBreakdown,
      allCoupons,
      chartLabels:    JSON.stringify(chartLabels),
      chartGross:     JSON.stringify(chartGross),
      chartNet:       JSON.stringify(chartNet),
      chartDiscount:  JSON.stringify(chartDiscount),
      // filter state
      type, from, to, coupon,
      currentPage,
      totalPages: Math.ceil(total / LIMIT),
      total,
      start, end
    })

  } catch (err) {
    console.error('loadSalesReport error:', err)
    res.status(500).render('error', { message: 'Could not load sales report.' })
  }
}

// ─── PDF download ─────────────────────────────────────────────────────────────

const downloadPDF = async (req, res) => {
  try {
    const { type = 'daily', from = '', to = '', coupon = '' } = req.query
    const { start, end } = getDateRange(type, from, to)

    const match = {
      createdAt:   { $gte: start, $lte: end },
      orderStatus: { $nin: ['Cancelled'] }
    }
    if (coupon) match.couponCode = coupon.toUpperCase()

    const [summary] = await Order.aggregate([
      { $match: match },
      {
        $group: {
          _id:            null,
          totalOrders:    { $sum: 1 },
          grossRevenue:   { $sum: '$subtotal' },
          totalDiscount:  { $sum: '$totalDiscount' },
          couponDiscount: { $sum: '$couponDiscount' },
          netRevenue:     { $sum: '$finalAmount' }
        }
      }
    ])

    const stats = summary || { totalOrders: 0, grossRevenue: 0, totalDiscount: 0, couponDiscount: 0, netRevenue: 0 }

    const orders = await Order.find(match)
      .populate('userId', 'name email')
      .sort({ createdAt: -1 })
      .limit(200)
      .lean()

    const doc = new PDFDocument({ margin: 40, size: 'A4' })
    res.setHeader('Content-Type', 'application/pdf')
    res.setHeader('Content-Disposition', `attachment; filename="sales-report-${type}-${Date.now()}.pdf"`)
    doc.pipe(res)

    // header
    doc.fontSize(20).font('Helvetica-Bold').text('Blush-Berry — Sales Report', { align: 'center' })
    doc.fontSize(10).font('Helvetica').fillColor('#888')
       .text(`Period: ${start.toDateString()} → ${end.toDateString()}  |  Generated: ${new Date().toLocaleString()}`, { align: 'center' })
    doc.moveDown(1.2)

    // summary boxes
    doc.fontSize(11).font('Helvetica-Bold').fillColor('#333').text('Summary', { underline: true })
    doc.moveDown(0.4)
    const summaryRows = [
      ['Total Orders',    stats.totalOrders],
      ['Gross Revenue',   `₹${stats.grossRevenue.toLocaleString('en-IN')}`],
      ['Total Discounts', `₹${stats.totalDiscount.toLocaleString('en-IN')}`],
      ['Coupon Discounts',`₹${stats.couponDiscount.toLocaleString('en-IN')}`],
      ['Net Revenue',     `₹${stats.netRevenue.toLocaleString('en-IN')}`],
    ]
    summaryRows.forEach(([label, val]) => {
      doc.font('Helvetica').fontSize(10).fillColor('#555').text(label, { continued: true, width: 180 })
      doc.font('Helvetica-Bold').fillColor('#111').text(String(val))
    })

    doc.moveDown(1.2)

    // order table header
    doc.fontSize(11).font('Helvetica-Bold').fillColor('#333').text('Order Details', { underline: true })
    doc.moveDown(0.4)

    const cols = [40, 110, 210, 300, 370, 440, 510]
    const headers = ['#', 'Order ID', 'Customer', 'Gross', 'Discount', 'Net', 'Status']
    doc.fontSize(9).font('Helvetica-Bold').fillColor('#fff')
    doc.rect(40, doc.y, 515, 16).fill('#c93060')
    const hY = doc.y - 14
    headers.forEach((h, i) => doc.fillColor('#fff').text(h, cols[i], hY, { width: cols[i + 1] - cols[i] - 4 }))
    doc.moveDown(0.1)

    orders.forEach((o, idx) => {
      const rowY = doc.y
      if (rowY > 750) { doc.addPage(); }
      const bg = idx % 2 === 0 ? '#fdf0f4' : '#fff'
      doc.rect(40, doc.y, 515, 14).fill(bg)
      const y = doc.y - 12
      doc.fontSize(8).font('Helvetica').fillColor('#222')
      doc.text(String(idx + 1),                         cols[0], y, { width: 65 })
      doc.text(o.orderId || '',                         cols[1], y, { width: 95 })
      doc.text(o.userId?.name || o.shippingAddress?.name || '—', cols[2], y, { width: 85 })
      doc.text(`₹${(o.subtotal||0).toLocaleString('en-IN')}`,    cols[3], y, { width: 65 })
      doc.text(`₹${(o.totalDiscount||0).toLocaleString('en-IN')}`,cols[4], y, { width: 65 })
      doc.text(`₹${(o.finalAmount||0).toLocaleString('en-IN')}`, cols[5], y, { width: 65 })
      doc.text(o.orderStatus || '',                     cols[6], y, { width: 55 })
      doc.moveDown(0.05)
    })

    doc.end()
  } catch (err) {
    console.error('downloadPDF error:', err)
    res.status(500).json({ success: false, message: 'Could not generate PDF.' })
  }
}

// ─── Excel download ───────────────────────────────────────────────────────────

const downloadExcel = async (req, res) => {
  try {
    const { type = 'daily', from = '', to = '', coupon = '' } = req.query
    const { start, end } = getDateRange(type, from, to)

    const match = {
      createdAt:   { $gte: start, $lte: end },
      orderStatus: { $nin: ['Cancelled'] }
    }
    if (coupon) match.couponCode = coupon.toUpperCase()

    const [summary] = await Order.aggregate([
      { $match: match },
      {
        $group: {
          _id:            null,
          totalOrders:    { $sum: 1 },
          grossRevenue:   { $sum: '$subtotal' },
          totalDiscount:  { $sum: '$totalDiscount' },
          couponDiscount: { $sum: '$couponDiscount' },
          netRevenue:     { $sum: '$finalAmount' }
        }
      }
    ])

    const stats = summary || { totalOrders: 0, grossRevenue: 0, totalDiscount: 0, couponDiscount: 0, netRevenue: 0 }

    const orders = await Order.find(match)
      .populate('userId', 'name email')
      .sort({ createdAt: -1 })
      .limit(5000)
      .lean()

    const couponBreakdown = await Order.aggregate([
      { $match: { ...match, couponCode: { $ne: null } } },
      {
        $group: {
          _id:           '$couponCode',
          uses:          { $sum: 1 },
          totalDeducted: { $sum: '$couponDiscount' },
          netRevenue:    { $sum: '$finalAmount' }
        }
      },
      { $sort: { totalDeducted: -1 } }
    ])

    const wb = new ExcelJS.Workbook()
    wb.creator = 'Blush-Berry'

    // ── Sheet 1: Summary ─────────────────────────────────────────────────────
    const ws1 = wb.addWorksheet('Summary')
    ws1.columns = [
      { header: 'Metric', key: 'metric', width: 28 },
      { header: 'Value',  key: 'value',  width: 22 }
    ]
    ws1.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } }
    ws1.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC93060' } }

    ws1.addRow({ metric: 'Report Period',     value: `${start.toDateString()} → ${end.toDateString()}` })
    ws1.addRow({ metric: 'Total Orders',      value: stats.totalOrders })
    ws1.addRow({ metric: 'Gross Revenue (₹)', value: stats.grossRevenue })
    ws1.addRow({ metric: 'Total Discounts (₹)',value: stats.totalDiscount })
    ws1.addRow({ metric: 'Coupon Discounts (₹)',value: stats.couponDiscount })
    ws1.addRow({ metric: 'Net Revenue (₹)',   value: stats.netRevenue })

    // ── Sheet 2: Orders ──────────────────────────────────────────────────────
    const ws2 = wb.addWorksheet('Orders')
    ws2.columns = [
      { header: 'Order ID',      key: 'orderId',      width: 22 },
      { header: 'Date',          key: 'date',         width: 18 },
      { header: 'Customer',      key: 'customer',     width: 22 },
      { header: 'Email',         key: 'email',        width: 26 },
      { header: 'Payment',       key: 'payment',      width: 14 },
      { header: 'Gross (₹)',     key: 'gross',        width: 14 },
      { header: 'Coupon',        key: 'coupon',       width: 14 },
      { header: 'Discount (₹)',  key: 'discount',     width: 16 },
      { header: 'Net (₹)',       key: 'net',          width: 14 },
      { header: 'Status',        key: 'status',       width: 14 },
    ]
    ws2.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } }
    ws2.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC93060' } }

    orders.forEach((o, idx) => {
      const row = ws2.addRow({
        orderId:  o.orderId || '',
        date:     new Date(o.createdAt).toLocaleDateString('en-IN'),
        customer: o.userId?.name || o.shippingAddress?.name || '—',
        email:    o.userId?.email || o.shippingAddress?.email || '—',
        payment:  o.paymentMethod || '',
        gross:    o.subtotal || 0,
        coupon:   o.couponCode || '—',
        discount: o.totalDiscount || 0,
        net:      o.finalAmount || 0,
        status:   o.orderStatus || '',
      })
      if (idx % 2 === 0) {
        row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDF0F4' } }
      }
    })

    // ── Sheet 3: Coupons ─────────────────────────────────────────────────────
    const ws3 = wb.addWorksheet('Coupon Performance')
    ws3.columns = [
      { header: 'Coupon Code',     key: 'code',     width: 18 },
      { header: 'Uses',            key: 'uses',     width: 10 },
      { header: 'Total Deducted (₹)', key: 'deducted', width: 22 },
      { header: 'Net Revenue (₹)', key: 'revenue',  width: 20 },
    ]
    ws3.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } }
    ws3.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC93060' } }

    couponBreakdown.forEach(c => {
      ws3.addRow({ code: c._id, uses: c.uses, deducted: c.totalDeducted, revenue: c.netRevenue })
    })

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    res.setHeader('Content-Disposition', `attachment; filename="sales-report-${type}-${Date.now()}.xlsx"`)
    await wb.xlsx.write(res)
    res.end()

  } catch (err) {
    console.error('downloadExcel error:', err)
    res.status(500).json({ success: false, message: 'Could not generate Excel.' })
  }
}

module.exports = { loadSalesReport, downloadPDF, downloadExcel }