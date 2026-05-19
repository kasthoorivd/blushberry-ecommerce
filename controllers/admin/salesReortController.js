const Order = require('../../models/user/orderModel')
const ExcelJS = require('exceljs')
const PDFDocument = require('pdfkit')
const { HttpStatus } = require('../../utils/statusCode')

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

function getDateRange(type, from, to) {
  const now = new Date()
  let start, end

  if (type === 'daily') {
    start = new Date(now); start.setHours(0, 0, 0, 0)
    end = new Date(now); end.setHours(23, 59, 59, 999)
  } else if (type === 'weekly') {
    const day = now.getDay()
    start = new Date(now); start.setDate(now.getDate() - day); start.setHours(0, 0, 0, 0)
    end = new Date(now); end.setHours(23, 59, 59, 999)
  } else if (type === 'yearly') {
    start = new Date(now.getFullYear(), 0, 1)
    end = new Date(now.getFullYear(), 11, 31, 23, 59, 59, 999)
  } else {
    start = from ? new Date(from) : new Date(0)
    end = to ? new Date(new Date(to).setHours(23, 59, 59, 999)) : new Date()
  }
  return { start, end }
}

function inr(n) {
  return 'Rs.' + (n || 0).toLocaleString('en-IN')
}

function buildMatch(start, end, coupon) {
  const match = { createdAt: { $gte: start, $lte: end } }
  if (coupon) match.couponCode = coupon.toUpperCase()
  return match
}

/**
 * Extract partial-return rows from an order.
 *
 * Handles ALL common return patterns:
 *   a) item.returnStatus === 'Completed' / 'completed'
 *   b) item.isReturned === true
 *   c) item.status === 'Returned' / 'returned'  ← newly added
 *
 * Also handles:
 *   - item.productId being a populated object OR a plain ObjectId string
 *   - item.salePrice / item.price / item.offerPrice for unit price
 *   - item.discountAmount / item.discount / computed from offerPrice
 *
 * Returns a flat array, one object per returned item.
 */
function extractReturnedItems(order) {
  if (!Array.isArray(order.items)) return []

  return order.items
    .filter(item =>
      item.returnStatus === 'Completed' ||
      item.returnStatus === 'completed' ||
      item.isReturned === true ||
      item.status === 'Returned' ||
      item.status === 'returned'
    )
    .map(item => {
      // ── Resolve product name ───────────────────────────────────────────────
      let productName = '—'
      if (item.productName) productName = item.productName
      else if (item.name) productName = item.name
      else if (item.productId && typeof item.productId === 'object' && item.productId.name)
        productName = item.productId.name

      // ── Resolve unit price ─────────────────────────────────────────────────
      // Try every common field name, prefer the actual paid price
      const unitPrice =
        item.price ??
        item.salePrice ??
        item.offerPrice ??
        item.mrp ??
        0

      const qty = item.quantity || 1
      const linePrice = unitPrice * qty   // total for this line before discount

      // ── Resolve item-level discount ────────────────────────────────────────
      const itemDiscount =
        item.discountAmount ??
        item.discount ??
        item.couponDiscount ??
        0

      // ── Net refund for this item ───────────────────────────────────────────
      const itemNet = linePrice - itemDiscount

      return {
        // order-level
        orderId: order.orderId,
        orderDate: order.createdAt,
        customer: order.userId?.name || order.shippingAddress?.name || '—',
        email: order.userId?.email || order.shippingAddress?.email || '—',
        paymentMethod: order.paymentMethod || '—',
        couponCode: order.couponCode || '—',
        orderReturnStatus: order.returnStatus,
        // item-level
        productName,
        variantInfo: [item.size, item.color, item.variant].filter(Boolean).join(' / ') || '—',
        quantity: qty,
        unitPrice,
        linePrice,        // unitPrice × qty
        itemDiscount,
        itemNet,
        returnReason: item.returnReason || order.returnReason || '—',
        returnedAt: item.returnedAt || order.updatedAt || order.createdAt,
      }
    })
}

// ─────────────────────────────────────────────────────────────────────────────
// loadSalesReport
// ─────────────────────────────────────────────────────────────────────────────
const loadSalesReport = async (req, res) => {
  try {
    const {
      type = 'daily',
      from = '',
      to = '',
      coupon = '',
      page = 1
    } = req.query

    const LIMIT = 10
    const currentPage = Math.max(1, parseInt(page))
    const { start, end } = getDateRange(type, from, to)
    const match = buildMatch(start, end, coupon)

    const [summary] = await Order.aggregate([
      { $match: match },
      {
        $group: {
          _id: null,
          totalOrders: { $sum: 1 },
          grossRevenue: { $sum: '$subtotal' },
          totalDiscount: { $sum: '$totalDiscount' },
          couponDiscount: { $sum: '$couponDiscount' },
          netRevenue: { $sum: '$finalAmount' },
          cancelledOrders: { $sum: { $cond: [{ $eq: ['$orderStatus', 'Cancelled'] }, 1, 0] } },
          returnedOrders: { $sum: { $cond: [{ $eq: ['$returnStatus', 'Completed'] }, 1, 0] } },
          cancelledAmount: { $sum: { $cond: [{ $eq: ['$orderStatus', 'Cancelled'] }, '$finalAmount', 0] } },
          returnedAmount: { $sum: { $cond: [{ $eq: ['$returnStatus', 'Completed'] }, '$finalAmount', 0] } },
        }
      }
    ])

    const stats = summary || {
      totalOrders: 0, grossRevenue: 0, totalDiscount: 0,
      couponDiscount: 0, netRevenue: 0,
      cancelledOrders: 0, returnedOrders: 0,
      cancelledAmount: 0, returnedAmount: 0,
    }

    const couponBreakdown = await Order.aggregate([
      { $match: { ...match, couponCode: { $ne: null } } },
      {
        $group: {
          _id: '$couponCode',
          uses: { $sum: 1 },
          totalDeducted: { $sum: '$couponDiscount' },
          netRevenue: { $sum: '$finalAmount' }
        }
      },
      { $sort: { totalDeducted: -1 } }
    ])

    let groupId
    if (type === 'daily') groupId = { hour: { $hour: '$createdAt' } }
    else if (type === 'weekly') groupId = { dayOfWeek: { $dayOfWeek: '$createdAt' } }
    else groupId = { month: { $month: '$createdAt' } }

    const chartRaw = await Order.aggregate([
      { $match: match },
      {
        $group: {
          _id: groupId,
          gross: { $sum: '$subtotal' },
          net: { $sum: '$finalAmount' },
          discount: { $sum: '$totalDiscount' },
        }
      },
      { $sort: { '_id.hour': 1, '_id.dayOfWeek': 1, '_id.month': 1 } }
    ])

    const DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']
    const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

    let chartLabels, chartGross, chartNet, chartDiscount
    if (type === 'daily') {
      chartLabels = chartRaw.map(d => `${d._id.hour}:00`)
      chartGross = chartRaw.map(d => d.gross)
      chartNet = chartRaw.map(d => d.net)
      chartDiscount = chartRaw.map(d => d.discount)
    } else if (type === 'weekly') {
      chartLabels = chartRaw.map(d => DAYS[(d._id.dayOfWeek - 1 + 7) % 7])
      chartGross = chartRaw.map(d => d.gross)
      chartNet = chartRaw.map(d => d.net)
      chartDiscount = chartRaw.map(d => d.discount)
    } else {
      chartLabels = chartRaw.map(d => MONTHS[d._id.month - 1])
      chartGross = chartRaw.map(d => d.gross)
      chartNet = chartRaw.map(d => d.net)
      chartDiscount = chartRaw.map(d => d.discount)
    }

    const total = await Order.countDocuments(match)

    // Paginated orders for the main table
    const orders = await Order.find(match)
      .populate('userId', 'name email')
      .sort({ createdAt: -1 })
      .skip((currentPage - 1) * LIMIT)
      .limit(LIMIT)
      .lean()

    // ALL orders (unpaginated) for cancelled / returned sections
    const allOrders = await Order.find(match)
      .populate('userId', 'name email')
      .sort({ createdAt: -1 })
      .lean()

    const cancelledOrders = allOrders.filter(o => o.orderStatus === 'Cancelled')

    // ── Build item-level returned rows (partial-return aware) ──────────────
    // This picks up individual items where only ONE item in a multi-item
    // order was returned, giving exact per-item pricing & reason.
    const returnedItemRows = allOrders.flatMap(o => extractReturnedItems(o))

    // Order-level fallback: orders where the WHOLE order is returned
    // (used only when zero item-level data is found)
    const returnedOrders = allOrders.filter(o => o.returnStatus === 'Completed')

    const allCoupons = await Order.distinct('couponCode', {
      createdAt: { $gte: new Date(Date.now() - 365 * 24 * 60 * 60 * 1000) },
      couponCode: { $ne: null }
    })

    res.render('admin/salesReport', {
      user: req.session.admin || null,
      stats,
      orders,
      couponBreakdown,
      allCoupons,
      cancelledOrders,
      returnedItemRows,   // ← item-level rows (partial-return aware) — PRIMARY
      returnedOrders,     // ← order-level fallback (whole-order returns)
      chartLabels: JSON.stringify(chartLabels),
      chartGross: JSON.stringify(chartGross),
      chartNet: JSON.stringify(chartNet),
      chartDiscount: JSON.stringify(chartDiscount),
      type, from, to, coupon,
      currentPage,
      totalPages: Math.ceil(total / LIMIT),
      total,
      start, end
    })
  } catch (err) {
    console.error('loadSalesReport error:', err)
    res.status(HttpStatus.INTERNAL_SERVER_ERROR).render('error', { message: 'Could not load sales report.' })
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// downloadPDF
// ─────────────────────────────────────────────────────────────────────────────
const downloadPDF = async (req, res) => {
  try {
    const { type = 'daily', from = '', to = '', coupon = '' } = req.query
    const { start, end } = getDateRange(type, from, to)
    const match = buildMatch(start, end, coupon)

    const [summary] = await Order.aggregate([
      { $match: match },
      {
        $group: {
          _id: null,
          totalOrders: { $sum: 1 },
          grossRevenue: { $sum: '$subtotal' },
          totalDiscount: { $sum: '$totalDiscount' },
          couponDiscount: { $sum: '$couponDiscount' },
          netRevenue: { $sum: '$finalAmount' },
          cancelledOrders: { $sum: { $cond: [{ $eq: ['$orderStatus', 'Cancelled'] }, 1, 0] } },
          returnedOrders: { $sum: { $cond: [{ $eq: ['$returnStatus', 'Completed'] }, 1, 0] } },
          cancelledAmount: { $sum: { $cond: [{ $eq: ['$orderStatus', 'Cancelled'] }, '$finalAmount', 0] } },
          returnedAmount: { $sum: { $cond: [{ $eq: ['$returnStatus', 'Completed'] }, '$finalAmount', 0] } },
        }
      }
    ])

    const stats = summary || {
      totalOrders: 0, grossRevenue: 0, totalDiscount: 0,
      couponDiscount: 0, netRevenue: 0,
      cancelledOrders: 0, returnedOrders: 0,
      cancelledAmount: 0, returnedAmount: 0,
    }

    const orders = await Order.find(match)
      .populate('userId', 'name email')
      .sort({ createdAt: -1 })
      .limit(500)
      .lean()

    // ── Item-level returned rows (partial-return aware) ────────────────────
    const returnedItemRows = orders.flatMap(o => extractReturnedItems(o))
    const cancelledOrders = orders.filter(o => o.orderStatus === 'Cancelled')
    // Fallback only when NO item-level data exists at all
    const returnedOrders = returnedItemRows.length === 0
      ? orders.filter(o => o.returnStatus === 'Completed')
      : []

    // ── PDF setup ──────────────────────────────────────────────────────────
    const doc = new PDFDocument({
      margin: 40,
      size: 'A4',
      layout: 'landscape',
      bufferPages: true,
      autoFirstPage: true,
    })

    res.setHeader('Content-Type', 'application/pdf')
    res.setHeader('Content-Disposition', `attachment; filename="blushberry-sales-${type}-${Date.now()}.pdf"`)
    doc.pipe(res)

    const MARGIN = 40
    const PAGE_W = doc.page.width
    const PAGE_H = doc.page.height
    const USABLE_W = PAGE_W - MARGIN * 2
    const LEFT = MARGIN
    const FOOT_H = 30
    const SAFE_BOTTOM = PAGE_H - FOOT_H - 4

    function newPage() {
      doc.addPage()
      return MARGIN
    }

    // ── Header banner ──────────────────────────────────────────────────────
    doc.rect(0, 0, PAGE_W, 54).fill('#c93060')
    doc.fontSize(20).font('Helvetica-Bold').fillColor('#ffffff')
      .text('Blush-Berry', LEFT, 14, { lineBreak: false })
    doc.fontSize(10).font('Helvetica').fillColor('rgba(255,255,255,0.85)')
      .text('Sales Report', LEFT, 36, { lineBreak: false })
    doc.fontSize(9).font('Helvetica').fillColor('rgba(255,255,255,0.75)')
      .text(`Generated: ${new Date().toLocaleString('en-IN')}`, LEFT, 20,
        { align: 'right', width: USABLE_W, lineBreak: false })

    doc.fontSize(9).font('Helvetica').fillColor('#555555')
      .text(`Period: ${start.toDateString()}  →  ${end.toDateString()}`, LEFT, 60, { lineBreak: false })

    // ── Summary cards ──────────────────────────────────────────────────────
    const CARD_GAP = 8
    const CARD_W = (USABLE_W - CARD_GAP * 2) / 3
    const CARD_H = 52

    const row1Cards = [
      { label: 'Total Orders', value: String(stats.totalOrders), color: '#3a1a2e' },
      { label: 'Gross Revenue', value: inr(stats.grossRevenue), color: '#3a1a2e' },
      { label: 'Net Revenue', value: inr(stats.netRevenue), color: '#c93060' },
    ]
    const row2Cards = [
      { label: 'Total Discounts', value: inr(stats.totalDiscount), color: '#e8527a' },
      { label: `Cancelled (${stats.cancelledOrders})`, value: inr(stats.cancelledAmount), color: '#a0222a' },
      {
        label: `Returned Orders (${stats.returnedOrders}) / Items (${returnedItemRows.length})`,
        value: inr(stats.returnedAmount), color: '#9a6200'
      },
    ]

    function drawCards(cards, startY) {
      cards.forEach((c, i) => {
        const x = LEFT + i * (CARD_W + CARD_GAP)
        doc.roundedRect(x, startY, CARD_W, CARD_H, 6).fill('#fdf0f4')
        doc.fontSize(8).font('Helvetica').fillColor('#b07090')
          .text(c.label, x + 8, startY + 8, { width: CARD_W - 16, lineBreak: false })
        doc.fontSize(12).font('Helvetica-Bold').fillColor(c.color)
          .text(c.value, x + 8, startY + 26, { width: CARD_W - 16, lineBreak: false })
      })
      return startY + CARD_H + 10
    }

    let mainY = 74
    mainY = drawCards(row1Cards, mainY)
    mainY = drawCards(row2Cards, mainY)
    mainY += 10

    // ── Column definitions ─────────────────────────────────────────────────
    const RAW_COLS = [
      { label: '#', raw: 20, align: 'center' },
      { label: 'Order ID', raw: 105, align: 'left' },
      { label: 'Date', raw: 60, align: 'left' },
      { label: 'Customer', raw: 95, align: 'left' },
      { label: 'Payment', raw: 55, align: 'left' },
      { label: 'Coupon', raw: 55, align: 'left' },
      { label: 'Gross', raw: 72, align: 'right' },
      { label: 'Discount', raw: 72, align: 'right' },
      { label: 'Net', raw: 72, align: 'right' },
      { label: 'Status', raw: 65, align: 'left' },
    ]
    const rawTotal = RAW_COLS.reduce((s, c) => s + c.raw, 0)
    const ORDER_COLS = RAW_COLS.map(c => ({
      ...c, width: Math.floor((c.raw / rawTotal) * USABLE_W)
    }))
    const orderColsTotal = ORDER_COLS.reduce((s, c) => s + c.width, 0)
    ORDER_COLS[ORDER_COLS.length - 1].width += USABLE_W - orderColsTotal

    // Returned-item columns — now includes Unit Price column
    const RET_RAW = [
      { label: '#', raw: 18, align: 'center' },
      { label: 'Order ID', raw: 85, align: 'left' },
      { label: 'Date', raw: 50, align: 'left' },
      { label: 'Customer', raw: 80, align: 'left' },
      { label: 'Product', raw: 105, align: 'left' },
      { label: 'Variant', raw: 55, align: 'left' },
      { label: 'Qty', raw: 24, align: 'center' },
      { label: 'Unit Price', raw: 58, align: 'right' },
      { label: 'Line Total', raw: 58, align: 'right' },
      { label: 'Discount', raw: 55, align: 'right' },
      { label: 'Net Refund', raw: 60, align: 'right' },
      { label: 'Status', raw: 50, align: 'left' },
      { label: 'Reason', raw: 103, align: 'left' },
    ]
    const retRawTotal = RET_RAW.reduce((s, c) => s + c.raw, 0)
    const RET_COLS = RET_RAW.map(c => ({
      ...c, width: Math.floor((c.raw / retRawTotal) * USABLE_W)
    }))
    const retColsTotal = RET_COLS.reduce((s, c) => s + c.width, 0)
    RET_COLS[RET_COLS.length - 1].width += USABLE_W - retColsTotal

    const ROW_H = 18
    const THEAD_H = 20

    const STATUS_COLORS = {
      delivered: '#1a7a46',
      placed: '#2249a0',
      processing: '#9a6200',
      shipped: '#6c27b0',
      cancelled: '#a0222a',
      'partial return': '#cc8800', 
      returned: '#9a6200',
      failed: '#a0222a',
    }

    function drawHeader(cols, y, fill = '#c93060') {
      doc.rect(LEFT, y, USABLE_W, THEAD_H).fill(fill)
      let x = LEFT
      cols.forEach(c => {
        doc.fontSize(7.5).font('Helvetica-Bold').fillColor('#ffffff')
          .text(c.label, x + 3, y + 6, {
            width: c.width - 6, align: c.align, lineBreak: false, ellipsis: true
          })
        x += c.width
      })
      return y + THEAD_H
    }

    function drawOrderRow(order, idx, y, headerFill = '#c93060') {
      if (y + ROW_H > SAFE_BOTTOM) {
        y = newPage()
        y = drawHeader(ORDER_COLS, y, headerFill)
      }
      const bg = idx % 2 === 0 ? '#ffffff' : '#fdf5f7'
      doc.rect(LEFT, y, USABLE_W, ROW_H).fill(bg)

      const returnedItemCount = Array.isArray(order.items)
        ? order.items.filter(item =>
          item.returnStatus === 'Completed' || item.returnStatus === 'completed' ||
          item.isReturned === true ||
          item.status === 'Returned' || item.status === 'returned'
        ).length
        : 0

      const status = order.returnStatus === 'Completed'
        ? 'Returned'
        : returnedItemCount > 0
          ? `Partial Return (${returnedItemCount}/${(order.items || []).length})`
          : (order.orderStatus || '—')

      const cells = [
        String(idx + 1),
        order.orderId || '—',
        new Date(order.createdAt).toLocaleDateString('en-IN'),
        order.userId?.name || order.shippingAddress?.name || '—',
        order.paymentMethod || '—',
        order.couponCode || '—',
        inr(order.subtotal),
        order.totalDiscount > 0 ? inr(order.totalDiscount) : '—',
        inr(order.finalAmount),
        status,
      ]

      let x = LEFT
      cells.forEach((cell, i) => {
        const c = ORDER_COLS[i]
        let colour = '#3a1a2e'
        if (i === 9) colour = STATUS_COLORS[status.toLowerCase()] || '#3a1a2e'
        if (i === 7 && cell !== '—') colour = '#e8527a'
        if (i === 8) colour = '#c93060'
        doc.fontSize(7).font(i === 9 ? 'Helvetica-Bold' : 'Helvetica').fillColor(colour)
          .text(cell, x + 3, y + 5, {
            width: c.width - 6, align: c.align, lineBreak: false, ellipsis: true
          })
        x += c.width
      })

      doc.moveTo(LEFT, y + ROW_H)
        .lineTo(LEFT + USABLE_W, y + ROW_H)
        .strokeColor('#f4d0da').lineWidth(0.3).stroke()

      return y + ROW_H
    }

    // ── Draw one returned-item row (now with unit price + line total) ──────
    function drawReturnedItemRow(row, idx, y, headerFill = '#9a6200') {
      if (y + ROW_H > SAFE_BOTTOM) {
        y = newPage()
        y = drawHeader(RET_COLS, y, headerFill)
      }
      const bg = idx % 2 === 0 ? '#ffffff' : '#fdf9f0'
      doc.rect(LEFT, y, USABLE_W, ROW_H).fill(bg)

      const cells = [
        String(idx + 1),
        row.orderId || '—',
        new Date(row.orderDate).toLocaleDateString('en-IN'),
        row.customer || '—',
        row.productName || '—',
        row.variantInfo || '—',
        String(row.quantity),
        inr(row.unitPrice),               // unit price (single item)
        inr(row.linePrice),               // qty × unit price
        row.itemDiscount > 0 ? inr(row.itemDiscount) : '—',
        inr(row.itemNet),                 // net refund for this item
        'Returned',
        row.returnReason || '—',
      ]

      let x = LEFT
      cells.forEach((cell, i) => {
        const c = RET_COLS[i]
        let colour = '#3a1a2e'
        if (i === 9 && cell !== '—') colour = '#e8527a'   // discount — pink
        if (i === 10) colour = '#9a6200'                    // net refund — amber
        if (i === 11) colour = '#9a6200'                    // status label
        doc.fontSize(7)
          .font(i === 11 ? 'Helvetica-Bold' : 'Helvetica')
          .fillColor(colour)
          .text(cell, x + 3, y + 5, {
            width: c.width - 6, align: c.align, lineBreak: false, ellipsis: true
          })
        x += c.width
      })

      doc.moveTo(LEFT, y + ROW_H)
        .lineTo(LEFT + USABLE_W, y + ROW_H)
        .strokeColor('#f4d0da').lineWidth(0.3).stroke()

      return y + ROW_H
    }

    function drawSectionHeading(title, y) {
      const HEADING_H = 16
      if (y + HEADING_H + THEAD_H + ROW_H > SAFE_BOTTOM) {
        y = newPage()
      }
      doc.fontSize(11).font('Helvetica-Bold').fillColor('#3a1a2e')
        .text(title, LEFT, y, { lineBreak: false })
      return y + HEADING_H
    }

    // ── Section 1: Order Details ───────────────────────────────────────────
    if (mainY + 16 + THEAD_H + ROW_H > SAFE_BOTTOM) mainY = newPage()
    doc.fontSize(11).font('Helvetica-Bold').fillColor('#3a1a2e')
      .text('Order Details', LEFT, mainY, { lineBreak: false })
    mainY += 16
    mainY = drawHeader(ORDER_COLS, mainY, '#c93060')
    orders.forEach((o, i) => { mainY = drawOrderRow(o, i, mainY, '#c93060') })
    mainY += 16

    // ── Section 2: Cancelled Orders ────────────────────────────────────────
    if (cancelledOrders.length > 0) {
      mainY = drawSectionHeading(`Cancelled Orders (${cancelledOrders.length})`, mainY)
      mainY = drawHeader(ORDER_COLS, mainY, '#a0222a')
      cancelledOrders.forEach((o, i) => { mainY = drawOrderRow(o, i, mainY, '#a0222a') })
      mainY += 16
    }

    // ── Section 3: Returned Items (item-level, partial-return aware) ───────
    if (returnedItemRows.length > 0) {
      mainY = drawSectionHeading(
        `Returned Items — Per Item Detail (${returnedItemRows.length} item(s) across ${[...new Set(returnedItemRows.map(r => r.orderId))].length} order(s))`,
        mainY
      )
      mainY = drawHeader(RET_COLS, mainY, '#9a6200')
      returnedItemRows.forEach((row, i) => { mainY = drawReturnedItemRow(row, i, mainY, '#9a6200') })
      mainY += 16
    } else if (returnedOrders.length > 0) {
      // Fallback: whole-order returns when no item-level data
      mainY = drawSectionHeading(`Returned Orders (${returnedOrders.length})`, mainY)
      mainY = drawHeader(ORDER_COLS, mainY, '#9a6200')
      returnedOrders.forEach((o, i) => { mainY = drawOrderRow(o, i, mainY, '#9a6200') })
      mainY += 16
    }

    // ── Page footers ───────────────────────────────────────────────────────
    const range = doc.bufferedPageRange()
    for (let i = range.start; i < range.start + range.count; i++) {
      doc.switchToPage(i)
      doc.fontSize(7.5).font('Helvetica').fillColor('#aaaaaa')
        .text(
          `Blush-Berry — Confidential Sales Report   |   Page ${i - range.start + 1} of ${range.count}`,
          LEFT, PAGE_H - 24,
          { width: USABLE_W, align: 'center', lineBreak: false }
        )
    }

    doc.end()

  } catch (err) {
    console.error('downloadPDF error:', err)
    res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({ success: false, message: 'Could not generate PDF.' })
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// downloadExcel
// ─────────────────────────────────────────────────────────────────────────────
const downloadExcel = async (req, res) => {
  try {
    const { type = 'daily', from = '', to = '', coupon = '' } = req.query
    const { start, end } = getDateRange(type, from, to)
    const match = buildMatch(start, end, coupon)

    const [summary] = await Order.aggregate([
      { $match: match },
      {
        $group: {
          _id: null,
          totalOrders: { $sum: 1 },
          grossRevenue: { $sum: '$subtotal' },
          totalDiscount: { $sum: '$totalDiscount' },
          couponDiscount: { $sum: '$couponDiscount' },
          netRevenue: { $sum: '$finalAmount' },
          cancelledOrders: { $sum: { $cond: [{ $eq: ['$orderStatus', 'Cancelled'] }, 1, 0] } },
          returnedOrders: { $sum: { $cond: [{ $eq: ['$returnStatus', 'Completed'] }, 1, 0] } },
          cancelledAmount: { $sum: { $cond: [{ $eq: ['$orderStatus', 'Cancelled'] }, '$finalAmount', 0] } },
          returnedAmount: { $sum: { $cond: [{ $eq: ['$returnStatus', 'Completed'] }, '$finalAmount', 0] } },
        }
      }
    ])

    const stats = summary || {
      totalOrders: 0, grossRevenue: 0, totalDiscount: 0,
      couponDiscount: 0, netRevenue: 0,
      cancelledOrders: 0, returnedOrders: 0,
      cancelledAmount: 0, returnedAmount: 0,
    }

    const orders = await Order.find(match)
      .populate('userId', 'name email')
      .sort({ createdAt: -1 })
      .limit(5000)
      .lean()

    // ── Item-level returned rows (partial-return aware) ────────────────────
    const returnedItemRows = orders.flatMap(o => extractReturnedItems(o))
    const cancelledOrders = orders.filter(o => o.orderStatus === 'Cancelled')
    const returnedOrders = orders.filter(o => o.returnStatus === 'Completed')

    const wb = new ExcelJS.Workbook()
    wb.creator = 'Blush-Berry'
    wb.created = new Date()

    // ── Style constants ────────────────────────────────────────────────────
    const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC93060' } }
    const cancelFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFA0222A' } }
    const returnFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9A6200' } }
    const altFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF0F4' } }
    const altRetFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9F0' } }
    const headerFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' }
    const bodyFont = { size: 10, name: 'Calibri' }
    const boldFont = { bold: true, size: 10, name: 'Calibri' }
    const pinkFont = { bold: true, color: { argb: 'FFC93060' }, size: 10, name: 'Calibri' }
    const redFont = { color: { argb: 'FFE8527A' }, size: 10, name: 'Calibri' }
    const cancelFont = { bold: true, color: { argb: 'FFA0222A' }, size: 10, name: 'Calibri' }
    const returnFont = { bold: true, color: { argb: 'FF9A6200' }, size: 10, name: 'Calibri' }
    const thinBorder = {
      top: { style: 'thin', color: { argb: 'FFF4D0DA' } },
      bottom: { style: 'thin', color: { argb: 'FFF4D0DA' } },
      left: { style: 'thin', color: { argb: 'FFF4D0DA' } },
      right: { style: 'thin', color: { argb: 'FFF4D0DA' } },
    }

    function styleHeaderRow(row, fill = headerFill) {
      row.eachCell(cell => {
        cell.fill = fill
        cell.font = headerFont
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
        cell.border = thinBorder
      })
      row.height = 24
    }

    function styleDataRow(row, isAlt, amountCols = [], altF = altFill) {
      row.eachCell({ includeEmpty: true }, (cell, colNum) => {
        if (isAlt) cell.fill = altF
        cell.font = amountCols.includes(colNum) ? boldFont : bodyFont
        cell.alignment = { vertical: 'middle', wrapText: false }
        cell.border = thinBorder
      })
      row.height = 18
    }

    // ── Sheet 1: Summary ──────────────────────────────────────────────────
    const ws1 = wb.addWorksheet('Summary', { properties: { tabColor: { argb: 'FFC93060' } } })
    ws1.mergeCells('A1:C1')
    const titleCell = ws1.getCell('A1')
    titleCell.value = 'Blush-Berry — Sales Report'
    titleCell.font = { bold: true, size: 16, color: { argb: 'FFC93060' }, name: 'Calibri' }
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
    ws1.getRow(1).height = 32

    ws1.mergeCells('A2:C2')
    const periodCell = ws1.getCell('A2')
    periodCell.value = `Period: ${start.toDateString()}  →  ${end.toDateString()}`
    periodCell.font = { italic: true, size: 10, color: { argb: 'FF7A4A5E' }, name: 'Calibri' }
    periodCell.alignment = { horizontal: 'center', vertical: 'middle' }
    ws1.getRow(2).height = 18

    ws1.addRow([])
    ws1.columns = [
      { key: 'metric', width: 40 },
      { key: 'value', width: 22 },
      { key: 'note', width: 42 },
    ]
    styleHeaderRow(ws1.addRow(['Metric', 'Value', 'Note']))

    const summaryData = [
      ['Total Orders', stats.totalOrders,
        `${start.toDateString()} to ${end.toDateString()}`],
      ['Gross Revenue (Rs.)', stats.grossRevenue,
        'Before any discounts'],
      ['Total Discounts (Rs.)', stats.totalDiscount,
        'Offer + coupon combined'],
      ['Coupon Discounts (Rs.)', stats.couponDiscount,
        'Coupon deductions only'],
      ['Item Offer Discounts (Rs.)', Math.max(0, stats.totalDiscount - stats.couponDiscount),
        'Product/category offers'],
      ['Net Revenue (Rs.)', stats.netRevenue,
        'After all deductions'],
      ['Cancelled Orders (count)', stats.cancelledOrders,
        'Fully cancelled orders'],
      ['Cancelled Orders Value (Rs.)', stats.cancelledAmount,
        'Revenue lost to cancellations'],
      ['Returned Orders (count)', stats.returnedOrders,
        'Fully returned orders'],
      ['Returned Orders Value (Rs.)', stats.returnedAmount,
        'Revenue lost to returns'],
      ['Returned Items — item count', returnedItemRows.length,
        'Individual items returned (partial returns included)'],
      ['Returned Items — unique orders', [...new Set(returnedItemRows.map(r => r.orderId))].length,
        'Orders that had at least one item returned'],
      ['Returned Items — total net refund (Rs.)', returnedItemRows.reduce((s, r) => s + r.itemNet, 0),
        'Sum of per-item net refund amounts'],
    ]

    summaryData.forEach((rowData, i) => {
      const r = ws1.addRow(rowData)
      styleDataRow(r, i % 2 === 1, [2])
      r.getCell(1).font = bodyFont
      r.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' }
      const label = rowData[0]
      if (label.includes('Net Revenue')) r.getCell(2).font = pinkFont
      else if (label.includes('Cancelled')) r.getCell(2).font = cancelFont
      else if (label.includes('Returned')) r.getCell(2).font = returnFont
      else if (label.includes('Discount')) r.getCell(2).font = redFont
      r.getCell(3).font = { italic: true, size: 9, color: { argb: 'FFB07090' }, name: 'Calibri' }
    })

    // ── Sheet 2: All Orders ───────────────────────────────────────────────
    const ws2 = wb.addWorksheet('All Orders', { properties: { tabColor: { argb: 'FFC93060' } } })
    ws2.columns = [
      { header: 'S.No', key: 'sno', width: 7 },
      { header: 'Order ID', key: 'orderId', width: 24 },
      { header: 'Date', key: 'date', width: 16 },
      { header: 'Customer Name', key: 'customer', width: 24 },
      { header: 'Email', key: 'email', width: 30 },
      { header: 'Payment', key: 'payment', width: 14 },
      { header: 'Coupon Code', key: 'coupon', width: 14 },
      { header: 'Gross (Rs.)', key: 'gross', width: 16 },
      { header: 'Discount (Rs.)', key: 'discount', width: 16 },
      { header: 'Net (Rs.)', key: 'net', width: 16 },
      { header: 'Status', key: 'status', width: 18 },
    ]
    styleHeaderRow(ws2.getRow(1))
    orders.forEach((o, idx) => {
      // Show "Partial Return" if some items returned but order not fully returned
      const retItems = extractReturnedItems(o)
const hasPartialReturn = retItems.length > 0 && o.returnStatus !== 'Completed'
const status = o.returnStatus === 'Completed'
  ? 'Returned'
  : hasPartialReturn
    ? `Partial Return (${retItems.length}/${(o.items || []).length})`
    : (o.orderStatus || '—')

      const r = ws2.addRow({
        sno: idx + 1,
        orderId: o.orderId || '—',
        date: new Date(o.createdAt).toLocaleDateString('en-IN'),
        customer: o.userId?.name || o.shippingAddress?.name || '—',
        email: o.userId?.email || o.shippingAddress?.email || '—',
        payment: o.paymentMethod || '—',
        coupon: o.couponCode || '—',
        gross: o.subtotal || 0,
        discount: o.totalDiscount || 0,
        net: o.finalAmount || 0,
        status,
      })
      styleDataRow(r, idx % 2 === 1, [8, 9, 10])
        ;[8, 9, 10].forEach(col => {
          r.getCell(col).alignment = { horizontal: 'right', vertical: 'middle' }
          r.getCell(col).numFmt = '#,##0.00'
        })
      if ((o.totalDiscount || 0) > 0) r.getCell(9).font = redFont
      r.getCell(10).font = pinkFont
      const statusCell = r.getCell(11)
      const sl = status.toLowerCase()
      if (sl === 'delivered') statusCell.font = { color: { argb: 'FF1A7A46' }, bold: true, size: 10, name: 'Calibri' }
      else if (sl === 'cancelled') statusCell.font = cancelFont
      else if (sl === 'returned') statusCell.font = returnFont
      else if (sl === 'partial return') statusCell.font = { color: { argb: 'FFCC8800' }, bold: true, size: 10, name: 'Calibri' }
      else if (sl === 'shipped') statusCell.font = { color: { argb: 'FF6C27B0' }, bold: true, size: 10, name: 'Calibri' }
      else if (sl === 'processing') statusCell.font = { color: { argb: 'FF9A6200' }, bold: true, size: 10, name: 'Calibri' }
      else statusCell.font = bodyFont
    })
    ws2.views = [{ state: 'frozen', ySplit: 1 }]
    ws2.autoFilter = { from: 'A1', to: 'K1' }

    // ── Sheet 3: Cancelled Orders ─────────────────────────────────────────
    const ws3 = wb.addWorksheet('Cancelled Orders', { properties: { tabColor: { argb: 'FFA0222A' } } })
    ws3.columns = [
      { header: 'S.No', key: 'sno', width: 7 },
      { header: 'Order ID', key: 'orderId', width: 24 },
      { header: 'Date', key: 'date', width: 16 },
      { header: 'Customer Name', key: 'customer', width: 24 },
      { header: 'Email', key: 'email', width: 30 },
      { header: 'Payment', key: 'payment', width: 14 },
      { header: 'Gross (Rs.)', key: 'gross', width: 16 },
      { header: 'Discount (Rs.)', key: 'discount', width: 16 },
      { header: 'Net (Rs.)', key: 'net', width: 16 },
      { header: 'Cancel Reason', key: 'reason', width: 30 },
    ]
    styleHeaderRow(ws3.getRow(1), cancelFill)
    if (cancelledOrders.length === 0) {
      const r = ws3.addRow(['No cancelled orders in this period.', '', '', '', '', '', '', '', '', ''])
      ws3.mergeCells('A2:J2')
      r.getCell(1).font = { italic: true, color: { argb: 'FFB07090' }, name: 'Calibri' }
      r.getCell(1).alignment = { horizontal: 'center' }
    } else {
      cancelledOrders.forEach((o, idx) => {
        const r = ws3.addRow({
          sno: idx + 1,
          orderId: o.orderId || '—',
          date: new Date(o.createdAt).toLocaleDateString('en-IN'),
          customer: o.userId?.name || o.shippingAddress?.name || '—',
          email: o.userId?.email || o.shippingAddress?.email || '—',
          payment: o.paymentMethod || '—',
          gross: o.subtotal || 0,
          discount: o.totalDiscount || 0,
          net: o.finalAmount || 0,
          reason: o.cancelReason || '—',
        })
        styleDataRow(r, idx % 2 === 1, [7, 8, 9])
          ;[7, 8, 9].forEach(col => {
            r.getCell(col).alignment = { horizontal: 'right', vertical: 'middle' }
            r.getCell(col).numFmt = '#,##0.00'
          })
        r.getCell(9).font = cancelFont
      })
    }
    ws3.views = [{ state: 'frozen', ySplit: 1 }]
    ws3.autoFilter = { from: 'A1', to: 'J1' }

    // ── Sheet 4: Returned Items (item-level, partial-return aware) ─────────
    // This sheet shows ONE ROW PER RETURNED ITEM.
    // e.g. if order has 2 items and only 1 was returned → 1 row with that
    // item's name, unit price, qty, discount, and net refund.
    const ws4 = wb.addWorksheet('Returned Items', { properties: { tabColor: { argb: 'FF9A6200' } } })
    ws4.columns = [
      { header: 'S.No', key: 'sno', width: 7 },
      { header: 'Order ID', key: 'orderId', width: 24 },
      { header: 'Order Date', key: 'date', width: 16 },
      { header: 'Customer Name', key: 'customer', width: 24 },
      { header: 'Email', key: 'email', width: 30 },
      { header: 'Payment', key: 'payment', width: 14 },
      { header: 'Product', key: 'product', width: 30 },
      { header: 'Variant', key: 'variant', width: 18 },
      { header: 'Qty Returned', key: 'qty', width: 12 },
      { header: 'Unit Price (Rs.)', key: 'unitPrice', width: 18 },
      { header: 'Line Total (Rs.)', key: 'linePrice', width: 18 },
      { header: 'Discount (Rs.)', key: 'discount', width: 16 },
      { header: 'Net Refund (Rs.)', key: 'net', width: 18 },
      { header: 'Item Status', key: 'itemStatus', width: 16 },
      { header: 'Return Reason', key: 'reason', width: 32 },
      { header: 'Returned At', key: 'returnedAt', width: 18 },
    ]
    styleHeaderRow(ws4.getRow(1), returnFill)

    if (returnedItemRows.length === 0) {
      // Fallback to order-level if genuinely no item-level data
      if (returnedOrders.length === 0) {
        const r = ws4.addRow(['No returned items in this period.', ...Array(15).fill('')])
        ws4.mergeCells('A2:P2')
        r.getCell(1).font = { italic: true, color: { argb: 'FFB07090' }, name: 'Calibri' }
        r.getCell(1).alignment = { horizontal: 'center' }
      } else {
        // Whole-order returns, no item data stored
        returnedOrders.forEach((o, idx) => {
          const r = ws4.addRow({
            sno: idx + 1,
            orderId: o.orderId || '—',
            date: new Date(o.createdAt).toLocaleDateString('en-IN'),
            customer: o.userId?.name || o.shippingAddress?.name || '—',
            email: o.userId?.email || o.shippingAddress?.email || '—',
            payment: o.paymentMethod || '—',
            product: '(Full order — no item detail)',
            variant: '—',
            qty: '—',
            unitPrice: '—',
            linePrice: o.subtotal || 0,
            discount: o.totalDiscount || 0,
            net: o.finalAmount || 0,
            itemStatus: 'Returned',
            reason: o.returnReason || '—',
            returnedAt: new Date(o.updatedAt || o.createdAt).toLocaleDateString('en-IN'),
          })
          styleDataRow(r, idx % 2 === 1, [11, 12, 13], altRetFill)
            ;[11, 12, 13].forEach(col => {
              r.getCell(col).alignment = { horizontal: 'right', vertical: 'middle' }
              r.getCell(col).numFmt = '#,##0.00'
            })
          r.getCell(13).font = returnFont
          r.getCell(14).font = returnFont
        })
      }
    } else {
      // ── PRIMARY PATH: item-level rows ────────────────────────────────────
      returnedItemRows.forEach((row, idx) => {
        const r = ws4.addRow({
          sno: idx + 1,
          orderId: row.orderId || '—',
          date: new Date(row.orderDate).toLocaleDateString('en-IN'),
          customer: row.customer || '—',
          email: row.email || '—',
          payment: row.paymentMethod || '—',
          product: row.productName || '—',
          variant: row.variantInfo || '—',
          qty: row.quantity,
          unitPrice: row.unitPrice,          // ← price for a single unit
          linePrice: row.linePrice,          // ← unitPrice × qty
          discount: row.itemDiscount || 0,
          net: row.itemNet,            // ← net refund for this item
          itemStatus: 'Returned',
          reason: row.returnReason || '—',
          returnedAt: new Date(row.returnedAt).toLocaleDateString('en-IN'),
        })
        styleDataRow(r, idx % 2 === 1, [10, 11, 12, 13], altRetFill)
          ;[10, 11, 12, 13].forEach(col => {
            r.getCell(col).alignment = { horizontal: 'right', vertical: 'middle' }
            r.getCell(col).numFmt = '#,##0.00'
          })
        if ((row.itemDiscount || 0) > 0) r.getCell(12).font = redFont
        r.getCell(13).font = returnFont   // net refund — amber
        r.getCell(14).font = returnFont   // item status — amber
      })
    }
    ws4.views = [{ state: 'frozen', ySplit: 1 }]
    ws4.autoFilter = { from: 'A1', to: 'P1' }

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    res.setHeader('Content-Disposition', `attachment; filename="blushberry-sales-${type}-${Date.now()}.xlsx"`)
    await wb.xlsx.write(res)
    res.end()

  } catch (err) {
    console.error('downloadExcel error:', err)
    res.status(HttpStatus.INTERNAL_SERVER_ERROR).json({ success: false, message: 'Could not generate Excel.' })
  }
}

module.exports = { loadSalesReport, downloadPDF, downloadExcel }