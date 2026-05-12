const Order       = require('../../models/user/orderModel')
const ExcelJS     = require('exceljs')
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
    end   = new Date(now); end.setHours(23, 59, 59, 999)
  } else if (type === 'weekly') {
    const day = now.getDay()
    start = new Date(now); start.setDate(now.getDate() - day); start.setHours(0, 0, 0, 0)
    end   = new Date(now); end.setHours(23, 59, 59, 999)
  } else if (type === 'yearly') {
    start = new Date(now.getFullYear(), 0, 1)
    end   = new Date(now.getFullYear(), 11, 31, 23, 59, 59, 999)
  } else {
    start = from ? new Date(from) : new Date(0)
    end   = to   ? new Date(new Date(to).setHours(23, 59, 59, 999)) : new Date()
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

// ─────────────────────────────────────────────────────────────────────────────
// loadSalesReport
// ─────────────────────────────────────────────────────────────────────────────
const loadSalesReport = async (req, res) => {
  try {
    const {
      type   = 'daily',
      from   = '',
      to     = '',
      coupon = '',
      page   = 1
    } = req.query

    const LIMIT = 10
    const currentPage = Math.max(1, parseInt(page))
    const { start, end } = getDateRange(type, from, to)
    const match = buildMatch(start, end, coupon)

    const [summary] = await Order.aggregate([
      { $match: match },
      {
        $group: {
          _id:              null,
          totalOrders:      { $sum: 1 },
          grossRevenue:     { $sum: '$subtotal' },
          totalDiscount:    { $sum: '$totalDiscount' },
          couponDiscount:   { $sum: '$couponDiscount' },
          netRevenue:       { $sum: '$finalAmount' },
          cancelledOrders:  { $sum: { $cond: [{ $eq: ['$orderStatus',  'Cancelled'] }, 1, 0] } },
          returnedOrders:   { $sum: { $cond: [{ $eq: ['$returnStatus', 'Completed'] }, 1, 0] } },
          cancelledAmount:  { $sum: { $cond: [{ $eq: ['$orderStatus',  'Cancelled'] }, '$finalAmount', 0] } },
          returnedAmount:   { $sum: { $cond: [{ $eq: ['$returnStatus', 'Completed'] }, '$finalAmount', 0] } },
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
          _id:           '$couponCode',
          uses:          { $sum: 1 },
          totalDeducted: { $sum: '$couponDiscount' },
          netRevenue:    { $sum: '$finalAmount' }
        }
      },
      { $sort: { totalDeducted: -1 } }
    ])

    let groupId
    if (type === 'daily')        groupId = { hour:      { $hour:      '$createdAt' } }
    else if (type === 'weekly')  groupId = { dayOfWeek: { $dayOfWeek: '$createdAt' } }
    else                         groupId = { month:     { $month:     '$createdAt' } }

    const chartRaw = await Order.aggregate([
      { $match: match },
      {
        $group: {
          _id:      groupId,
          gross:    { $sum: '$subtotal'     },
          net:      { $sum: '$finalAmount'  },
          discount: { $sum: '$totalDiscount'},
        }
      },
      { $sort: { '_id.hour': 1, '_id.dayOfWeek': 1, '_id.month': 1 } }
    ])

    const DAYS   = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat']
    const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

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

    const total  = await Order.countDocuments(match)
    const orders = await Order.find(match)
      .populate('userId', 'name email')
      .sort({ createdAt: -1 })
      .skip((currentPage - 1) * LIMIT)
      .limit(LIMIT)
      .lean()

    const allCoupons = await Order.distinct('couponCode', {
      createdAt:  { $gte: new Date(Date.now() - 365 * 24 * 60 * 60 * 1000) },
      couponCode: { $ne: null }
    })

    res.render('admin/salesReport', {
      user: req.session.admin || null,
      stats,
      orders,
      couponBreakdown,
      allCoupons,
      chartLabels:   JSON.stringify(chartLabels),
      chartGross:    JSON.stringify(chartGross),
      chartNet:      JSON.stringify(chartNet),
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
// downloadPDF  — FIXED: no blank pages, correct Y tracking, proper overflow
// ─────────────────────────────────────────────────────────────────────────────
const downloadPDF = async (req, res) => {
  try {
    const { type = 'daily', from = '', to = '', coupon = '' } = req.query
    const { start, end } = getDateRange(type, from, to)
    const match = buildMatch(start, end, coupon)

    // ── Data ──────────────────────────────────────────────────────────────────
    const [summary] = await Order.aggregate([
      { $match: match },
      {
        $group: {
          _id:             null,
          totalOrders:     { $sum: 1 },
          grossRevenue:    { $sum: '$subtotal'      },
          totalDiscount:   { $sum: '$totalDiscount' },
          couponDiscount:  { $sum: '$couponDiscount'},
          netRevenue:      { $sum: '$finalAmount'   },
          cancelledOrders: { $sum: { $cond: [{ $eq: ['$orderStatus',  'Cancelled'] }, 1, 0] } },
          returnedOrders:  { $sum: { $cond: [{ $eq: ['$returnStatus', 'Completed'] }, 1, 0] } },
          cancelledAmount: { $sum: { $cond: [{ $eq: ['$orderStatus',  'Cancelled'] }, '$finalAmount', 0] } },
          returnedAmount:  { $sum: { $cond: [{ $eq: ['$returnStatus', 'Completed'] }, '$finalAmount', 0] } },
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

    // ── PDF setup ─────────────────────────────────────────────────────────────
    // KEY FIX: bufferPages:true is required for footer page numbers,
    // but we must track Y manually and NEVER let pdfkit auto-flow text
    // past the safe area — otherwise it silently creates blank continuation pages.
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

    const MARGIN   = 40
    const PAGE_W   = doc.page.width    // landscape A4 = 841.89
    const PAGE_H   = doc.page.height   // landscape A4 = 595.28
    const USABLE_W = PAGE_W - MARGIN * 2
    const LEFT     = MARGIN
    const FOOT_H   = 30                // reserved at bottom for footer
    // Safe bottom boundary — stop rendering rows before this line
    const SAFE_BOTTOM = PAGE_H - FOOT_H - 4

    // ── Reusable: add a new page and reset Y to top margin ───────────────────
    function newPage() {
      doc.addPage()
      return MARGIN
    }

    // ── Header banner (first page only) ──────────────────────────────────────
    doc.rect(0, 0, PAGE_W, 54).fill('#c93060')
    doc.fontSize(20).font('Helvetica-Bold').fillColor('#ffffff')
       .text('Blush-Berry', LEFT, 14, { lineBreak: false })
    doc.fontSize(10).font('Helvetica').fillColor('rgba(255,255,255,0.85)')
       .text('Sales Report', LEFT, 36, { lineBreak: false })
    doc.fontSize(9).font('Helvetica').fillColor('rgba(255,255,255,0.75)')
       .text(`Generated: ${new Date().toLocaleString('en-IN')}`, LEFT, 20,
             { align: 'right', width: USABLE_W, lineBreak: false })

    // Period line sits just below banner
    doc.fontSize(9).font('Helvetica').fillColor('#555555')
       .text(`Period: ${start.toDateString()}  →  ${end.toDateString()}`, LEFT, 60, { lineBreak: false })

    // ── Summary cards ─────────────────────────────────────────────────────────
    const CARD_GAP = 8
    const CARD_W   = (USABLE_W - CARD_GAP * 2) / 3
    const CARD_H   = 52

    const row1Cards = [
      { label: 'Total Orders',    value: String(stats.totalOrders),  color: '#3a1a2e' },
      { label: 'Gross Revenue',   value: inr(stats.grossRevenue),    color: '#3a1a2e' },
      { label: 'Net Revenue',     value: inr(stats.netRevenue),      color: '#c93060' },
    ]
    const row2Cards = [
      { label: 'Total Discounts',                              value: inr(stats.totalDiscount),   color: '#e8527a' },
      { label: `Cancelled (${stats.cancelledOrders})`,        value: inr(stats.cancelledAmount), color: '#a0222a' },
      { label: `Returned (${stats.returnedOrders})`,          value: inr(stats.returnedAmount),  color: '#9a6200' },
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

    // Cards start at y=74 (below banner+period line)
    let mainY = 74
    mainY = drawCards(row1Cards, mainY)
    mainY = drawCards(row2Cards, mainY)
    mainY += 10  // breathing room before table

    // ── Column definitions ────────────────────────────────────────────────────
    const RAW_COLS = [
      { label: '#',        raw: 20,  align: 'center' },
      { label: 'Order ID', raw: 105, align: 'left'   },
      { label: 'Date',     raw: 60,  align: 'left'   },
      { label: 'Customer', raw: 95,  align: 'left'   },
      { label: 'Payment',  raw: 55,  align: 'left'   },
      { label: 'Coupon',   raw: 55,  align: 'left'   },
      { label: 'Gross',    raw: 72,  align: 'right'  },
      { label: 'Discount', raw: 72,  align: 'right'  },
      { label: 'Net',      raw: 72,  align: 'right'  },
      { label: 'Status',   raw: 65,  align: 'left'   },
    ]
    const rawTotal   = RAW_COLS.reduce((s, c) => s + c.raw, 0)
    const ORDER_COLS = RAW_COLS.map(c => ({
      ...c,
      width: Math.floor((c.raw / rawTotal) * USABLE_W)
    }))
    // Distribute rounding remainder to last column
    const orderColsTotal = ORDER_COLS.reduce((s, c) => s + c.width, 0)
    ORDER_COLS[ORDER_COLS.length - 1].width += USABLE_W - orderColsTotal

    const ROW_H   = 18
    const THEAD_H = 20

    const STATUS_COLORS = {
      delivered:  '#1a7a46',
      placed:     '#2249a0',
      processing: '#9a6200',
      shipped:    '#6c27b0',
      cancelled:  '#a0222a',
      returned:   '#555555',
      failed:     '#a0222a',
    }

    // ── Draw a table header row, returns Y after header ───────────────────────
    function drawOrderHeader(y, fill = '#c93060') {
      doc.rect(LEFT, y, USABLE_W, THEAD_H).fill(fill)
      let x = LEFT
      ORDER_COLS.forEach(c => {
        doc.fontSize(7.5).font('Helvetica-Bold').fillColor('#ffffff')
           .text(c.label, x + 3, y + 6, {
             width: c.width - 6, align: c.align, lineBreak: false, ellipsis: true
           })
        x += c.width
      })
      return y + THEAD_H
    }

    // ── Draw one data row — handles page break BEFORE drawing ─────────────────
    // FIX: check overflow BEFORE rendering, not after. Return updated Y.
    function drawOrderRow(order, idx, y, headerFill = '#c93060') {
      // Check if this row + a little buffer will overflow the safe area
      if (y + ROW_H > SAFE_BOTTOM) {
        y = newPage()
        y = drawOrderHeader(y, headerFill)   // repeat header on new page
      }

      const bg = idx % 2 === 0 ? '#ffffff' : '#fdf5f7'
      doc.rect(LEFT, y, USABLE_W, ROW_H).fill(bg)

      const status = order.returnStatus === 'Completed' ? 'Returned' : (order.orderStatus || '—')
      const cells  = [
        String(idx + 1),
        order.orderId || '—',
        new Date(order.createdAt).toLocaleDateString('en-IN'),
        order.userId?.name || order.shippingAddress?.name || '—',
        order.paymentMethod || '—',
        order.couponCode    || '—',
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

      // Subtle row separator line
      doc.moveTo(LEFT, y + ROW_H)
         .lineTo(LEFT + USABLE_W, y + ROW_H)
         .strokeColor('#f4d0da').lineWidth(0.3).stroke()

      return y + ROW_H
    }

    // ── Section heading helper ────────────────────────────────────────────────
    // FIX: check if section heading + at least one row fits; if not, new page first
    function drawSectionHeading(title, y, fill = '#c93060') {
      const HEADING_H = 16
      // Need space for: section title + header row + at least 1 data row
      if (y + HEADING_H + THEAD_H + ROW_H > SAFE_BOTTOM) {
        y = newPage()
      }
      doc.fontSize(11).font('Helvetica-Bold').fillColor('#3a1a2e')
         .text(title, LEFT, y, { lineBreak: false })
      y += HEADING_H
      y = drawOrderHeader(y, fill)
      return y
    }

    // ── Order Details section ─────────────────────────────────────────────────
    // Check if heading + header row fit on current page
    if (mainY + 16 + THEAD_H + ROW_H > SAFE_BOTTOM) {
      mainY = newPage()
    }
    doc.fontSize(11).font('Helvetica-Bold').fillColor('#3a1a2e')
       .text('Order Details', LEFT, mainY, { lineBreak: false })
    mainY += 16
    mainY = drawOrderHeader(mainY)

    orders.forEach((o, i) => {
      mainY = drawOrderRow(o, i, mainY, '#c93060')
    })
    mainY += 16

    // ── Cancelled Orders section ──────────────────────────────────────────────
    const cancelledOrders = orders.filter(o => o.orderStatus === 'Cancelled')
    if (cancelledOrders.length > 0) {
      mainY = drawSectionHeading('Cancelled Orders', mainY, '#a0222a')
      cancelledOrders.forEach((o, i) => {
        mainY = drawOrderRow(o, i, mainY, '#a0222a')
      })
      mainY += 16
    }

    // ── Returned Orders section ───────────────────────────────────────────────
    const returnedOrders = orders.filter(o => o.returnStatus === 'Completed')
    if (returnedOrders.length > 0) {
      mainY = drawSectionHeading('Returned Orders', mainY, '#9a6200')
      returnedOrders.forEach((o, i) => {
        mainY = drawOrderRow(o, i, mainY, '#9a6200')
      })
      mainY += 16
    }

    // ── Page footers (applied after all pages are buffered) ───────────────────
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
          _id:             null,
          totalOrders:     { $sum: 1 },
          grossRevenue:    { $sum: '$subtotal'      },
          totalDiscount:   { $sum: '$totalDiscount' },
          couponDiscount:  { $sum: '$couponDiscount'},
          netRevenue:      { $sum: '$finalAmount'   },
          cancelledOrders: { $sum: { $cond: [{ $eq: ['$orderStatus',  'Cancelled'] }, 1, 0] } },
          returnedOrders:  { $sum: { $cond: [{ $eq: ['$returnStatus', 'Completed'] }, 1, 0] } },
          cancelledAmount: { $sum: { $cond: [{ $eq: ['$orderStatus',  'Cancelled'] }, '$finalAmount', 0] } },
          returnedAmount:  { $sum: { $cond: [{ $eq: ['$returnStatus', 'Completed'] }, '$finalAmount', 0] } },
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

    const wb = new ExcelJS.Workbook()
    wb.creator = 'Blush-Berry'
    wb.created = new Date()

    const headerFill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC93060' } }
    const cancelFill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFA0222A' } }
    const returnFill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9A6200' } }
    const altFill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF0F4' } }
    const headerFont   = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' }
    const bodyFont     = { size: 10, name: 'Calibri' }
    const boldFont     = { bold: true, size: 10, name: 'Calibri' }
    const pinkFont     = { bold: true, color: { argb: 'FFC93060' }, size: 10, name: 'Calibri' }
    const redFont      = { color: { argb: 'FFE8527A' }, size: 10, name: 'Calibri' }
    const cancelFont   = { bold: true, color: { argb: 'FFA0222A' }, size: 10, name: 'Calibri' }
    const returnFont   = { bold: true, color: { argb: 'FF9A6200' }, size: 10, name: 'Calibri' }
    const thinBorder   = {
      top:    { style: 'thin', color: { argb: 'FFF4D0DA' } },
      bottom: { style: 'thin', color: { argb: 'FFF4D0DA' } },
      left:   { style: 'thin', color: { argb: 'FFF4D0DA' } },
      right:  { style: 'thin', color: { argb: 'FFF4D0DA' } },
    }

    function styleHeaderRow(row, fill = headerFill) {
      row.eachCell(cell => {
        cell.fill      = fill
        cell.font      = headerFont
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
        cell.border    = thinBorder
      })
      row.height = 24
    }

    function styleDataRow(row, isAlt, amountCols = []) {
      row.eachCell({ includeEmpty: true }, (cell, colNum) => {
        if (isAlt) cell.fill = altFill
        cell.font      = amountCols.includes(colNum) ? boldFont : bodyFont
        cell.alignment = { vertical: 'middle', wrapText: false }
        cell.border    = thinBorder
      })
      row.height = 18
    }

    // Sheet 1: Summary
    const ws1 = wb.addWorksheet('Summary', { properties: { tabColor: { argb: 'FFC93060' } } })

    ws1.mergeCells('A1:C1')
    const titleCell = ws1.getCell('A1')
    titleCell.value     = 'Blush-Berry — Sales Report'
    titleCell.font      = { bold: true, size: 16, color: { argb: 'FFC93060' }, name: 'Calibri' }
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
    ws1.getRow(1).height = 32

    ws1.mergeCells('A2:C2')
    const periodCell = ws1.getCell('A2')
    periodCell.value     = `Period: ${start.toDateString()}  →  ${end.toDateString()}`
    periodCell.font      = { italic: true, size: 10, color: { argb: 'FF7A4A5E' }, name: 'Calibri' }
    periodCell.alignment = { horizontal: 'center', vertical: 'middle' }
    ws1.getRow(2).height = 18

    ws1.addRow([])

    ws1.columns = [
      { key: 'metric', width: 35 },
      { key: 'value',  width: 22 },
      { key: 'note',   width: 32 },
    ]

    styleHeaderRow(ws1.addRow(['Metric', 'Value', 'Note']))

    const summaryData = [
      ['Total Orders',                    stats.totalOrders,        `${start.toDateString()} to ${end.toDateString()}`],
      ['Gross Revenue (Rs.)',              stats.grossRevenue,       'Before any discounts'],
      ['Total Discounts (Rs.)',            stats.totalDiscount,      'Offer + coupon combined'],
      ['Coupon Discounts (Rs.)',           stats.couponDiscount,     'Coupon deductions only'],
      ['Item Offer Discounts (Rs.)',       Math.max(0, stats.totalDiscount - stats.couponDiscount), 'Product/category offers'],
      ['Net Revenue (Rs.)',                stats.netRevenue,         'After all deductions'],
      ['Cancelled Orders (count)',         stats.cancelledOrders,    'Fully cancelled orders'],
      ['Cancelled Orders Value (Rs.)',     stats.cancelledAmount,    'Revenue lost to cancellations'],
      ['Returned Orders (count)',          stats.returnedOrders,     'Fully returned orders'],
      ['Returned Orders Value (Rs.)',      stats.returnedAmount,     'Revenue lost to returns'],
    ]

    summaryData.forEach((rowData, i) => {
      const r = ws1.addRow(rowData)
      styleDataRow(r, i % 2 === 1, [2])
      r.getCell(1).font = bodyFont
      r.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' }

      if (rowData[0].includes('Net')) {
        r.getCell(2).font = pinkFont
      } else if (rowData[0].includes('Cancelled')) {
        r.getCell(2).font = cancelFont
      } else if (rowData[0].includes('Returned')) {
        r.getCell(2).font = returnFont
      } else if (rowData[0].includes('Discount')) {
        r.getCell(2).font = redFont
      }

      r.getCell(3).font = { italic: true, size: 9, color: { argb: 'FFB07090' }, name: 'Calibri' }
    })

    // Sheet 2: All Orders
    const ws2 = wb.addWorksheet('All Orders', { properties: { tabColor: { argb: 'FFC93060' } } })

    ws2.columns = [
      { header: 'S.No',          key: 'sno',      width: 7  },
      { header: 'Order ID',      key: 'orderId',  width: 24 },
      { header: 'Date',          key: 'date',     width: 16 },
      { header: 'Customer Name', key: 'customer', width: 24 },
      { header: 'Email',         key: 'email',    width: 30 },
      { header: 'Payment',       key: 'payment',  width: 14 },
      { header: 'Coupon Code',   key: 'coupon',   width: 14 },
      { header: 'Gross (Rs.)',   key: 'gross',    width: 16 },
      { header: 'Discount (Rs.)',key: 'discount', width: 16 },
      { header: 'Net (Rs.)',     key: 'net',      width: 16 },
      { header: 'Status',        key: 'status',   width: 14 },
    ]

    styleHeaderRow(ws2.getRow(1))

    orders.forEach((o, idx) => {
      const r = ws2.addRow({
        sno:      idx + 1,
        orderId:  o.orderId  || '—',
        date:     new Date(o.createdAt).toLocaleDateString('en-IN'),
        customer: o.userId?.name || o.shippingAddress?.name || '—',
        email:    o.userId?.email || o.shippingAddress?.email || '—',
        payment:  o.paymentMethod || '—',
        coupon:   o.couponCode    || '—',
        gross:    o.subtotal      || 0,
        discount: o.totalDiscount || 0,
        net:      o.finalAmount   || 0,
        status:   o.returnStatus === 'Completed' ? 'Returned' : (o.orderStatus || '—'),
      })
      styleDataRow(r, idx % 2 === 1, [8, 9, 10])

      ;[8, 9, 10].forEach(col => {
        r.getCell(col).alignment = { horizontal: 'right', vertical: 'middle' }
        r.getCell(col).numFmt    = '#,##0.00'
      })

      if ((o.totalDiscount || 0) > 0) r.getCell(9).font = redFont
      r.getCell(10).font = pinkFont

      const statusCell = r.getCell(11)
      const s = o.returnStatus === 'Completed' ? 'returned' : (o.orderStatus || '').toLowerCase()
      if      (s === 'delivered')  statusCell.font = { color: { argb: 'FF1A7A46' }, bold: true, size: 10, name: 'Calibri' }
      else if (s === 'cancelled')  statusCell.font = cancelFont
      else if (s === 'returned')   statusCell.font = returnFont
      else if (s === 'shipped')    statusCell.font = { color: { argb: 'FF6C27B0' }, bold: true, size: 10, name: 'Calibri' }
      else if (s === 'processing') statusCell.font = { color: { argb: 'FF9A6200' }, bold: true, size: 10, name: 'Calibri' }
      else                         statusCell.font = bodyFont
    })

    ws2.views      = [{ state: 'frozen', ySplit: 1 }]
    ws2.autoFilter = { from: 'A1', to: 'K1' }

    // Sheet 3: Cancelled Orders
    const ws3 = wb.addWorksheet('Cancelled Orders', { properties: { tabColor: { argb: 'FFA0222A' } } })

    ws3.columns = [
      { header: 'S.No',          key: 'sno',      width: 7  },
      { header: 'Order ID',      key: 'orderId',  width: 24 },
      { header: 'Date',          key: 'date',     width: 16 },
      { header: 'Customer Name', key: 'customer', width: 24 },
      { header: 'Email',         key: 'email',    width: 30 },
      { header: 'Payment',       key: 'payment',  width: 14 },
      { header: 'Gross (Rs.)',   key: 'gross',    width: 16 },
      { header: 'Discount (Rs.)',key: 'discount', width: 16 },
      { header: 'Net (Rs.)',     key: 'net',      width: 16 },
      { header: 'Cancel Reason', key: 'reason',   width: 30 },
    ]

    styleHeaderRow(ws3.getRow(1), cancelFill)

    const cancelledOrders = orders.filter(o => o.orderStatus === 'Cancelled')
    if (cancelledOrders.length === 0) {
      const r = ws3.addRow(['No cancelled orders in this period.', '', '', '', '', '', '', '', '', ''])
      ws3.mergeCells('A2:J2')
      r.getCell(1).font      = { italic: true, color: { argb: 'FFB07090' }, name: 'Calibri' }
      r.getCell(1).alignment = { horizontal: 'center' }
    } else {
      cancelledOrders.forEach((o, idx) => {
        const r = ws3.addRow({
          sno:      idx + 1,
          orderId:  o.orderId  || '—',
          date:     new Date(o.createdAt).toLocaleDateString('en-IN'),
          customer: o.userId?.name || o.shippingAddress?.name || '—',
          email:    o.userId?.email || o.shippingAddress?.email || '—',
          payment:  o.paymentMethod || '—',
          gross:    o.subtotal      || 0,
          discount: o.totalDiscount || 0,
          net:      o.finalAmount   || 0,
          reason:   o.cancelReason  || '—',
        })
        styleDataRow(r, idx % 2 === 1, [7, 8, 9])
        ;[7, 8, 9].forEach(col => {
          r.getCell(col).alignment = { horizontal: 'right', vertical: 'middle' }
          r.getCell(col).numFmt    = '#,##0.00'
        })
        r.getCell(9).font = cancelFont
      })
    }

    ws3.views      = [{ state: 'frozen', ySplit: 1 }]
    ws3.autoFilter = { from: 'A1', to: 'J1' }

    // Sheet 4: Returned Orders
    const ws4 = wb.addWorksheet('Returned Orders', { properties: { tabColor: { argb: 'FF9A6200' } } })

    ws4.columns = [
      { header: 'S.No',          key: 'sno',      width: 7  },
      { header: 'Order ID',      key: 'orderId',  width: 24 },
      { header: 'Date',          key: 'date',     width: 16 },
      { header: 'Customer Name', key: 'customer', width: 24 },
      { header: 'Email',         key: 'email',    width: 30 },
      { header: 'Payment',       key: 'payment',  width: 14 },
      { header: 'Gross (Rs.)',   key: 'gross',    width: 16 },
      { header: 'Discount (Rs.)',key: 'discount', width: 16 },
      { header: 'Net (Rs.)',     key: 'net',      width: 16 },
      { header: 'Return Reason', key: 'reason',   width: 30 },
      { header: 'Return Status', key: 'retStatus',width: 18 },
    ]

    styleHeaderRow(ws4.getRow(1), returnFill)

    const returnedOrders = orders.filter(o => o.returnStatus === 'Completed')
    if (returnedOrders.length === 0) {
      const r = ws4.addRow(['No returned orders in this period.', '', '', '', '', '', '', '', '', '', ''])
      ws4.mergeCells('A2:K2')
      r.getCell(1).font      = { italic: true, color: { argb: 'FFB07090' }, name: 'Calibri' }
      r.getCell(1).alignment = { horizontal: 'center' }
    } else {
      returnedOrders.forEach((o, idx) => {
        const r = ws4.addRow({
          sno:       idx + 1,
          orderId:   o.orderId  || '—',
          date:      new Date(o.createdAt).toLocaleDateString('en-IN'),
          customer:  o.userId?.name || o.shippingAddress?.name || '—',
          email:     o.userId?.email || o.shippingAddress?.email || '—',
          payment:   o.paymentMethod || '—',
          gross:     o.subtotal      || 0,
          discount:  o.totalDiscount || 0,
          net:       o.finalAmount   || 0,
          reason:    o.returnReason  || '—',
          retStatus: o.returnStatus  || '—',
        })
        styleDataRow(r, idx % 2 === 1, [7, 8, 9])
        ;[7, 8, 9].forEach(col => {
          r.getCell(col).alignment = { horizontal: 'right', vertical: 'middle' }
          r.getCell(col).numFmt    = '#,##0.00'
        })
        r.getCell(9).font = returnFont
      })
    }

    ws4.views      = [{ state: 'frozen', ySplit: 1 }]
    ws4.autoFilter = { from: 'A1', to: 'K1' }

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