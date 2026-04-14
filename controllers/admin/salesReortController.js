const Order       = require('../../models/user/orderModel')
const ExcelJS     = require('exceljs')
const PDFDocument = require('pdfkit')



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

    const stats = summary || {
      totalOrders: 0, grossRevenue: 0,
      totalDiscount: 0, couponDiscount: 0, netRevenue: 0
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
          _id:      groupId,
          gross:    { $sum: '$subtotal' },
          net:      { $sum: '$finalAmount' },
          discount: { $sum: '$totalDiscount' }
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
      user:          req.session.admin || null,
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
    res.status(500).render('error', { message: 'Could not load sales report.' })
  }
}



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

    const stats = summary || {
      totalOrders: 0, grossRevenue: 0,
      totalDiscount: 0, couponDiscount: 0, netRevenue: 0
    }

    const orders = await Order.find(match)
      .populate('userId', 'name email')
      .sort({ createdAt: -1 })
      .limit(200)
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

    
    const doc = new PDFDocument({ margin: 40, size: 'A4', layout: 'landscape' })
    res.setHeader('Content-Type', 'application/pdf')
    res.setHeader('Content-Disposition', `attachment; filename="sales-report-${type}-${Date.now()}.pdf"`)
    doc.pipe(res)

    const MARGIN   = 40
    const PAGE_W   = doc.page.width                
    const PAGE_H   = doc.page.height               
    const USABLE_W = PAGE_W - MARGIN * 2          
    const LEFT     = MARGIN

   
    doc.rect(0, 0, PAGE_W, 54).fill('#c93060')
    doc.fontSize(20).font('Helvetica-Bold').fillColor('#ffffff')
       .text('Blush-Berry', LEFT, 14, { lineBreak: false })
    doc.fontSize(10).font('Helvetica').fillColor('rgba(255,255,255,0.85)')
       .text('Sales Report', LEFT, 36, { lineBreak: false })
    doc.fontSize(9).font('Helvetica').fillColor('rgba(255,255,255,0.75)')
       .text(`Generated: ${new Date().toLocaleString('en-IN')}`, LEFT, 20, { align: 'right', width: USABLE_W, lineBreak: false })
    doc.y = 70

    
    doc.fontSize(9).font('Helvetica').fillColor('#555555')
       .text(`Period: ${start.toDateString()}  →  ${end.toDateString()}`, LEFT, doc.y)
    doc.moveDown(0.8)

 
    const CARD_GAP = 8
    const CARD_W   = (USABLE_W - CARD_GAP * 3) / 4   
    const CARD_H   = 52
    const cardY    = doc.y

    const summaryCards = [
      { label: 'Total Orders',    value: String(stats.totalOrders) },
      { label: 'Gross Revenue',   value: inr(stats.grossRevenue)   },
      { label: 'Total Discounts', value: inr(stats.totalDiscount)  },
      { label: 'Net Revenue',     value: inr(stats.netRevenue)     },
    ]

    summaryCards.forEach((c, i) => {
      const x = LEFT + i * (CARD_W + CARD_GAP)
      doc.roundedRect(x, cardY, CARD_W, CARD_H, 6).fill('#fdf0f4')
      doc.fontSize(8).font('Helvetica').fillColor('#b07090')
         .text(c.label, x + 8, cardY + 8, { width: CARD_W - 16, lineBreak: false })
      doc.fontSize(12).font('Helvetica-Bold').fillColor('#c93060')
         .text(c.value, x + 8, cardY + 24, { width: CARD_W - 16, lineBreak: false })
    })

    doc.y = cardY + CARD_H + 18

    // ── Orders table ─────────────────────────────────────────────────────────
    // Column widths must sum exactly to USABLE_W (761.89)
    // We define them as fractions then scale to fill perfectly.
    const RAW_COLS = [
      { label: '#',        raw: 22,  align: 'center' },
      { label: 'Order ID', raw: 110, align: 'left'   },
      { label: 'Date',     raw: 62,  align: 'left'   },
      { label: 'Customer', raw: 100, align: 'left'   },
      { label: 'Payment',  raw: 58,  align: 'left'   },
      { label: 'Coupon',   raw: 55,  align: 'left'   },
      { label: 'Gross',    raw: 72,  align: 'right'  },
      { label: 'Discount', raw: 72,  align: 'right'  },
      { label: 'Net',      raw: 72,  align: 'right'  },
      { label: 'Status',   raw: 60,  align: 'left'   },
    ]
    const rawTotal = RAW_COLS.reduce((s, c) => s + c.raw, 0)
    const ORDER_COLS = RAW_COLS.map(c => ({
      ...c,
      width: Math.floor((c.raw / rawTotal) * USABLE_W)
    }))
    // Give any rounding remainder to the last column
    const orderColsTotal = ORDER_COLS.reduce((s, c) => s + c.width, 0)
    ORDER_COLS[ORDER_COLS.length - 1].width += USABLE_W - orderColsTotal

    const ROW_H   = 18
    const THEAD_H = 20
    const FOOT_H  = 30   // reserved space at bottom of page

    // Draw a table header row, returns y after header
    function drawOrderHeader(y) {
      doc.rect(LEFT, y, USABLE_W, THEAD_H).fill('#c93060')
      let x = LEFT
      ORDER_COLS.forEach(c => {
        doc.fontSize(7.5).font('Helvetica-Bold').fillColor('#ffffff')
           .text(c.label, x + 3, y + 6, {
             width:    c.width - 6,
             align:    c.align,
             lineBreak: false,
             ellipsis: true
           })
        x += c.width
      })
      return y + THEAD_H
    }

    // Draw one order row, returns y after row
    function drawOrderRow(order, idx, y) {
      // Page break check
      if (y + ROW_H > PAGE_H - FOOT_H) {
        doc.addPage()
        y = MARGIN
        y = drawOrderHeader(y)
      }

      const bg = idx % 2 === 0 ? '#ffffff' : '#fdf5f7'
      doc.rect(LEFT, y, USABLE_W, ROW_H).fill(bg)

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
        order.orderStatus || '—',
      ]

      let x = LEFT
      cells.forEach((cell, i) => {
        const c = ORDER_COLS[i]
        doc.fontSize(7).font('Helvetica').fillColor('#3a1a2e')
           .text(cell, x + 3, y + 5, {
             width:    c.width - 6,
             align:    c.align,
             lineBreak: false,
             ellipsis: true
           })
        x += c.width
      })

     
      doc.moveTo(LEFT, y + ROW_H)
         .lineTo(LEFT + USABLE_W, y + ROW_H)
         .strokeColor('#f4d0da').lineWidth(0.3).stroke()

      return y + ROW_H
    }

    // Section title
    doc.fontSize(11).font('Helvetica-Bold').fillColor('#3a1a2e')
       .text('Order Details', LEFT, doc.y)
    doc.moveDown(0.4)

    let rowY = drawOrderHeader(doc.y)
    orders.forEach((o, i) => { rowY = drawOrderRow(o, i, rowY) })
    doc.y = rowY + 16

   
    if (couponBreakdown.length > 0) {
      // Page break if not enough room
      if (doc.y + THEAD_H + ROW_H * couponBreakdown.length + 40 > PAGE_H - FOOT_H) {
        doc.addPage()
        doc.y = MARGIN
      }

      doc.fontSize(11).font('Helvetica-Bold').fillColor('#3a1a2e')
         .text('Coupon Performance', LEFT, doc.y)
      doc.moveDown(0.4)

     
      const RAW_CP = [
        { label: 'Coupon Code',              raw: 160, align: 'left'   },
        { label: 'Uses',                     raw: 60,  align: 'center' },
        { label: 'Total Deducted',           raw: 140, align: 'right'  },
        { label: 'Net Revenue (after coupon)',raw: 140, align: 'right'  },
      ]
      const cpRawTotal = RAW_CP.reduce((s, c) => s + c.raw, 0)
      const CP_COLS = RAW_CP.map(c => ({
        ...c,
        width: Math.floor((c.raw / cpRawTotal) * USABLE_W)
      }))
      const cpColsTotal = CP_COLS.reduce((s, c) => s + c.width, 0)
      CP_COLS[CP_COLS.length - 1].width += USABLE_W - cpColsTotal

      // Header
      const cpHeaderY = doc.y
      doc.rect(LEFT, cpHeaderY, USABLE_W, THEAD_H).fill('#c93060')
      let cx = LEFT
      CP_COLS.forEach(c => {
        doc.fontSize(7.5).font('Helvetica-Bold').fillColor('#ffffff')
           .text(c.label, cx + 3, cpHeaderY + 6, {
             width:    c.width - 6,
             align:    c.align,
             lineBreak: false,
             ellipsis: true
           })
        cx += c.width
      })

      let cpY = cpHeaderY + THEAD_H

      couponBreakdown.forEach((c, i) => {
        if (cpY + ROW_H > PAGE_H - FOOT_H) {
          doc.addPage()
          cpY = MARGIN
          doc.rect(LEFT, cpY, USABLE_W, THEAD_H).fill('#c93060')
          let hx = LEFT
          CP_COLS.forEach(col => {
            doc.fontSize(7.5).font('Helvetica-Bold').fillColor('#ffffff')
               .text(col.label, hx + 3, cpY + 6, {
                 width:    col.width - 6,
                 align:    col.align,
                 lineBreak: false,
                 ellipsis: true
               })
            hx += col.width
          })
          cpY += THEAD_H
        }

        const bg = i % 2 === 0 ? '#ffffff' : '#fdf5f7'
        doc.rect(LEFT, cpY, USABLE_W, ROW_H).fill(bg)

        const cells = [c._id, String(c.uses), inr(c.totalDeducted), inr(c.netRevenue)]
        let x = LEFT
        cells.forEach((cell, j) => {
          const col = CP_COLS[j]
          doc.fontSize(8).font('Helvetica').fillColor('#3a1a2e')
             .text(cell, x + 3, cpY + 5, {
               width:    col.width - 6,
               align:    col.align,
               lineBreak: false,
               ellipsis: true
             })
          x += col.width
        })

        doc.moveTo(LEFT, cpY + ROW_H)
           .lineTo(LEFT + USABLE_W, cpY + ROW_H)
           .strokeColor('#f4d0da').lineWidth(0.3).stroke()

        cpY += ROW_H
      })

      doc.y = cpY + 12
    }

    const range = doc.bufferedPageRange()
    for (let i = range.start; i < range.start + range.count; i++) {
      doc.switchToPage(i)
      doc.fontSize(7.5).font('Helvetica').fillColor('#aaaaaa')
         .text(
           `Blush-Berry — Confidential Sales Report   |   Page ${i - range.start + 1} of ${range.count}`,
           LEFT,
           PAGE_H - 24,
           { width: USABLE_W, align: 'center', lineBreak: false }
         )
    }

    doc.end()

  } catch (err) {
    console.error('downloadPDF error:', err)
    res.status(500).json({ success: false, message: 'Could not generate PDF.' })
  }
}



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

    const stats = summary || {
      totalOrders: 0, grossRevenue: 0,
      totalDiscount: 0, couponDiscount: 0, netRevenue: 0
    }

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
    wb.created = new Date()

    const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC93060' } }
    const altFill    = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF0F4' } }
    const headerFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11, name: 'Calibri' }
    const bodyFont   = { size: 10, name: 'Calibri' }
    const boldFont   = { bold: true, size: 10, name: 'Calibri' }
    const pinkFont   = { bold: true, color: { argb: 'FFC93060' }, size: 10, name: 'Calibri' }
    const redFont    = { color: { argb: 'FFE8527A' }, size: 10, name: 'Calibri' }
    const thinBorder = {
      top:    { style: 'thin', color: { argb: 'FFF4D0DA' } },
      bottom: { style: 'thin', color: { argb: 'FFF4D0DA' } },
      left:   { style: 'thin', color: { argb: 'FFF4D0DA' } },
      right:  { style: 'thin', color: { argb: 'FFF4D0DA' } },
    }

    function styleHeaderRow(row) {
      row.eachCell(cell => {
        cell.fill      = headerFill
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
      { key: 'metric', width: 30 },
      { key: 'value',  width: 22 },
      { key: 'note',   width: 28 },
    ]

    const hRow = ws1.addRow(['Metric', 'Value', 'Note'])
    styleHeaderRow(hRow)

    const summaryData = [
      ['Total Orders',               stats.totalOrders,        `${start.toDateString()} to ${end.toDateString()}`],
      ['Gross Revenue (Rs.)',         stats.grossRevenue,        'Before any discounts'],
      ['Total Discounts (Rs.)',       stats.totalDiscount,       'Offer + coupon combined'],
      ['Coupon Discounts (Rs.)',      stats.couponDiscount,      'Coupon deductions only'],
      ['Item Offer Discounts (Rs.)',  Math.max(0, stats.totalDiscount - stats.couponDiscount), 'Product/category offers'],
      ['Net Revenue (Rs.)',           stats.netRevenue,          'After all deductions'],
    ]
    summaryData.forEach((rowData, i) => {
      const r = ws1.addRow(rowData)
      styleDataRow(r, i % 2 === 1, [2])
      r.getCell(1).font = bodyFont
      if (rowData[0].includes('Net')) {
        r.getCell(2).font      = pinkFont
        r.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' }
      } else if (rowData[0].includes('Discount')) {
        r.getCell(2).font      = redFont
        r.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' }
      } else {
        r.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' }
      }
      r.getCell(3).font = { italic: true, size: 9, color: { argb: 'FFB07090' }, name: 'Calibri' }
    })

    // ── Sheet 2: Orders ───────────────────────────────────────────────────────
    const ws2 = wb.addWorksheet('Orders', { properties: { tabColor: { argb: 'FFC93060' } } })

    ws2.columns = [
      { header: 'S.No',          key: 'sno',      width: 7  },
      { header: 'Order ID',      key: 'orderId',  width: 24 },
      { header: 'Date',          key: 'date',     width: 16 },
      { header: 'Customer Name', key: 'customer', width: 24 },
      { header: 'Email',         key: 'email',    width: 28 },
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
        coupon:   o.couponCode || '—',
        gross:    o.subtotal      || 0,
        discount: o.totalDiscount || 0,
        net:      o.finalAmount   || 0,
        status:   o.orderStatus   || '—',
      })
      styleDataRow(r, idx % 2 === 1, [8, 9, 10])

      ;[8, 9, 10].forEach(col => {
        r.getCell(col).alignment = { horizontal: 'right', vertical: 'middle' }
        r.getCell(col).numFmt    = '#,##0.00'
      })

      if ((o.totalDiscount || 0) > 0) r.getCell(9).font = redFont
      r.getCell(10).font = pinkFont

      const statusCell = r.getCell(11)
      const s = (o.orderStatus || '').toLowerCase()
      if      (s === 'delivered')   statusCell.font = { color: { argb: 'FF1A7A46' }, bold: true, size: 10, name: 'Calibri' }
      else if (s === 'cancelled')   statusCell.font = { color: { argb: 'FFA0222A' }, bold: true, size: 10, name: 'Calibri' }
      else if (s === 'shipped')     statusCell.font = { color: { argb: 'FF6C27B0' }, bold: true, size: 10, name: 'Calibri' }
      else if (s === 'processing')  statusCell.font = { color: { argb: 'FF9A6200' }, bold: true, size: 10, name: 'Calibri' }
      else statusCell.font = bodyFont
    })

    ws2.views      = [{ state: 'frozen', ySplit: 1 }]
    ws2.autoFilter = { from: 'A1', to: 'K1' }

    // ── Sheet 3: Coupon Performance ───────────────────────────────────────────
    const ws3 = wb.addWorksheet('Coupon Performance', { properties: { tabColor: { argb: 'FFC93060' } } })

    ws3.columns = [
      { header: 'Coupon Code',           key: 'code',     width: 20 },
      { header: 'Total Uses',            key: 'uses',     width: 14 },
      { header: 'Total Deducted (Rs.)',  key: 'deducted', width: 24 },
      { header: 'Net Revenue (Rs.)',     key: 'revenue',  width: 22 },
    ]

    styleHeaderRow(ws3.getRow(1))

    if (couponBreakdown.length === 0) {
      const r = ws3.addRow(['No coupon usage in this period.', '', '', ''])
      ws3.mergeCells('A2:D2')
      r.getCell(1).font      = { italic: true, color: { argb: 'FFB07090' }, name: 'Calibri' }
      r.getCell(1).alignment = { horizontal: 'center' }
    } else {
      couponBreakdown.forEach((c, i) => {
        const r = ws3.addRow({
          code:     c._id,
          uses:     c.uses,
          deducted: c.totalDeducted,
          revenue:  c.netRevenue,
        })
        styleDataRow(r, i % 2 === 1, [3, 4])
        r.getCell(1).font      = { bold: true, color: { argb: 'FFC93060' }, size: 10, name: 'Calibri' }
        r.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' }
        r.getCell(3).alignment = { horizontal: 'right',  vertical: 'middle' }
        r.getCell(3).numFmt    = '#,##0.00'
        r.getCell(3).font      = redFont
        r.getCell(4).alignment = { horizontal: 'right',  vertical: 'middle' }
        r.getCell(4).numFmt    = '#,##0.00'
        r.getCell(4).font      = pinkFont
      })
    }

    ws3.views = [{ state: 'frozen', ySplit: 1 }]

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    res.setHeader('Content-Disposition', `attachment; filename="blushberry-sales-${type}-${Date.now()}.xlsx"`)
    await wb.xlsx.write(res)
    res.end()

  } catch (err) {
    console.error('downloadExcel error:', err)
    res.status(500).json({ success: false, message: 'Could not generate Excel.' })
  }
}

module.exports = { loadSalesReport, downloadPDF, downloadExcel }