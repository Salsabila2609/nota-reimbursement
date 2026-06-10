/**
 * excel-generator.ts
 * Lokasi: src/lib/excel-generator.ts
 *
 * Generate laporan rekap BBM driver.
 * - Nomor urut berdasarkan tanggal struk (bill_date), bukan tanggal submit
 * - Nomor konsisten dengan laporan PDF
 * - Kategori "lainnya" ditampilkan dengan deskripsi dari driver
 * - Warna palette sesuai IOH brand (#ED1C24, #FFCB05, #32BCAD, #C6168D, #4D4D4F)
 *
 * Dependency: exceljs  →  npm install exceljs
 */

import * as ExcelJS from 'exceljs'

// ─── Types ────────────────────────────────────────────────────────────────────

export type BillEntry = {
  no: number          // nomor urut berdasarkan tanggal struk
  date: string        // ISO date — WAJIB adalah bill_date (tanggal struk), bukan submission_date
  category: string    // "BIAYA BBM" | "BIAYA TOLL" | "BIAYA PARKIR" | "LAINNYA: ..."
  description: string // e.g. "BIAYA BBM OPERS MOBIL B 1960 RZP"
  amount: number      // nominal in Rp (integer)
}

export type DriverData = {
  name: string
  plateNumber: string
  vehicleType: string
  entries: BillEntry[]  // harus sudah diurutkan ascending by bill_date sebelum dikirim
}

export type RekapExcelParams = {
  title: string
  subtitle: string
  drivers: DriverData[]
  cashAdvance?: number
  reportDate?: string
  createdBy?: string
  createdByTitle?: string
  approvedBy?: string
  approvedByTitle?: string
  acknowledgedBy?: string
  acknowledgedByTitle?: string
  month: string
}

// ─── IOH Brand Palette (ARGB — prefix FF = fully opaque) ─────────────────────

const COLOR = {
  // IOH brand
  iohRed:       'FFED1C24',   // #ED1C24
  iohYellow:    'FFFFCB05',   // #FFCB05
  iohTeal:      'FF32BCAD',   // #32BCAD
  iohMagenta:   'FFC6168D',   // #C6168D
  iohCharcoal:  'FF4D4D4F',   // #4D4D4F

  // Document structural colors (remain dark for readability)
  headerBg:     'FF1A1A1C',   // near-black header
  headerText:   'FFFFFFFF',
  subheaderBg:  'FF2D2D30',
  driverBg:     'FF1E3A5F',   // dark navy (secondary info)
  driverText:   'FFFFFFFF',
  totalBg:      'FF0F2D4A',
  totalText:    'FFFFFFFF',
  altRow:       'FFF5F5F7',   // IOH bg color
  border:       'FFD0D0D4',
  grandTotalBg: 'FF0A1628',
  summaryBg:    'FFD6E4F0',

  // Accent overrides matching IOH
  accentYellow: 'FFFFCB05',   // IOH yellow (was FFC000)
  accentTeal:   'FF32BCAD',   // IOH teal for subtotals
}

const FONT_NAME = 'Arial'

function border(style: ExcelJS.BorderStyle = 'thin'): ExcelJS.Borders {
  const s = { style, color: { argb: COLOR.border } }
  const b: any = { top: s, bottom: s, left: s, right: s, diagonal: undefined }
  return b as ExcelJS.Borders
}

function centerAlign(wrap = false): Partial<ExcelJS.Alignment> {
  return { horizontal: 'center', vertical: 'middle', wrapText: wrap }
}
function leftAlign(wrap = false): Partial<ExcelJS.Alignment> {
  return { horizontal: 'left', vertical: 'middle', wrapText: wrap }
}
function rightAlign(): Partial<ExcelJS.Alignment> {
  return { horizontal: 'right', vertical: 'middle' }
}
function currency(amount: number): string {
  return new Intl.NumberFormat('id-ID').format(amount)
}

// ─── Fungsi utama ─────────────────────────────────────────────────────────────

export async function generateRekapExcel(params: RekapExcelParams): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'Sistem Reimburse IOH'
  workbook.created = new Date()

  await buildSheetPTG(workbook, params)

  for (const driver of params.drivers) {
    await buildSheetDriver(workbook, driver, params.title, params.subtitle)
  }

  const buffer = await workbook.xlsx.writeBuffer()
  return Buffer.from(buffer)
}

// ─────────────────────────────────────────────────────────────────────────────
// SHEET PTG — Layout horizontal multi-driver
// ─────────────────────────────────────────────────────────────────────────────

async function buildSheetPTG(wb: ExcelJS.Workbook, params: RekapExcelParams) {
  const {
    title, subtitle, drivers, cashAdvance,
    reportDate, createdBy, createdByTitle,
    approvedBy, approvedByTitle,
    acknowledgedBy, acknowledgedByTitle,
  } = params

  const ws = wb.addWorksheet('PTG', {
    pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1 },
    views: [{ state: 'normal', zoomScale: 85 }],
  })

  const BLOCK_COLS = 7

  ws.getColumn(1).width = 5
  ws.getColumn(2).width = 14
  ws.getColumn(3).width = 18   // wider for LAINNYA: ...
  ws.getColumn(4).width = 40
  ws.getColumn(5).width = 5
  ws.getColumn(6).width = 16
  ws.getColumn(7).width = 3

  drivers.forEach((_, di) => {
    const base = 8 + di * BLOCK_COLS
    ws.getColumn(base).width = 5
    ws.getColumn(base + 1).width = 12
    ws.getColumn(base + 2).width = 18
    ws.getColumn(base + 3).width = 40
    ws.getColumn(base + 4).width = 5
    ws.getColumn(base + 5).width = 16
    ws.getColumn(base + 6).width = 3
  })

  // ── Row 1: Title ──────────────────────────────────────────────────────────
  const row1 = ws.addRow([])
  row1.height = 28

  setCellStyle(ws, row1.number, 1, title, {
    font: { bold: true, size: 11, color: { argb: COLOR.headerText }, name: FONT_NAME },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.headerBg } },
    alignment: centerAlign(),
    border: border(),
  })
  ws.mergeCells(row1.number, 1, row1.number, 6)

  // IOH red accent cell on left edge of title
  const titleAccentCell = ws.getCell(row1.number, 1)
  // (merged cells — accent handled by left border color below)

  drivers.forEach((_, di) => {
    const base = 8 + di * BLOCK_COLS
    setCellStyle(ws, row1.number, base, title, {
      font: { bold: true, size: 11, color: { argb: COLOR.headerText }, name: FONT_NAME },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.headerBg } },
      alignment: centerAlign(),
      border: border(),
    })
    ws.mergeCells(row1.number, base, row1.number, base + 5)
  })

  // ── Row 2: Subtitle ───────────────────────────────────────────────────────
  const row2 = ws.addRow([])
  row2.height = 22

  setCellStyle(ws, row2.number, 1, subtitle, {
    font: { size: 9, color: { argb: COLOR.headerText }, name: FONT_NAME },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.subheaderBg } },
    alignment: centerAlign(true),
    border: border(),
  })
  ws.mergeCells(row2.number, 1, row2.number, 6)

  drivers.forEach((_, di) => {
    const base = 8 + di * BLOCK_COLS
    setCellStyle(ws, row2.number, base, subtitle, {
      font: { size: 9, color: { argb: COLOR.headerText }, name: FONT_NAME },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.subheaderBg } },
      alignment: centerAlign(true),
      border: border(),
    })
    ws.mergeCells(row2.number, base, row2.number, base + 5)
  })

  // ── Row 3: JENIS PEMAKAIAN labels ─────────────────────────────────────────
  const row3 = ws.addRow([])
  row3.height = 20

  ;[1, 2, 3, 4, 5, 6].forEach(c => {
    setCellStyle(ws, row3.number, c, '', {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EEF5' } },
      border: border(),
    })
  })

  drivers.forEach((_, di) => {
    const base = 8 + di * BLOCK_COLS
    setCellStyle(ws, row3.number, base + 1, 'JENIS PEMAKAIAN', {
      font: { bold: true, size: 9, name: FONT_NAME },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EEF5' } },
      alignment: centerAlign(),
      border: border(),
    })
    ws.mergeCells(row3.number, base + 1, row3.number, base + 2)
    setCellStyle(ws, row3.number, base + 3, 'Data Keseluruhan', {
      font: { bold: true, size: 9, name: FONT_NAME },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD6E4F0' } },
      alignment: centerAlign(),
      border: border(),
    })
    ws.mergeCells(row3.number, base + 3, row3.number, base + 5)
    ;[base].forEach(c => {
      setCellStyle(ws, row3.number, c, '', {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EEF5' } },
        border: border(),
      })
    })
  })

  // ── Row 4: Summary header + driver info header ─────────────────────────────
  const row4 = ws.addRow([])
  row4.height = 20

  const summaryHeaders = ['NO.', 'NAMA', 'BIAYA', 'DESCRIPTION', 'Rp.', 'NOMINAL']
  summaryHeaders.forEach((h, i) => {
    setCellStyle(ws, row4.number, i + 1, h, {
      font: { bold: true, size: 9, color: { argb: COLOR.headerText }, name: FONT_NAME },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.headerBg } },
      alignment: centerAlign(),
      border: border(),
    })
  })

  drivers.forEach((driver, di) => {
    const base = 8 + di * BLOCK_COLS
    setCellStyle(ws, row4.number, base + 3, 'NAMA DRIVER', {
      font: { bold: true, size: 9, name: FONT_NAME, color: { argb: COLOR.headerText } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.driverBg } },
      alignment: centerAlign(), border: border(),
    })
    setCellStyle(ws, row4.number, base + 4, 'PLAT NOMOR', {
      font: { bold: true, size: 9, name: FONT_NAME, color: { argb: COLOR.headerText } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.driverBg } },
      alignment: centerAlign(), border: border(),
    })
    setCellStyle(ws, row4.number, base + 5, 'TYPE', {
      font: { bold: true, size: 9, name: FONT_NAME, color: { argb: COLOR.headerText } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.driverBg } },
      alignment: centerAlign(), border: border(),
    })
    ;[base, base + 1, base + 2].forEach(c => {
      setCellStyle(ws, row4.number, c, '', {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.driverBg } },
        border: border(),
      })
    })
  })

  // ── Row 5: Detail header + driver values ──────────────────────────────────
  const row5 = ws.addRow([])
  row5.height = 20

  for (let c = 1; c <= 6; c++) {
    setCellStyle(ws, row5.number, c, '', {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A3A5C' } },
      border: border(),
    })
  }

  drivers.forEach((driver, di) => {
    const base = 8 + di * BLOCK_COLS
    const infoStyle = {
      font: { bold: true, size: 9, name: FONT_NAME, color: { argb: 'FFFFFFFF' } },
      fill: { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FF1A3A5C' } },
      alignment: centerAlign(), border: border(),
    }
    setCellStyle(ws, row5.number, base + 3, driver.name, infoStyle)
    setCellStyle(ws, row5.number, base + 4, driver.plateNumber, infoStyle)
    setCellStyle(ws, row5.number, base + 5, driver.vehicleType, infoStyle)
    ;[base, base + 1, base + 2].forEach(c => {
      setCellStyle(ws, row5.number, c, '', {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A3A5C' } },
        border: border(),
      })
    })
  })

  // ── Row 6: Column headers for driver detail ────────────────────────────────
  const row6 = ws.addRow([])
  row6.height = 20

  ;[1, 2, 3, 4, 5, 6].forEach(c => { ws.getCell(row6.number, c).border = border() })

  drivers.forEach((_, di) => {
    const base = 8 + di * BLOCK_COLS
    const detailHdrs = ['NO.', 'TGL STRUK', 'BIAYA', 'DESCRIPTION', 'Rp.', 'NOMINAL']
    detailHdrs.forEach((h, i) => {
      setCellStyle(ws, row6.number, base + i, h, {
        font: { bold: true, size: 8, color: { argb: COLOR.headerText }, name: FONT_NAME },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3A3A3E' } },
        alignment: centerAlign(), border: border(),
      })
    })
  })

  // ── Build summary rows ─────────────────────────────────────────────────────
  const summaryRows: {
    no: number; nama: string; biaya: string; description: string; amount: number; isTotal?: boolean
  }[] = []
  let summaryNo = 1

  for (const driver of drivers) {
    const grouped: Record<string, number> = {}
    for (const e of driver.entries) {
      const cat = e.category  // already normalized
      grouped[cat] = (grouped[cat] || 0) + e.amount
    }
    let isFirst = true
    for (const cat of Object.keys(grouped)) {
      summaryRows.push({
        no: summaryNo++, nama: isFirst ? driver.name : '',
        biaya: cat, description: `${cat} OPERS MOBIL`, amount: grouped[cat],
      })
      isFirst = false
    }
    const total = driver.entries.reduce((s, e) => s + e.amount, 0)
    summaryRows.push({ no: 0, nama: driver.name, biaya: '', description: 'TOTAL', amount: total, isTotal: true })
  }

  const grandByCategory: Record<string, number> = {}
  let grandTotal = 0
  for (const driver of drivers) {
    for (const e of driver.entries) {
      grandByCategory[e.category] = (grandByCategory[e.category] || 0) + e.amount
      grandTotal += e.amount
    }
  }
  let grandNo = 1
  for (const cat of Object.keys(grandByCategory)) {
    summaryRows.push({ no: grandNo++, nama: 'ALL', biaya: cat, description: `${cat} OPERS MOBIL`, amount: grandByCategory[cat] })
  }

  const maxEntries = Math.max(...drivers.map(d => d.entries.length), 0)
  const DATA_START_ROW = row6.number + 1
  const totalDataRows = Math.max(summaryRows.length, maxEntries) + 5

  for (let ri = 0; ri < totalDataRows; ri++) {
    const row = ws.addRow([])
    row.height = 22
    const rowNum = DATA_START_ROW + ri
    const altBg = ri % 2 === 0 ? 'FFFFFFFF' : COLOR.altRow

    const sRow = summaryRows[ri]
    if (sRow) {
      const isTot = sRow.isTotal
      const bg = isTot ? COLOR.totalBg : altBg
      const fg = isTot ? COLOR.accentYellow : 'FF1A1A1C'
      const bold = !!isTot

      setCellStyle(ws, rowNum, 1, isTot ? '' : (sRow.no || ''), {
        font: { size: 9, name: FONT_NAME, bold, color: { argb: isTot ? 'FFFFFFFF' : 'FF1A1A1C' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } },
        alignment: centerAlign(), border: border(),
      })
      setCellStyle(ws, rowNum, 2, sRow.nama, {
        font: { size: 9, name: FONT_NAME, bold, color: { argb: fg } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } },
        alignment: leftAlign(), border: border(),
      })
      setCellStyle(ws, rowNum, 3, isTot ? '' : sRow.biaya, {
        font: { size: 9, name: FONT_NAME, bold, color: { argb: isTot ? 'FFFFFFFF' : 'FF1A1A1C' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } },
        alignment: leftAlign(true), border: border(),
      })
      setCellStyle(ws, rowNum, 4, sRow.description, {
        font: { size: 9, name: FONT_NAME, bold, color: { argb: isTot ? 'FFFFFFFF' : 'FF1A1A1C' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } },
        alignment: leftAlign(true), border: border(),
      })
      setCellStyle(ws, rowNum, 5, 'Rp.', {
        font: { size: 9, name: FONT_NAME, bold, color: { argb: isTot ? 'FFFFFFFF' : 'FF1A1A1C' } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } },
        alignment: centerAlign(), border: border(),
      })
      setCellStyle(ws, rowNum, 6, sRow.amount, {
        font: { size: 9, name: FONT_NAME, bold, color: { argb: fg } },
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } },
        alignment: rightAlign(), border: border(), numFmt: '#,##0',
      })
    } else {
      for (let c = 1; c <= 6; c++) {
        setCellStyle(ws, rowNum, c, '', {
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: altBg } },
          border: border(),
        })
      }
    }

    // Per-driver entries
    drivers.forEach((driver, di) => {
      const base = 8 + di * BLOCK_COLS
      const entry = driver.entries[ri]

      if (entry) {
        setCellStyle(ws, rowNum, base, entry.no, {
          font: { size: 8, name: FONT_NAME },
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: altBg } },
          alignment: centerAlign(), border: border(),
        })
        // ── Kolom tanggal: gunakan tanggal struk (bill_date) yang sudah dimasukkan ke entry.date ──
        setCellStyle(ws, rowNum, base + 1, formatExcelDate(entry.date), {
          font: { size: 8, name: FONT_NAME },
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: altBg } },
          alignment: centerAlign(), border: border(),
        })
        setCellStyle(ws, rowNum, base + 2, entry.category, {
          font: { size: 8, name: FONT_NAME },
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: altBg } },
          alignment: leftAlign(true), border: border(),
        })
        setCellStyle(ws, rowNum, base + 3, entry.description, {
          font: { size: 8, name: FONT_NAME },
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: altBg } },
          alignment: leftAlign(true), border: border(),
        })
        setCellStyle(ws, rowNum, base + 4, 'Rp.', {
          font: { size: 8, name: FONT_NAME },
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: altBg } },
          alignment: centerAlign(), border: border(),
        })
        setCellStyle(ws, rowNum, base + 5, entry.amount, {
          font: { size: 8, name: FONT_NAME },
          fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: altBg } },
          alignment: rightAlign(), border: border(), numFmt: '#,##0',
        })
      } else {
        for (let c = base; c <= base + 5; c++) {
          setCellStyle(ws, rowNum, c, '', {
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: altBg } },
            border: border(),
          })
        }
      }
    })
  }

  // ── Subtotals per driver ───────────────────────────────────────────────────
  const subTotalRowNum = DATA_START_ROW + totalDataRows

  drivers.forEach((driver, di) => {
    const base = 8 + di * BLOCK_COLS
    const grouped: Record<string, number> = {}
    const driverTotal = driver.entries.reduce((s, e) => {
      grouped[e.category] = (grouped[e.category] || 0) + e.amount
      return s + e.amount
    }, 0)

    const subLabelStyle = {
      font: { size: 8, name: FONT_NAME, bold: true, color: { argb: COLOR.headerText } },
      fill: { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: COLOR.totalBg } },
      alignment: leftAlign() as Partial<ExcelJS.Alignment>,
      border: border(),
    }

    let subtotalRow = subTotalRowNum
    for (const [cat, amt] of Object.entries(grouped)) {
      const sr = ws.getRow(subtotalRow)
      sr.height = 20
      setCellStyle(ws, subtotalRow, base, '', {
        fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.totalBg } },
        border: border(),
      })
      setCellStyle(ws, subtotalRow, base + 1, cat, subLabelStyle)
      ws.mergeCells(subtotalRow, base + 1, subtotalRow, base + 3)
      setCellStyle(ws, subtotalRow, base + 4, 'Rp.', subLabelStyle)
      setCellStyle(ws, subtotalRow, base + 5, amt, {
        ...subLabelStyle,
        numFmt: '#,##0',
        alignment: rightAlign(),
        font: { size: 8, name: FONT_NAME, bold: true, color: { argb: COLOR.accentYellow } },
      })
      subtotalRow++
    }

    // TOTAL driver — IOH teal accent
    const sr = ws.getRow(subtotalRow)
    sr.height = 22
    setCellStyle(ws, subtotalRow, base, '', {
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.grandTotalBg } },
      border: border(),
    })
    setCellStyle(ws, subtotalRow, base + 1, 'TOTAL', {
      font: { size: 9, name: FONT_NAME, bold: true, color: { argb: COLOR.accentTeal } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.grandTotalBg } },
      alignment: leftAlign(), border: border(),
    })
    ws.mergeCells(subtotalRow, base + 1, subtotalRow, base + 3)
    setCellStyle(ws, subtotalRow, base + 4, 'Rp.', {
      font: { size: 9, name: FONT_NAME, bold: true, color: { argb: COLOR.accentTeal } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.grandTotalBg } },
      alignment: centerAlign(), border: border(),
    })
    setCellStyle(ws, subtotalRow, base + 5, driverTotal, {
      font: { size: 9, name: FONT_NAME, bold: true, color: { argb: COLOR.accentYellow } },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.grandTotalBg } },
      alignment: rightAlign(), border: border(), numFmt: '#,##0',
    })
  })

  // ── Grand total (summary kiri) ─────────────────────────────────────────────
  const grandRows = [
    { label: 'TOTAL',         amount: grandTotal, color: COLOR.accentYellow },
    { label: 'CASH ADVANCE',  amount: cashAdvance || 0, color: 'FF2563EB' },
    { label: 'SISA (SELISIH)', amount: (cashAdvance || 0) - grandTotal, color: grandTotal > (cashAdvance || 0) ? 'FFDC2626' : 'FF16A34A' },
  ]

  let grandRowNum = DATA_START_ROW + totalDataRows
  for (const gr of grandRows) {
    const grRow = ws.getRow(grandRowNum)
    grRow.height = 24
    const grStyle = {
      fill: { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FF0F1929' } },
      border: border('medium' as ExcelJS.BorderStyle),
    }
    setCellStyle(ws, grandRowNum, 1, gr.label, {
      ...grStyle,
      font: { bold: true, size: 10, name: FONT_NAME, color: { argb: 'FFFFFFFF' } },
      alignment: leftAlign(),
    })
    ws.mergeCells(grandRowNum, 1, grandRowNum, 4)
    setCellStyle(ws, grandRowNum, 5, 'Rp.', {
      ...grStyle,
      font: { bold: true, size: 10, name: FONT_NAME, color: { argb: 'FFFFFFFF' } },
      alignment: centerAlign(),
    })
    setCellStyle(ws, grandRowNum, 6, gr.amount, {
      ...grStyle,
      font: { bold: true, size: 11, name: FONT_NAME, color: { argb: gr.color } },
      alignment: rightAlign(), numFmt: '#,##0',
    })
    grandRowNum++
  }

  // ── Signature ──────────────────────────────────────────────────────────────
  grandRowNum += 2
  if (reportDate) {
    setCellStyle(ws, grandRowNum, 5, reportDate, {
      font: { size: 9, name: FONT_NAME, italic: true },
      alignment: leftAlign(),
    })
    ws.mergeCells(grandRowNum, 5, grandRowNum, 6)
  }
  grandRowNum += 2
  setCellStyle(ws, grandRowNum, 1, 'Dibuat oleh', { font: { size: 9, name: FONT_NAME }, alignment: centerAlign() })
  ws.mergeCells(grandRowNum, 1, grandRowNum, 2)
  setCellStyle(ws, grandRowNum, 4, 'Mengetahui', { font: { size: 9, name: FONT_NAME }, alignment: centerAlign() })

  grandRowNum += 4
  if (createdBy) {
    setCellStyle(ws, grandRowNum, 1, createdBy, { font: { size: 9, name: FONT_NAME, bold: true }, alignment: centerAlign() })
    ws.mergeCells(grandRowNum, 1, grandRowNum, 2)
  }
  if (approvedBy) {
    setCellStyle(ws, grandRowNum, 4, approvedBy, { font: { size: 9, name: FONT_NAME, bold: true }, alignment: centerAlign() })
    ws.mergeCells(grandRowNum, 4, grandRowNum, 5)
  }
  grandRowNum++
  if (createdByTitle) {
    setCellStyle(ws, grandRowNum, 1, createdByTitle, {
      font: { size: 8, name: FONT_NAME, color: { argb: 'FF555558' } }, alignment: centerAlign(),
    })
    ws.mergeCells(grandRowNum, 1, grandRowNum, 2)
  }
  if (approvedByTitle) {
    setCellStyle(ws, grandRowNum, 4, approvedByTitle, {
      font: { size: 8, name: FONT_NAME, color: { argb: 'FF555558' } }, alignment: centerAlign(),
    })
    ws.mergeCells(grandRowNum, 4, grandRowNum, 5)
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// SHEET PER DRIVER — format vertikal
// ─────────────────────────────────────────────────────────────────────────────

async function buildSheetDriver(
  wb: ExcelJS.Workbook,
  driver: DriverData,
  title: string,
  subtitle: string,
) {
  const sheetName = driver.name.length > 28 ? driver.name.slice(0, 28) : driver.name
  const ws = wb.addWorksheet(sheetName, {
    pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1 },
  })

  ws.getColumn(1).width = 6
  ws.getColumn(2).width = 14
  ws.getColumn(3).width = 20   // wider for LAINNYA: ...
  ws.getColumn(4).width = 44
  ws.getColumn(5).width = 6
  ws.getColumn(6).width = 16

  // Judul
  const r1 = ws.addRow([])
  r1.height = 28
  setCellStyle(ws, r1.number, 1, title, {
    font: { bold: true, size: 12, name: FONT_NAME, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.headerBg } },
    alignment: centerAlign(), border: border(),
  })
  ws.mergeCells(r1.number, 1, r1.number, 6)

  const r2 = ws.addRow([])
  r2.height = 26
  setCellStyle(ws, r2.number, 1, subtitle, {
    font: { size: 9, name: FONT_NAME, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.subheaderBg } },
    alignment: centerAlign(true), border: border(),
  })
  ws.mergeCells(r2.number, 1, r2.number, 6)

  const r3 = ws.addRow([])
  r3.height = 26
  setCellStyle(ws, r3.number, 2, 'JENIS PEMAKAIAN', {
    font: { bold: true, size: 9, name: FONT_NAME },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EEF5' } },
    alignment: centerAlign(), border: border(),
  })
  ws.mergeCells(r3.number, 2, r3.number, 3)
  setCellStyle(ws, r3.number, 4, 'Data Keseluruhan', {
    font: { bold: true, size: 9, name: FONT_NAME },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD6E4F0' } },
    alignment: centerAlign(), border: border(),
  })
  ws.mergeCells(r3.number, 4, r3.number, 6)
  setCellStyle(ws, r3.number, 1, '', {
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8EEF5' } },
    border: border(),
  })

  // Info driver header
  const r4 = ws.addRow([])
  r4.height = 26
  ;[1, 2, 3].forEach(c => setCellStyle(ws, r4.number, c, '', {
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.driverBg } }, border: border(),
  }))
  setCellStyle(ws, r4.number, 4, 'NAMA DRIVER', {
    font: { bold: true, size: 9, name: FONT_NAME, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.driverBg } },
    alignment: centerAlign(), border: border(),
  })
  setCellStyle(ws, r4.number, 5, 'PLAT NOMOR', {
    font: { bold: true, size: 9, name: FONT_NAME, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.driverBg } },
    alignment: centerAlign(), border: border(),
  })
  setCellStyle(ws, r4.number, 6, 'TYPE', {
    font: { bold: true, size: 9, name: FONT_NAME, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.driverBg } },
    alignment: centerAlign(), border: border(),
  })

  const r5 = ws.addRow([])
  r5.height = 26
  ;[1, 2, 3].forEach(c => setCellStyle(ws, r5.number, c, '', {
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A3A5C' } }, border: border(),
  }))
  setCellStyle(ws, r5.number, 4, driver.name, {
    font: { bold: true, size: 9, name: FONT_NAME, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A3A5C' } },
    alignment: centerAlign(), border: border(),
  })
  setCellStyle(ws, r5.number, 5, driver.plateNumber, {
    font: { bold: true, size: 9, name: FONT_NAME, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A3A5C' } },
    alignment: centerAlign(), border: border(),
  })
  setCellStyle(ws, r5.number, 6, driver.vehicleType, {
    font: { bold: true, size: 9, name: FONT_NAME, color: { argb: 'FFFFFFFF' } },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A3A5C' } },
    alignment: centerAlign(), border: border(),
  })

  // Column headers — "TGL STRUK" menggantikan "MTA / TGL"
  const r6 = ws.addRow([])
  r6.height = 26
  const hdrs = ['NO.', 'TGL STRUK', 'BIAYA', 'DESCRIPTION', 'Rp.', 'NOMINAL']
  hdrs.forEach((h, i) => {
    setCellStyle(ws, r6.number, i + 1, h, {
      font: { bold: true, size: 9, color: { argb: 'FFFFFFFF' }, name: FONT_NAME },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2D2D30' } },
      alignment: centerAlign(), border: border(),
    })
  })

  // Data rows — ordered by bill_date (entry.date = bill_date, no = urutan berdasarkan tgl struk)
  driver.entries.forEach((entry, i) => {
    const row = ws.addRow([])
    row.height = 26
    const bg = i % 2 === 0 ? 'FFFFFFFF' : COLOR.altRow

    setCellStyle(ws, row.number, 1, entry.no, {
      font: { size: 9, name: FONT_NAME },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } },
      alignment: centerAlign(), border: border(),
    })
    // ── Kolom 2: tanggal struk bukan tanggal submit ──
    setCellStyle(ws, row.number, 2, formatExcelDate(entry.date), {
      font: { size: 9, name: FONT_NAME },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } },
      alignment: centerAlign(), border: border(),
    })
    setCellStyle(ws, row.number, 3, entry.category, {
      font: { size: 9, name: FONT_NAME },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } },
      alignment: leftAlign(true), border: border(),
    })
    setCellStyle(ws, row.number, 4, entry.description, {
      font: { size: 9, name: FONT_NAME },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } },
      alignment: leftAlign(true), border: border(),
    })
    setCellStyle(ws, row.number, 5, 'Rp.', {
      font: { size: 9, name: FONT_NAME },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } },
      alignment: centerAlign(), border: border(),
    })
    setCellStyle(ws, row.number, 6, entry.amount, {
      font: { size: 9, name: FONT_NAME },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } },
      alignment: rightAlign(), border: border(), numFmt: '#,##0',
    })
  })

  // Subtotal per kategori
  const grouped: Record<string, number> = {}
  const total = driver.entries.reduce((s, e) => {
    grouped[e.category] = (grouped[e.category] || 0) + e.amount
    return s + e.amount
  }, 0)

  ws.addRow([])

  for (const [cat, amt] of Object.entries(grouped)) {
    const sr = ws.addRow([])
    sr.height = 26
    const subStyle = {
      font: { bold: true, size: 9, name: FONT_NAME, color: { argb: COLOR.headerText } },
      fill: { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: COLOR.totalBg } },
      border: border(),
    }
    setCellStyle(ws, sr.number, 1, '', { ...subStyle })
    setCellStyle(ws, sr.number, 2, '', { ...subStyle })
    setCellStyle(ws, sr.number, 3, cat, { ...subStyle, alignment: leftAlign() })
    setCellStyle(ws, sr.number, 4, `${cat} OPERS MOBIL ${driver.plateNumber}`, { ...subStyle, alignment: leftAlign(true) })
    setCellStyle(ws, sr.number, 5, 'Rp.', { ...subStyle, alignment: centerAlign() })
    setCellStyle(ws, sr.number, 6, amt, {
      ...subStyle, alignment: rightAlign(), numFmt: '#,##0',
      font: { bold: true, size: 9, name: FONT_NAME, color: { argb: COLOR.accentYellow } },
    })
  }

  // TOTAL — IOH teal accent
  const totalRow = ws.addRow([])
  totalRow.height = 24
  const totStyle = {
    font: { bold: true, size: 10, name: FONT_NAME, color: { argb: COLOR.accentTeal } },
    fill: { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: COLOR.grandTotalBg } },
    border: border('medium' as ExcelJS.BorderStyle),
  }
  setCellStyle(ws, totalRow.number, 1, '', { ...totStyle })
  setCellStyle(ws, totalRow.number, 2, driver.name, { ...totStyle, alignment: leftAlign() })
  ws.mergeCells(totalRow.number, 2, totalRow.number, 3)
  setCellStyle(ws, totalRow.number, 4, 'TOTAL', { ...totStyle, alignment: leftAlign() })
  setCellStyle(ws, totalRow.number, 5, 'Rp.', { ...totStyle, alignment: centerAlign() })
  setCellStyle(ws, totalRow.number, 6, total, {
    ...totStyle, alignment: rightAlign(), numFmt: '#,##0',
    font: { bold: true, size: 10, name: FONT_NAME, color: { argb: COLOR.accentYellow } },
  })
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function setCellStyle(
  ws: ExcelJS.Worksheet,
  row: number,
  col: number,
  value: string | number | Date | null | undefined,
  style: {
    font?: Partial<ExcelJS.Font>
    fill?: ExcelJS.Fill
    alignment?: Partial<ExcelJS.Alignment>
    border?: Partial<ExcelJS.Borders>
    numFmt?: string
  }
) {
  const cell = ws.getCell(row, col)
  if (value !== undefined && value !== null) cell.value = value
  if (style.font) cell.font = style.font
  if (style.fill) cell.fill = style.fill
  if (style.alignment) cell.alignment = style.alignment
  if (style.border) cell.border = style.border
  if (style.numFmt) cell.numFmt = style.numFmt
}

/**
 * Normalize category untuk Excel.
 * - parkir / tol / bensin → uppercase standard
 * - lainnya: gunakan description sebagai suffix (misal "LAINNYA: TAMBAL BAN")
 */
export function normalizeCategory(cat: string, description?: string): string {
  const lower = cat.toLowerCase().trim()
  if (lower.includes('bbm') || lower.includes('bensin')) return 'BIAYA BBM'
  if (lower.includes('toll') || lower.includes('tol')) return 'BIAYA TOLL'
  if (lower.includes('parkir') || lower.includes('parking')) return 'BIAYA PARKIR'
  if (lower === 'lainnya') {
    return description ? `LAINNYA: ${description.toUpperCase()}` : 'LAINNYA'
  }
  return cat.toUpperCase()
}

/**
 * Format tanggal untuk Excel — menerima ISO date string.
 * Hasilnya: DD/MM/YYYY (format Indonesia standar).
 */
function formatExcelDate(dateStr: string): string {
  if (!dateStr) return ''
  try {
    const d = new Date(dateStr)
    return d.toLocaleDateString('id-ID', { day: '2-digit', month: '2-digit', year: 'numeric' })
  } catch {
    return dateStr
  }
}

/**
 * Helper untuk caller: sort entries berdasarkan bill_date ascending
 * dan assign nomor urut yang konsisten dengan PDF.
 *
 * Panggil ini sebelum membuat BillEntry[] untuk setiap driver:
 *
 *   const entries = prepareBillEntries(rawSubmissions, driverPlate)
 */
export function prepareBillEntries(
  submissions: Array<{
    bill_date?: string | null
    submission_date: string
    category: string
    description?: string | null
    amount?: number | null
    image_url?: string
  }>,
  platNumber: string,
): BillEntry[] {
  // Sort by bill_date, fallback submission_date
  const sorted = [...submissions].sort((a, b) => {
    const da = new Date(a.bill_date || a.submission_date).getTime()
    const db = new Date(b.bill_date || b.submission_date).getTime()
    return da - db
  })

  return sorted.map((sub, i) => ({
    no: i + 1,
    date: sub.bill_date || sub.submission_date,   // ← selalu tanggal struk
    category: normalizeCategory(sub.category, sub.description || undefined),
    description: buildDescription(sub.category, sub.description, platNumber),
    amount: sub.amount || 0,
  }))
}

function buildDescription(category: string, description?: string | null, plate?: string): string {
  const cat = normalizeCategory(category, description || undefined)
  if (category === 'lainnya' && description) {
    return `${cat} OPERS MOBIL${plate ? ' ' + plate : ''}`
  }
  return `${cat} OPERS MOBIL${plate ? ' ' + plate : ''}`
}