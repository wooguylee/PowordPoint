import { jsPDF } from 'jspdf'

const pxToMm = (value) => value * 0.264583
let notoFontLoaded = false

const imageCache = new Map()

const ensurePdfFonts = async (pdf) => {
  if (notoFontLoaded) {
    return
  }

  const response = await fetch('/NotoSansKR.ttf')
  const buffer = await response.arrayBuffer()
  const bytes = new Uint8Array(buffer)
  let binary = ''

  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte)
  })

  const base64 = btoa(binary)
  pdf.addFileToVFS('NotoSansKR.ttf', base64)
  pdf.addFont('NotoSansKR.ttf', 'NotoSansKR', 'normal')
  pdf.addFont('NotoSansKR.ttf', 'NotoSansKR', 'bold')
  notoFontLoaded = true
}

const loadImageSource = async (src) => {
  if (!src) {
    return null
  }

  if (src.startsWith('data:')) {
    return src
  }

  if (imageCache.has(src)) {
    return imageCache.get(src)
  }

  const response = await fetch(src)
  const blob = await response.blob()
  const dataUrl = await new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(`${reader.result || ''}`)
    reader.onerror = () => reject(new Error('Failed to read image for PDF export.'))
    reader.readAsDataURL(blob)
  })

  imageCache.set(src, dataUrl)
  return dataUrl
}

const hexToRgb = (value) => {
  const input = `${value || ''}`.trim()

  if (/^#([a-f0-9]{3}|[a-f0-9]{6})$/i.test(input)) {
    const normalized = input.length === 4
      ? `#${input[1]}${input[1]}${input[2]}${input[2]}${input[3]}${input[3]}`
      : input

    return {
      r: Number.parseInt(normalized.slice(1, 3), 16),
      g: Number.parseInt(normalized.slice(3, 5), 16),
      b: Number.parseInt(normalized.slice(5, 7), 16),
    }
  }

  const rgbMatch = input.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/i)

  if (rgbMatch) {
    return {
      r: Number(rgbMatch[1]),
      g: Number(rgbMatch[2]),
      b: Number(rgbMatch[3]),
    }
  }

  return { r: 0, g: 0, b: 0 }
}

const applyDrawColor = (pdf, color) => {
  const { r, g, b } = hexToRgb(color)
  pdf.setDrawColor(r, g, b)
}

const applyFillColor = (pdf, color) => {
  const { r, g, b } = hexToRgb(color)
  pdf.setFillColor(r, g, b)
}

const applyTextColor = (pdf, color) => {
  const { r, g, b } = hexToRgb(color)
  pdf.setTextColor(r, g, b)
}

const getPdfFont = (fontStyle) => {
  if (fontStyle === 'bold') {
    return 'bold'
  }

  if (fontStyle === 'italic') {
    return 'italic'
  }

  return 'normal'
}

const drawTextBlock = (pdf, element) => {
  const x = pxToMm(element.x + element.padding)
  const y = pxToMm(element.y + element.padding + element.fontSize * 0.85)
  const maxWidth = pxToMm(Math.max(40, element.width - element.padding * 2))
  const lineHeight = pxToMm(element.fontSize * (element.lineHeight || 1.2))
  const fontSize = Math.max(8, element.fontSize * 0.72)
  const lines = pdf.splitTextToSize(`${element.text || ''}`, maxWidth)
  const align = element.align || 'left'

  applyTextColor(pdf, element.fill)
  pdf.setFont('NotoSansKR', getPdfFont(element.fontStyle))
  pdf.setFontSize(fontSize)

  lines.forEach((line, index) => {
    const lineWidth = pdf.getTextWidth(line)
    const lineX =
      align === 'center'
        ? x + (maxWidth - lineWidth) / 2
        : align === 'right'
          ? x + maxWidth - lineWidth
          : x

    pdf.text(line, lineX, y + lineHeight * index)
  })
}

const withRotation = (pdf, element, draw) => {
  const rotation = element.rotation || 0

  if (!rotation) {
    draw()
    return
  }

  pdf.saveGraphicsState()
  pdf.setCurrentTransformationMatrix(
    pdf.Matrix(
      Math.cos((rotation * Math.PI) / 180),
      Math.sin((rotation * Math.PI) / 180),
      -Math.sin((rotation * Math.PI) / 180),
      Math.cos((rotation * Math.PI) / 180),
      pxToMm(element.x + (element.width || 0) / 2),
      pxToMm(element.y + (element.height || 0) / 2),
    ),
  )
  draw(-(element.width || 0) / 2, -(element.height || 0) / 2)
  pdf.restoreGraphicsState()
}

const addArrowHead = (pdf, fromX, fromY, toX, toY, size) => {
  const angle = Math.atan2(toY - fromY, toX - fromX)
  const leftX = toX - size * Math.cos(angle - Math.PI / 6)
  const leftY = toY - size * Math.sin(angle - Math.PI / 6)
  const rightX = toX - size * Math.cos(angle + Math.PI / 6)
  const rightY = toY - size * Math.sin(angle + Math.PI / 6)

  pdf.triangle(toX, toY, leftX, leftY, rightX, rightY, 'F')
}

const drawTable = (pdf, element) => {
  const cellWidth = element.width / element.cols
  const cellHeight = element.height / element.rows

  applyDrawColor(pdf, element.stroke)
  pdf.setLineWidth(Math.max(0.2, pxToMm(element.strokeWidth)))
  applyFillColor(pdf, element.fill)
  pdf.roundedRect(pxToMm(element.x), pxToMm(element.y), pxToMm(element.width), pxToMm(element.height), 2.4, 2.4, 'FD')
  applyFillColor(pdf, element.headerFill)
  pdf.rect(pxToMm(element.x), pxToMm(element.y), pxToMm(element.width), pxToMm(cellHeight), 'F')

  for (let column = 1; column < element.cols; column += 1) {
    const x = element.x + cellWidth * column
    pdf.line(pxToMm(x), pxToMm(element.y), pxToMm(x), pxToMm(element.y + element.height))
  }

  for (let row = 1; row < element.rows; row += 1) {
    const y = element.y + cellHeight * row
    pdf.line(pxToMm(element.x), pxToMm(y), pxToMm(element.x + element.width), pxToMm(y))
  }

  applyTextColor(pdf, element.textColor)

  element.cells.forEach((row, rowIndex) => {
    row.forEach((cell, columnIndex) => {
      pdf.setFont('NotoSansKR', rowIndex === 0 ? 'bold' : 'normal')
      pdf.setFontSize(Math.max(8, element.fontSize * 0.72))
      pdf.text(
        `${cell || ''}`,
        pxToMm(element.x + columnIndex * cellWidth + 10),
        pxToMm(element.y + rowIndex * cellHeight + 22),
        { maxWidth: pxToMm(cellWidth - 16) },
      )
    })
  })
}

const drawPen = (pdf, element) => {
  if (!Array.isArray(element.points) || element.points.length < 4) {
    return
  }

  applyDrawColor(pdf, element.stroke)
  pdf.setLineWidth(Math.max(0.3, pxToMm(element.strokeWidth)))

  for (let index = 2; index < element.points.length; index += 2) {
    const fromX = element.x + element.points[index - 2]
    const fromY = element.y + element.points[index - 1]
    const toX = element.x + element.points[index]
    const toY = element.y + element.points[index + 1]

    pdf.line(pxToMm(fromX), pxToMm(fromY), pxToMm(toX), pxToMm(toY))
  }
}

const drawElement = async (pdf, element) => {
  if (element.type === 'text') {
    withRotation(pdf, element, (offsetX = 0, offsetY = 0) => {
      drawTextBlock(pdf, {
        ...element,
        x: (element.x || 0) + offsetX,
        y: (element.y || 0) + offsetY,
      })
    })
    return
  }

  if (element.type === 'rect') {
    withRotation(pdf, element, (offsetX = 0, offsetY = 0) => {
      applyDrawColor(pdf, element.stroke)
      applyFillColor(pdf, element.fill)
      pdf.setLineWidth(Math.max(0.2, pxToMm(element.strokeWidth)))
      pdf.roundedRect(pxToMm(element.x + offsetX), pxToMm(element.y + offsetY), pxToMm(element.width), pxToMm(element.height), 3.2, 3.2, 'FD')
    })
    return
  }

  if (element.type === 'ellipse') {
    withRotation(pdf, element, (offsetX = 0, offsetY = 0) => {
      applyDrawColor(pdf, element.stroke)
      applyFillColor(pdf, element.fill)
      pdf.setLineWidth(Math.max(0.2, pxToMm(element.strokeWidth)))
      pdf.ellipse(pxToMm(element.x + offsetX + element.width / 2), pxToMm(element.y + offsetY + element.height / 2), pxToMm(element.width / 2), pxToMm(element.height / 2), 'FD')
    })
    return
  }

  if (element.type === 'arrow') {
    withRotation(pdf, element, (offsetX = 0, offsetY = 0) => {
      const fromX = pxToMm(element.x + offsetX)
      const fromY = pxToMm(element.y + offsetY)
      const toX = pxToMm(element.x + offsetX + element.width)
      const toY = pxToMm(element.y + offsetY + element.height)
      const size = Math.max(2, pxToMm(element.pointerLength || 18))
      applyDrawColor(pdf, element.stroke)
      applyFillColor(pdf, element.stroke)
      pdf.setLineWidth(Math.max(0.2, pxToMm(element.strokeWidth)))
      pdf.line(fromX, fromY, toX, toY)
      addArrowHead(pdf, fromX, fromY, toX, toY, size)
    })
    return
  }

  if (element.type === 'pen') {
    drawPen(pdf, element)
    return
  }

  if (element.type === 'table') {
    withRotation(pdf, element, (offsetX = 0, offsetY = 0) => {
      drawTable(pdf, {
        ...element,
        x: (element.x || 0) + offsetX,
        y: (element.y || 0) + offsetY,
      })
    })
    return
  }

  if (element.type === 'image' && element.src) {
    withRotation(pdf, element, (offsetX = 0, offsetY = 0) => {
      pdf.addImage(element.__pdfSrc, 'PNG', pxToMm(element.x + offsetX), pxToMm(element.y + offsetY), pxToMm(element.width), pxToMm(element.height))
    })
  }
}

export const exportDocumentToPdf = async (documentData) => {
  const firstPage = documentData.pages[0]
  const pdf = new jsPDF({
    orientation: firstPage.width > firstPage.height ? 'landscape' : 'portrait',
    unit: 'mm',
    format: [pxToMm(firstPage.width), pxToMm(firstPage.height)],
    compress: true,
  })

  await ensurePdfFonts(pdf)

  for (let pageIndex = 0; pageIndex < documentData.pages.length; pageIndex += 1) {
    const page = documentData.pages[pageIndex]

    if (pageIndex > 0) {
      pdf.addPage([pxToMm(page.width), pxToMm(page.height)], page.width > page.height ? 'landscape' : 'portrait')
    }

    applyFillColor(pdf, page.background)
    pdf.rect(0, 0, pxToMm(page.width), pxToMm(page.height), 'F')

    for (const element of page.elements) {
      const nextElement =
        element.type === 'image' && element.src
          ? { ...element, __pdfSrc: await loadImageSource(element.src) }
          : element
      await drawElement(pdf, nextElement)
    }
  }

  pdf.save(`${documentData.title || 'powordpointer-document'}.pdf`)
}
