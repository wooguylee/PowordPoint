const clampNumber = (value, fallback, min = -Infinity, max = Infinity) => {
  const numeric = Number(value)

  if (!Number.isFinite(numeric)) {
    return fallback
  }

  return Math.min(max, Math.max(min, numeric))
}

export const PAGE_WIDTH = 960
export const PAGE_HEIGHT = 1280
export const LOCAL_DRAFT_KEY = 'powordpointer.autosave'
export const LOCAL_LLM_KEY = 'powordpointer.llm'

export const TOOL_OPTIONS = [
  { id: 'select', label: 'Select' },
  { id: 'text', label: 'Text' },
  { id: 'rect', label: 'Rect' },
  { id: 'ellipse', label: 'Circle' },
  { id: 'arrow', label: 'Arrow' },
  { id: 'pen', label: 'Pen' },
  { id: 'table', label: 'Table' },
]

export const LLM_ACTIONS = [
  { id: 'draft', label: 'Draft writing' },
  { id: 'layout', label: 'Shape layout' },
  { id: 'pages', label: 'Page generation' },
  { id: 'selection', label: 'Edit selection' },
]

export const TRANSFORMABLE_TYPES = new Set([
  'text',
  'rect',
  'ellipse',
  'arrow',
  'image',
  'table',
])

const DEFAULT_FONT = 'Avenir Next, Segoe UI Variable, Noto Sans KR, sans-serif'
const DEFAULT_HEADING = 'Space Grotesk, Avenir Next, sans-serif'

export const createId = (prefix = 'item') =>
  `${prefix}-${Math.random().toString(36).slice(2, 10)}`

export const createTextElement = (x = 120, y = 120, overrides = {}) => ({
  id: createId('text'),
  type: 'text',
  x,
  y,
  width: 360,
  height: 120,
  rotation: 0,
  text: 'New text block',
  fontSize: 28,
  fontFamily: DEFAULT_HEADING,
  fill: '#142b36',
  align: 'left',
  lineHeight: 1.18,
  fontStyle: 'normal',
  padding: 8,
  ...overrides,
})

export const createShapeElement = (type, x = 120, y = 120, overrides = {}) => ({
  id: createId(type),
  type,
  x,
  y,
  width: 220,
  height: 140,
  rotation: 0,
  fill: type === 'rect' ? '#c7e7ef' : '#ffe1c6',
  stroke: '#23404d',
  strokeWidth: 2,
  opacity: 1,
  ...overrides,
})

export const createArrowElement = (x = 120, y = 120, overrides = {}) => ({
  id: createId('arrow'),
  type: 'arrow',
  x,
  y,
  width: 240,
  height: 0,
  rotation: 0,
  stroke: '#d56f3e',
  strokeWidth: 4,
  pointerLength: 18,
  pointerWidth: 18,
  opacity: 1,
  ...overrides,
})

export const createPenElement = (x = 120, y = 120, overrides = {}) => ({
  id: createId('pen'),
  type: 'pen',
  x,
  y,
  points: [0, 0],
  stroke: '#142b36',
  strokeWidth: 3,
  opacity: 1,
  ...overrides,
})

export const createTableElement = (x = 120, y = 120, overrides = {}) => ({
  id: createId('table'),
  type: 'table',
  x,
  y,
  width: 420,
  height: 240,
  rotation: 0,
  rows: 4,
  cols: 3,
  stroke: '#23404d',
  strokeWidth: 1,
  fill: '#fffdf8',
  headerFill: '#ddecef',
  textColor: '#17323d',
  fontSize: 18,
  cells: [
    ['Header 1', 'Header 2', 'Header 3'],
    ['Value', 'Value', 'Value'],
    ['Value', 'Value', 'Value'],
    ['Value', 'Value', 'Value'],
  ],
  ...overrides,
})

export const createImageElement = (src, x = 120, y = 120, overrides = {}) => ({
  id: createId('image'),
  type: 'image',
  x,
  y,
  width: 320,
  height: 220,
  rotation: 0,
  opacity: 1,
  src,
  ...overrides,
})

export const createBlankPage = (name = 'Untitled page') => ({
  id: createId('page'),
  name,
  width: PAGE_WIDTH,
  height: PAGE_HEIGHT,
  background: '#fffaf2',
  elements: [
    createTextElement(88, 86, {
      text: 'PowordPointer canvas document',
      fontSize: 40,
      width: 640,
      height: 80,
    }),
    createTextElement(92, 182, {
      text: 'Mix text, shapes, diagrams, and AI-generated layouts in one document surface.',
      fontSize: 22,
      width: 660,
      height: 90,
      fontFamily: DEFAULT_FONT,
      fill: '#48606b',
      lineHeight: 1.35,
    }),
    createShapeElement('rect', 88, 326, {
      width: 320,
      height: 180,
      fill: '#d8eef2',
      stroke: '#275063',
    }),
    createArrowElement(430, 414, {
      width: 148,
      height: 0,
    }),
    createShapeElement('ellipse', 598, 326, {
      width: 228,
      height: 180,
      fill: '#ffd8bd',
      stroke: '#b95d32',
    }),
    createTableElement(88, 566, {
      width: 740,
      height: 236,
    }),
  ],
})

export const createDefaultDocument = () => ({
  id: createId('doc'),
  version: 1,
  title: 'PowordPointer Document',
  description: 'Canvas document with AI-assisted writing and layout',
  createdAt: new Date().toISOString(),
  updatedAt: new Date().toISOString(),
  pages: [createBlankPage('Cover')],
})

export const normalizeBox = (start, end, minWidth = 48, minHeight = 48) => {
  const width = Math.max(minWidth, Math.abs(end.x - start.x))
  const height = Math.max(minHeight, Math.abs(end.y - start.y))

  return {
    x: Math.min(start.x, end.x),
    y: Math.min(start.y, end.y),
    width,
    height,
  }
}

export const createDrawElement = (tool, point) => {
  if (tool === 'rect') {
    return createShapeElement('rect', point.x, point.y, { width: 1, height: 1 })
  }

  if (tool === 'ellipse') {
    return createShapeElement('ellipse', point.x, point.y, { width: 1, height: 1 })
  }

  if (tool === 'arrow') {
    return createArrowElement(point.x, point.y, { width: 1, height: 1 })
  }

  if (tool === 'pen') {
    return createPenElement(point.x, point.y)
  }

  return null
}

export const finalizeDrawElement = (element, start, end) => {
  if (!element) {
    return element
  }

  if (element.type === 'arrow') {
    return {
      ...element,
      x: start.x,
      y: start.y,
      width: Math.abs(end.x - start.x) < 18 ? 220 : end.x - start.x,
      height: Math.abs(end.y - start.y) < 18 ? 0 : end.y - start.y,
    }
  }

  if (element.type === 'pen') {
    return element
  }

  const box = normalizeBox(start, end)

  return {
    ...element,
    ...box,
  }
}

const coerceCells = (cells, rows, cols) => {
  const safeRows = clampNumber(rows, 3, 1, 20)
  const safeCols = clampNumber(cols, 3, 1, 10)

  return Array.from({ length: safeRows }, (_, rowIndex) =>
    Array.from({ length: safeCols }, (_, colIndex) => {
      const nextValue = cells?.[rowIndex]?.[colIndex]
      return typeof nextValue === 'string'
        ? nextValue
        : rowIndex === 0
          ? `Header ${colIndex + 1}`
          : 'Value'
    }),
  )
}

export const sanitizeElement = (rawElement, fallbackPosition = { x: 120, y: 120 }) => {
  if (!rawElement || typeof rawElement !== 'object') {
    return null
  }

  const type = rawElement.type

  if (type === 'text') {
    return createTextElement(fallbackPosition.x, fallbackPosition.y, {
      id: rawElement.id || createId('text'),
      x: clampNumber(rawElement.x, fallbackPosition.x, 0, PAGE_WIDTH),
      y: clampNumber(rawElement.y, fallbackPosition.y, 0, PAGE_HEIGHT),
      width: clampNumber(rawElement.width, 360, 80, PAGE_WIDTH),
      height: clampNumber(rawElement.height, 120, 32, PAGE_HEIGHT),
      rotation: clampNumber(rawElement.rotation, 0, -360, 360),
      text: typeof rawElement.text === 'string' ? rawElement.text : 'Generated text',
      fontSize: clampNumber(rawElement.fontSize, 26, 10, 120),
      fontFamily:
        typeof rawElement.fontFamily === 'string' && rawElement.fontFamily.trim()
          ? rawElement.fontFamily
          : DEFAULT_FONT,
      fill: typeof rawElement.fill === 'string' ? rawElement.fill : '#142b36',
      align: ['left', 'center', 'right'].includes(rawElement.align)
        ? rawElement.align
        : 'left',
      lineHeight: clampNumber(rawElement.lineHeight, 1.25, 0.8, 2),
      fontStyle: ['normal', 'bold', 'italic'].includes(rawElement.fontStyle)
        ? rawElement.fontStyle
        : 'normal',
      padding: clampNumber(rawElement.padding, 8, 0, 40),
      groupId: typeof rawElement.groupId === 'string' ? rawElement.groupId : null,
    })
  }

  if (type === 'rect' || type === 'ellipse') {
    return createShapeElement(type, fallbackPosition.x, fallbackPosition.y, {
      id: rawElement.id || createId(type),
      x: clampNumber(rawElement.x, fallbackPosition.x, 0, PAGE_WIDTH),
      y: clampNumber(rawElement.y, fallbackPosition.y, 0, PAGE_HEIGHT),
      width: clampNumber(rawElement.width, 220, 20, PAGE_WIDTH),
      height: clampNumber(rawElement.height, 140, 20, PAGE_HEIGHT),
      rotation: clampNumber(rawElement.rotation, 0, -360, 360),
      fill: typeof rawElement.fill === 'string' ? rawElement.fill : '#c7e7ef',
      stroke: typeof rawElement.stroke === 'string' ? rawElement.stroke : '#23404d',
      strokeWidth: clampNumber(rawElement.strokeWidth, 2, 0, 20),
      opacity: clampNumber(rawElement.opacity, 1, 0.1, 1),
      groupId: typeof rawElement.groupId === 'string' ? rawElement.groupId : null,
    })
  }

  if (type === 'arrow') {
    return createArrowElement(fallbackPosition.x, fallbackPosition.y, {
      id: rawElement.id || createId('arrow'),
      x: clampNumber(rawElement.x, fallbackPosition.x, 0, PAGE_WIDTH),
      y: clampNumber(rawElement.y, fallbackPosition.y, 0, PAGE_HEIGHT),
      width: clampNumber(rawElement.width, 220, -PAGE_WIDTH, PAGE_WIDTH),
      height: clampNumber(rawElement.height, 0, -PAGE_HEIGHT, PAGE_HEIGHT),
      rotation: clampNumber(rawElement.rotation, 0, -360, 360),
      stroke: typeof rawElement.stroke === 'string' ? rawElement.stroke : '#d56f3e',
      strokeWidth: clampNumber(rawElement.strokeWidth, 4, 1, 20),
      pointerLength: clampNumber(rawElement.pointerLength, 18, 6, 64),
      pointerWidth: clampNumber(rawElement.pointerWidth, 18, 6, 64),
      opacity: clampNumber(rawElement.opacity, 1, 0.1, 1),
      groupId: typeof rawElement.groupId === 'string' ? rawElement.groupId : null,
    })
  }

  if (type === 'pen') {
    const points = Array.isArray(rawElement.points)
      ? rawElement.points.filter((value) => Number.isFinite(Number(value))).map(Number)
      : [0, 0]

    return createPenElement(fallbackPosition.x, fallbackPosition.y, {
      id: rawElement.id || createId('pen'),
      x: clampNumber(rawElement.x, fallbackPosition.x, 0, PAGE_WIDTH),
      y: clampNumber(rawElement.y, fallbackPosition.y, 0, PAGE_HEIGHT),
      points: points.length >= 4 ? points : [0, 0, 120, 10],
      stroke: typeof rawElement.stroke === 'string' ? rawElement.stroke : '#142b36',
      strokeWidth: clampNumber(rawElement.strokeWidth, 3, 1, 24),
      opacity: clampNumber(rawElement.opacity, 1, 0.1, 1),
      groupId: typeof rawElement.groupId === 'string' ? rawElement.groupId : null,
    })
  }

  if (type === 'table') {
    const rows = clampNumber(rawElement.rows, 4, 1, 20)
    const cols = clampNumber(rawElement.cols, 3, 1, 10)

    return createTableElement(fallbackPosition.x, fallbackPosition.y, {
      id: rawElement.id || createId('table'),
      x: clampNumber(rawElement.x, fallbackPosition.x, 0, PAGE_WIDTH),
      y: clampNumber(rawElement.y, fallbackPosition.y, 0, PAGE_HEIGHT),
      width: clampNumber(rawElement.width, 420, 120, PAGE_WIDTH),
      height: clampNumber(rawElement.height, 240, 80, PAGE_HEIGHT),
      rotation: clampNumber(rawElement.rotation, 0, -360, 360),
      rows,
      cols,
      stroke: typeof rawElement.stroke === 'string' ? rawElement.stroke : '#23404d',
      strokeWidth: clampNumber(rawElement.strokeWidth, 1, 1, 12),
      fill: typeof rawElement.fill === 'string' ? rawElement.fill : '#fffdf8',
      headerFill: typeof rawElement.headerFill === 'string' ? rawElement.headerFill : '#ddecef',
      textColor: typeof rawElement.textColor === 'string' ? rawElement.textColor : '#17323d',
      fontSize: clampNumber(rawElement.fontSize, 18, 10, 48),
      cells: coerceCells(rawElement.cells, rows, cols),
      groupId: typeof rawElement.groupId === 'string' ? rawElement.groupId : null,
    })
  }

  if (type === 'image' && typeof rawElement.src === 'string' && rawElement.src.trim()) {
    return createImageElement(rawElement.src, fallbackPosition.x, fallbackPosition.y, {
      id: rawElement.id || createId('image'),
      x: clampNumber(rawElement.x, fallbackPosition.x, 0, PAGE_WIDTH),
      y: clampNumber(rawElement.y, fallbackPosition.y, 0, PAGE_HEIGHT),
      width: clampNumber(rawElement.width, 320, 40, PAGE_WIDTH),
      height: clampNumber(rawElement.height, 220, 40, PAGE_HEIGHT),
      rotation: clampNumber(rawElement.rotation, 0, -360, 360),
      opacity: clampNumber(rawElement.opacity, 1, 0.1, 1),
      groupId: typeof rawElement.groupId === 'string' ? rawElement.groupId : null,
    })
  }

  return null
}

export const sanitizePage = (rawPage, index = 0) => {
  if (!rawPage || typeof rawPage !== 'object') {
    return createBlankPage(`Page ${index + 1}`)
  }

  const width = clampNumber(rawPage.width, PAGE_WIDTH, 480, 1600)
  const height = clampNumber(rawPage.height, PAGE_HEIGHT, 640, 2000)
  const elements = Array.isArray(rawPage.elements)
    ? rawPage.elements
        .map((element, elementIndex) =>
          sanitizeElement(element, {
            x: 80 + (elementIndex % 3) * 40,
            y: 80 + (elementIndex % 4) * 48,
          }),
        )
        .filter(Boolean)
    : []

  return {
    id: rawPage.id || createId('page'),
    name:
      typeof rawPage.name === 'string' && rawPage.name.trim()
        ? rawPage.name
        : `Page ${index + 1}`,
    width,
    height,
    background: typeof rawPage.background === 'string' ? rawPage.background : '#fffaf2',
    elements,
  }
}

export const sanitizeDocument = (rawDocument) => {
  if (!rawDocument || typeof rawDocument !== 'object') {
    return createDefaultDocument()
  }

  const pages = Array.isArray(rawDocument.pages)
    ? rawDocument.pages.map((page, index) => sanitizePage(page, index))
    : []

  return {
    id:
      typeof rawDocument.id === 'string' && rawDocument.id.trim()
        ? rawDocument.id
        : createId('doc'),
    version: 1,
    title:
      typeof rawDocument.title === 'string' && rawDocument.title.trim()
        ? rawDocument.title
        : 'PowordPointer Document',
    description:
      typeof rawDocument.description === 'string'
        ? rawDocument.description
        : 'Canvas document with AI-assisted writing and layout',
    createdAt:
      typeof rawDocument.createdAt === 'string'
        ? rawDocument.createdAt
        : new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    pages: pages.length > 0 ? pages : [createBlankPage('Cover')],
  }
}

export const duplicateElement = (element) => {
  const cloned = JSON.parse(JSON.stringify(element))
  cloned.id = createId(element.type)
  cloned.x += 28
  cloned.y += 28
  return cloned
}

export const serializeCells = (cells) =>
  (Array.isArray(cells) ? cells : []).map((row) => row.join(' | ')).join('\n')

export const parseCells = (text, rows, cols) => {
  const rawRows = `${text || ''}`.split(/\r?\n/)
  return Array.from({ length: rows }, (_, rowIndex) => {
    const rawCols = rawRows[rowIndex] ? rawRows[rowIndex].split('|') : []
    return Array.from({ length: cols }, (_, colIndex) => rawCols[colIndex]?.trim() || '')
  })
}

export const updateTimestamp = (documentData) => ({
  ...documentData,
  updatedAt: new Date().toISOString(),
})

export const getElementBounds = (element) => {
  if (!element) {
    return null
  }

  if (element.type === 'pen') {
    const xs = []
    const ys = []

    for (let index = 0; index < element.points.length; index += 2) {
      xs.push(element.x + element.points[index])
      ys.push(element.y + element.points[index + 1])
    }

    return {
      x: Math.min(...xs),
      y: Math.min(...ys),
      width: Math.max(...xs) - Math.min(...xs),
      height: Math.max(...ys) - Math.min(...ys),
    }
  }

  return {
    x: element.x,
    y: element.y,
    width: element.width || 0,
    height: element.height || 0,
  }
}

export const getSelectionBounds = (elements) => {
  if (!Array.isArray(elements) || elements.length === 0) {
    return null
  }

  const bounds = elements.map(getElementBounds).filter(Boolean)
  const left = Math.min(...bounds.map((item) => item.x))
  const top = Math.min(...bounds.map((item) => item.y))
  const right = Math.max(...bounds.map((item) => item.x + item.width))
  const bottom = Math.max(...bounds.map((item) => item.y + item.height))

  return {
    x: left,
    y: top,
    width: right - left,
    height: bottom - top,
    centerX: left + (right - left) / 2,
    centerY: top + (bottom - top) / 2,
  }
}

export const intersectsBox = (left, right) => {
  return !(
    left.x > right.x + right.width ||
    left.x + left.width < right.x ||
    left.y > right.y + right.height ||
    left.y + left.height < right.y
  )
}

export const alignElements = (elements, selectedIds, mode) => {
  const selected = elements.filter((element) => selectedIds.includes(element.id))
  const bounds = getSelectionBounds(selected)

  if (!bounds) {
    return elements
  }

  return elements.map((element) => {
    if (!selectedIds.includes(element.id)) {
      return element
    }

    const box = getElementBounds(element)

    if (!box) {
      return element
    }

    if (mode === 'left') {
      return { ...element, x: bounds.x }
    }

    if (mode === 'center') {
      return { ...element, x: bounds.centerX - box.width / 2 }
    }

    if (mode === 'right') {
      return { ...element, x: bounds.x + bounds.width - box.width }
    }

    if (mode === 'top') {
      return { ...element, y: bounds.y }
    }

    if (mode === 'middle') {
      return { ...element, y: bounds.centerY - box.height / 2 }
    }

    if (mode === 'bottom') {
      return { ...element, y: bounds.y + bounds.height - box.height }
    }

    return element
  })
}

export const distributeElements = (elements, selectedIds, direction) => {
  const selected = elements.filter((element) => selectedIds.includes(element.id))

  if (selected.length < 3) {
    return elements
  }

  const items = selected
    .map((element) => ({ element, bounds: getElementBounds(element) }))
    .filter((item) => item.bounds)
    .sort((left, right) =>
      direction === 'horizontal'
        ? left.bounds.x - right.bounds.x
        : left.bounds.y - right.bounds.y,
    )

  const first = items[0]
  const last = items[items.length - 1]

  if (!first || !last) {
    return elements
  }

  const totalSpan =
    direction === 'horizontal'
      ? last.bounds.x - first.bounds.x
      : last.bounds.y - first.bounds.y
  const gap = totalSpan / (items.length - 1)

  const nextPositions = new Map()
  items.forEach((item, index) => {
    if (index === 0 || index === items.length - 1) {
      nextPositions.set(item.element.id, item.element)
      return
    }

    nextPositions.set(item.element.id, {
      ...item.element,
      x: direction === 'horizontal' ? first.bounds.x + gap * index : item.element.x,
      y: direction === 'vertical' ? first.bounds.y + gap * index : item.element.y,
    })
  })

  return elements.map((element) => nextPositions.get(element.id) || element)
}
