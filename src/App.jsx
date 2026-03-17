import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import {
  Arrow,
  Ellipse,
  Group,
  Image,
  Layer,
  Line,
  Rect,
  Stage,
  Text,
  Transformer,
} from 'react-konva'
import './App.css'
import {
  LLM_ACTIONS,
  LOCAL_DRAFT_KEY,
  LOCAL_LLM_KEY,
  TOOL_OPTIONS,
  TRANSFORMABLE_TYPES,
  alignElements,
  createArrowElement,
  createBlankPage,
  createDefaultDocument,
  createDrawElement,
  createId,
  createImageElement,
  createShapeElement,
  createTableElement,
  createTextElement,
  duplicateElement,
  finalizeDrawElement,
  getElementBounds,
  getSelectionBounds,
  intersectsBox,
  normalizeBox,
  parseCells,
  sanitizeDocument,
  sanitizeElement,
  serializeCells,
  updateTimestamp,
} from './lib/editor'
import {
  cleanupUploadsOnServer,
  deleteDocumentFromServer,
  deleteUploadFromServer,
  fetchDocumentVersions,
  fetchDocumentById,
  fetchDocumentLibrary,
  fetchUploads,
  renameDocumentOnServer,
  requestLlmProxy,
  restoreDocumentVersion as restoreDocumentVersionApi,
  saveDocumentToServer,
  uploadImageToServer,
} from './lib/api'
import { normalizeLlmPayload, buildPrompts } from './lib/llm'
import { exportDocumentToPdf } from './lib/pdf'

const defaultLlmState = {
  apiKey: '',
  baseUrl: 'https://api.openai.com/v1',
  model: 'gpt-4.1-mini',
}

const snapTolerance = 8
const gridSize = 24

const toSlug = (value) =>
  `${value || 'document'}`
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/(^-|-$)/g, '') || 'document'

const readStoredDocument = () => {
  const raw = window.localStorage.getItem(LOCAL_DRAFT_KEY)

  if (!raw) {
    return createDefaultDocument()
  }

  try {
    return sanitizeDocument(JSON.parse(raw))
  } catch {
    return createDefaultDocument()
  }
}

const readStoredLlmConfig = () => {
  const raw = window.localStorage.getItem(LOCAL_LLM_KEY)

  if (!raw) {
    return defaultLlmState
  }

  try {
    return { ...defaultLlmState, ...JSON.parse(raw) }
  } catch {
    return defaultLlmState
  }
}

const downloadBlob = (blob, name) => {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = name
  anchor.click()
  URL.revokeObjectURL(url)
}

const readFileAsText = (file) =>
  new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(`${reader.result || ''}`)
    reader.onerror = () => reject(new Error('Failed to read file.'))
    reader.readAsText(file)
  })

const getPointerPosition = (stage, scale) => {
  const pointer = stage.getPointerPosition()

  if (!pointer) {
    return null
  }

  return {
    x: pointer.x / scale,
    y: pointer.y / scale,
  }
}

const getTargetIsCanvas = (target, stage) => target === stage || target.name() === 'page-background'

const extractElementPatchFromNode = (shapeNode, element) => {
  const patch = {
    x: shapeNode.x(),
    y: shapeNode.y(),
    rotation: shapeNode.rotation(),
  }

  if ('width' in element) {
    patch.width = Math.max(20, (element.width || 0) * Math.abs(shapeNode.scaleX()))
  }

  if ('height' in element) {
    patch.height = Math.max(20, (element.height || 0) * Math.abs(shapeNode.scaleY()))
  }

  if (element.type === 'text') {
    patch.fontSize = Math.max(10, element.fontSize * Math.abs(shapeNode.scaleY()))
  }

  if (element.type === 'arrow') {
    patch.width = element.width * shapeNode.scaleX()
    patch.height = element.height * shapeNode.scaleY()
  }

  return patch
}

const getIdsForElementSelection = (page, elementId) => {
  const target = page.elements.find((element) => element.id === elementId)

  if (!target) {
    return []
  }

  if (!target.groupId) {
    return [target.id]
  }

  return page.elements.filter((element) => element.groupId === target.groupId).map((element) => element.id)
}

const getDuplicateGroupMap = (elements) => {
  const map = new Map()

  elements.forEach((element) => {
    if (element.groupId && !map.has(element.groupId)) {
      map.set(element.groupId, createId('group'))
    }
  })

  return map
}

const snapBoundsToPage = (bounds, deltaX, deltaY, page) => {
  const moved = {
    x: bounds.x + deltaX,
    y: bounds.y + deltaY,
    width: bounds.width,
    height: bounds.height,
  }

  let nextX = moved.x
  let nextY = moved.y
  const vertical = []
  const horizontal = []

  const gridX = Math.round(moved.x / gridSize) * gridSize
  if (Math.abs(gridX - moved.x) <= snapTolerance) {
    nextX = gridX
    vertical.push(gridX)
  }

  const gridY = Math.round(moved.y / gridSize) * gridSize
  if (Math.abs(gridY - moved.y) <= snapTolerance) {
    nextY = gridY
    horizontal.push(gridY)
  }

  const centerX = moved.x + moved.width / 2
  const pageCenterX = page.width / 2
  if (Math.abs(centerX - pageCenterX) <= snapTolerance) {
    nextX = pageCenterX - moved.width / 2
    vertical.push(pageCenterX)
  }

  const centerY = moved.y + moved.height / 2
  const pageCenterY = page.height / 2
  if (Math.abs(centerY - pageCenterY) <= snapTolerance) {
    nextY = pageCenterY - moved.height / 2
    horizontal.push(pageCenterY)
  }

  const right = moved.x + moved.width
  if (Math.abs(right - page.width) <= snapTolerance) {
    nextX = page.width - moved.width
    vertical.push(page.width)
  }

  const bottom = moved.y + moved.height
  if (Math.abs(bottom - page.height) <= snapTolerance) {
    nextY = page.height - moved.height
    horizontal.push(page.height)
  }

  return {
    deltaX: nextX - bounds.x,
    deltaY: nextY - bounds.y,
    guides: { vertical, horizontal },
  }
}

const CanvasImageNode = ({ element, selected, draggable, onSelect, onDragStart, onDragMove, onDragEnd, onTransformEnd, registerRef }) => {
  const [image, setImage] = useState(null)

  useEffect(() => {
    if (!element.src) {
      return undefined
    }

    const nextImage = new window.Image()
    nextImage.onload = () => setImage(nextImage)
    nextImage.src = element.src

    return () => {
      nextImage.onload = null
    }
  }, [element.src])

  return (
    <Group
      ref={registerRef(element.id)}
      x={element.x}
      y={element.y}
      rotation={element.rotation}
      draggable={draggable}
      onClick={onSelect}
      onTap={onSelect}
      onDragStart={onDragStart}
      onDragMove={onDragMove}
      onDragEnd={onDragEnd}
      onTransformEnd={onTransformEnd}
    >
      <Rect
        width={element.width}
        height={element.height}
        fill="#f3f0e7"
        cornerRadius={18}
        shadowColor={selected ? '#da7c46' : '#19313a'}
        shadowBlur={selected ? 20 : 8}
        shadowOpacity={0.15}
      />
      {element.src && image?.src === element.src ? (
        <Image image={image} width={element.width} height={element.height} cornerRadius={18} opacity={element.opacity} />
      ) : (
        <Text text="Loading image" x={16} y={16} width={Math.max(40, element.width - 32)} fontSize={20} fill="#526670" />
      )}
    </Group>
  )
}

const TableNode = ({ element, selected, draggable, onSelect, onDragStart, onDragMove, onDragEnd, onTransformEnd, registerRef }) => {
  const cellWidth = element.width / element.cols
  const cellHeight = element.height / element.rows

  return (
    <Group
      ref={registerRef(element.id)}
      x={element.x}
      y={element.y}
      rotation={element.rotation}
      draggable={draggable}
      onClick={onSelect}
      onTap={onSelect}
      onDragStart={onDragStart}
      onDragMove={onDragMove}
      onDragEnd={onDragEnd}
      onTransformEnd={onTransformEnd}
    >
      <Rect
        width={element.width}
        height={element.height}
        fill={element.fill}
        stroke={element.stroke}
        strokeWidth={element.strokeWidth}
        cornerRadius={16}
        shadowColor={selected ? '#da7c46' : '#19313a'}
        shadowBlur={selected ? 18 : 8}
        shadowOpacity={0.12}
      />
      <Rect width={element.width} height={cellHeight} fill={element.headerFill} cornerRadius={[16, 16, 0, 0]} />
      {Array.from({ length: element.cols - 1 }, (_, index) => (
        <Line
          key={`table-column-${element.id}-${index}`}
          points={[(index + 1) * cellWidth, 0, (index + 1) * cellWidth, element.height]}
          stroke={element.stroke}
          strokeWidth={element.strokeWidth}
        />
      ))}
      {Array.from({ length: element.rows - 1 }, (_, index) => (
        <Line
          key={`table-row-${element.id}-${index}`}
          points={[0, (index + 1) * cellHeight, element.width, (index + 1) * cellHeight]}
          stroke={element.stroke}
          strokeWidth={element.strokeWidth}
        />
      ))}
      {element.cells.map((row, rowIndex) =>
        row.map((cell, colIndex) => (
          <Text
            key={`cell-${element.id}-${rowIndex}-${colIndex}`}
            x={colIndex * cellWidth + 12}
            y={rowIndex * cellHeight + 10}
            width={cellWidth - 24}
            height={cellHeight - 20}
            fontSize={element.fontSize}
            fontStyle={rowIndex === 0 ? 'bold' : 'normal'}
            fill={element.textColor}
            verticalAlign="middle"
            text={cell}
          />
        )),
      )}
    </Group>
  )
}

function App() {
  const initialDocument = useMemo(() => readStoredDocument(), [])
  const [documentData, setDocumentData] = useState(initialDocument)
  const [currentPageId, setCurrentPageId] = useState(initialDocument.pages[0].id)
  const [selectedIds, setSelectedIds] = useState([])
  const [activeTool, setActiveTool] = useState('select')
  const [drawingState, setDrawingState] = useState(null)
  const [selectionBox, setSelectionBox] = useState(null)
  const [viewportWidth, setViewportWidth] = useState(900)
  const [snapGuides, setSnapGuides] = useState({ vertical: [], horizontal: [] })
  const [llmConfig, setLlmConfig] = useState(() => readStoredLlmConfig())
  const [llmAction, setLlmAction] = useState('draft')
  const [llmPrompt, setLlmPrompt] = useState('Create a polished proposal with title, body copy, diagram blocks, and KPI table.')
  const [llmStatus, setLlmStatus] = useState({ type: 'idle', message: '' })
  const [library, setLibrary] = useState([])
  const [libraryQuery, setLibraryQuery] = useState('')
  const [libraryStatus, setLibraryStatus] = useState({ type: 'idle', message: '' })
  const [saveStatus, setSaveStatus] = useState({ type: 'idle', message: '' })
  const [imageTargetId, setImageTargetId] = useState(null)
  const [textEditor, setTextEditor] = useState(null)
  const [uploads, setUploads] = useState([])
  const [uploadStatus, setUploadStatus] = useState({ type: 'idle', message: '' })
  const [versions, setVersions] = useState([])
  const [versionStatus, setVersionStatus] = useState({ type: 'idle', message: '' })

  const stageRef = useRef(null)
  const transformerRef = useRef(null)
  const canvasShellRef = useRef(null)
  const canvasFrameRef = useRef(null)
  const imageInputRef = useRef(null)
  const documentInputRef = useRef(null)
  const nodeRefs = useRef(new Map())
  const dragSessionRef = useRef(null)

  const currentPage = useMemo(
    () => documentData.pages.find((page) => page.id === currentPageId) || documentData.pages[0],
    [currentPageId, documentData.pages],
  )

  const selectedElements = useMemo(
    () => currentPage.elements.filter((element) => selectedIds.includes(element.id)),
    [currentPage.elements, selectedIds],
  )

  const primarySelectedElement = selectedElements.length === 1 ? selectedElements[0] : null

  const canvasScale = useMemo(() => {
    const maxWidth = Math.max(320, viewportWidth - 28)
    return Math.min(1, maxWidth / currentPage.width)
  }, [currentPage.width, viewportWidth])

  useEffect(() => {
    window.localStorage.setItem(LOCAL_DRAFT_KEY, JSON.stringify(documentData))
  }, [documentData])

  useEffect(() => {
    window.localStorage.setItem(LOCAL_LLM_KEY, JSON.stringify(llmConfig))
  }, [llmConfig])

  useEffect(() => {
    if (!documentData.pages.some((page) => page.id === currentPageId)) {
      setCurrentPageId(documentData.pages[0].id)
      setSelectedIds([])
    }
  }, [currentPageId, documentData.pages])

  useEffect(() => {
    if (!canvasShellRef.current) {
      return undefined
    }

    const updateWidth = () => setViewportWidth(canvasShellRef.current?.clientWidth || 900)
    updateWidth()

    const observer = new ResizeObserver(updateWidth)
    observer.observe(canvasShellRef.current)

    return () => observer.disconnect()
  }, [])

  useEffect(() => {
    const loadLibrary = async () => {
      try {
        setLibraryStatus({ type: 'loading', message: 'Loading server documents...' })
        const documents = await fetchDocumentLibrary(libraryQuery)
        setLibrary(documents)
        setLibraryStatus({ type: 'success', message: 'Server library ready.' })
      } catch (error) {
        setLibraryStatus({
          type: 'error',
          message: error instanceof Error ? error.message : 'Could not connect to server.',
        })
      }
    }

    loadLibrary()
  }, [libraryQuery])

  useEffect(() => {
    const loadUploads = async () => {
      try {
        setUploadStatus({ type: 'loading', message: 'Loading uploads...' })
        const nextUploads = await fetchUploads()
        setUploads(nextUploads)
        setUploadStatus({ type: 'success', message: 'Upload library ready.' })
      } catch (error) {
        setUploadStatus({
          type: 'error',
          message: error instanceof Error ? error.message : 'Could not load uploads.',
        })
      }
    }

    loadUploads()
  }, [])

  useEffect(() => {
    const loadVersions = async () => {
      if (!documentData.id) {
        setVersions([])
        return
      }

      try {
        setVersionStatus({ type: 'loading', message: 'Loading version history...' })
        const nextVersions = await fetchDocumentVersions(documentData.id)
        setVersions(nextVersions)
        setVersionStatus({ type: 'success', message: 'Version history ready.' })
      } catch (error) {
        setVersionStatus({
          type: 'error',
          message: error instanceof Error ? error.message : 'Could not load version history.',
        })
      }
    }

    loadVersions()
  }, [documentData.id, documentData.updatedAt])

  useEffect(() => {
    const transformer = transformerRef.current

    if (!transformer) {
      return
    }

    if (selectedElements.length > 1) {
      const nodes = selectedElements
        .map((element) => nodeRefs.current.get(element.id))
        .filter(Boolean)

      if (nodes.length > 0) {
        transformer.nodes(nodes)
        transformer.getLayer()?.batchDraw()
        return
      }
    }

    if (selectedElements.length === 1 && TRANSFORMABLE_TYPES.has(selectedElements[0].type)) {
      const node = nodeRefs.current.get(selectedElements[0].id)

      if (node) {
        transformer.nodes([node])
        transformer.getLayer()?.batchDraw()
        return
      }
    }

    transformer.nodes([])
    transformer.getLayer()?.batchDraw()
  }, [selectedElements])

  useEffect(() => {
    const handleKeyDown = (event) => {
      const targetTag = event.target?.tagName
      const isTyping = targetTag === 'INPUT' || targetTag === 'TEXTAREA'

      if ((event.key === 'Escape' || event.key === 'Esc') && textEditor) {
        setTextEditor(null)
        return
      }

      if (isTyping) {
        return
      }

      if ((event.key === 'Delete' || event.key === 'Backspace') && selectedIds.length > 0) {
        setDocumentData((prev) =>
          updateTimestamp({
            ...prev,
            pages: prev.pages.map((page) =>
              page.id === currentPage.id
                ? { ...page, elements: page.elements.filter((element) => !selectedIds.includes(element.id)) }
                : page,
            ),
          }),
        )
        setSelectedIds([])
        return
      }

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'd' && selectedIds.length > 0) {
        event.preventDefault()
        const groupMap = getDuplicateGroupMap(selectedElements)
        const duplicates = selectedElements.map((element) => {
          const duplicate = duplicateElement(element)
          return element.groupId ? { ...duplicate, groupId: groupMap.get(element.groupId) } : duplicate
        })

        updateCurrentPage((page) => ({ ...page, elements: [...page.elements, ...duplicates] }))
        setSelectedIds(duplicates.map((element) => element.id))
        return
      }

      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'g' && !event.shiftKey && selectedIds.length > 1) {
        event.preventDefault()
        const groupId = createId('group')
        patchElementsByIds(selectedIds, (element) => ({ ...element, groupId }))
        return
      }

      if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.key.toLowerCase() === 'g' && selectedIds.length > 0) {
        event.preventDefault()
        patchElementsByIds(selectedIds, (element) => ({ ...element, groupId: null }))
        return
      }
    }

    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [currentPage.id, patchElementsByIds, selectedElements, selectedIds, textEditor, updateCurrentPage])

  const updatePages = useCallback((updater) => {
    setDocumentData((prev) => updateTimestamp(updater(prev)))
  }, [])

  const updateCurrentPage = useCallback((pageUpdater) => {
    updatePages((prev) => ({
      ...prev,
      pages: prev.pages.map((page) => (page.id === currentPage.id ? pageUpdater(page) : page)),
    }))
  }, [currentPage.id, updatePages])

  const replaceCurrentElements = (elements) => {
    updateCurrentPage((page) => ({ ...page, elements }))
  }

  const patchElementsByIds = useCallback((elementIds, updater) => {
    updateCurrentPage((page) => ({
      ...page,
      elements: page.elements.map((element) => {
        if (!elementIds.includes(element.id)) {
          return element
        }

        return sanitizeElement(updater(element), element)
      }),
    }))
  }, [updateCurrentPage])

  const registerRef = (id) => (node) => {
    if (node) {
      nodeRefs.current.set(id, node)
    } else {
      nodeRefs.current.delete(id)
    }
  }

  const applyLoadedDocument = (nextDocument) => {
    const safeDocument = sanitizeDocument(nextDocument)
    setDocumentData(safeDocument)
    setCurrentPageId(safeDocument.pages[0].id)
    setSelectedIds([])
    setTextEditor(null)
  }

  const refreshLibrary = async () => {
    try {
      const documents = await fetchDocumentLibrary(libraryQuery)
      setLibrary(documents)
      setLibraryStatus({ type: 'success', message: 'Server library refreshed.' })
    } catch (error) {
      setLibraryStatus({
        type: 'error',
        message: error instanceof Error ? error.message : 'Could not refresh library.',
      })
    }
  }

  const refreshUploads = async () => {
    try {
      const nextUploads = await fetchUploads()
      setUploads(nextUploads)
      setUploadStatus({ type: 'success', message: 'Upload library refreshed.' })
    } catch (error) {
      setUploadStatus({
        type: 'error',
        message: error instanceof Error ? error.message : 'Could not refresh uploads.',
      })
    }
  }

  const refreshVersions = async () => {
    if (!documentData.id) {
      setVersions([])
      return
    }

    try {
      const nextVersions = await fetchDocumentVersions(documentData.id)
      setVersions(nextVersions)
      setVersionStatus({ type: 'success', message: 'Version history refreshed.' })
    } catch (error) {
      setVersionStatus({
        type: 'error',
        message: error instanceof Error ? error.message : 'Could not refresh version history.',
      })
    }
  }

  const saveDocumentJson = () => {
    const blob = new Blob([JSON.stringify(documentData, null, 2)], { type: 'application/json' })
    downloadBlob(blob, `${toSlug(documentData.title)}.poword.json`)
  }

  const loadDocumentJson = async (event) => {
    const file = event.target.files?.[0]
    event.target.value = ''

    if (!file) {
      return
    }

    const text = await readFileAsText(file)
    applyLoadedDocument(JSON.parse(text))
  }

  const saveServerDocument = async () => {
    try {
      setSaveStatus({ type: 'loading', message: 'Saving document to server...' })
      const saved = await saveDocumentToServer(documentData)
      setDocumentData(saved)
      setSaveStatus({ type: 'success', message: 'Saved to backend JSON storage.' })
      await refreshLibrary()
      await refreshVersions()
    } catch (error) {
      setSaveStatus({
        type: 'error',
        message: error instanceof Error ? error.message : 'Failed to save document.',
      })
    }
  }

  const loadServerDocument = async (documentId) => {
    try {
      setLibraryStatus({ type: 'loading', message: 'Loading server document...' })
      const nextDocument = await fetchDocumentById(documentId)
      applyLoadedDocument(nextDocument)
      setLibraryStatus({ type: 'success', message: 'Loaded document from server.' })
    } catch (error) {
      setLibraryStatus({
        type: 'error',
        message: error instanceof Error ? error.message : 'Failed to load document.',
      })
    }
  }

  const openTextEditor = (element) => {
    if (!canvasFrameRef.current) {
      return
    }

    setTextEditor({
      id: element.id,
      value: element.text,
      x: element.x * canvasScale,
      y: element.y * canvasScale,
      width: Math.max(120, element.width * canvasScale),
      height: Math.max(48, element.height * canvasScale),
      fontSize: Math.max(14, element.fontSize * canvasScale),
      fontFamily: element.fontFamily,
      color: element.fill,
      align: element.align,
      rotation: element.rotation || 0,
    })
  }

  const commitTextEditor = () => {
    if (!textEditor) {
      return
    }

    patchElementsByIds([textEditor.id], (element) => ({ ...element, text: textEditor.value }))
    setTextEditor(null)
  }

  const clearSelection = () => setSelectedIds([])

  const setElementSelection = (elementId, additive = false) => {
    const ids = getIdsForElementSelection(currentPage, elementId)

    if (!additive) {
      setSelectedIds(ids)
      return
    }

    setSelectedIds((prev) => {
      const everySelected = ids.every((id) => prev.includes(id))

      if (everySelected) {
        return prev.filter((id) => !ids.includes(id))
      }

      return Array.from(new Set([...prev, ...ids]))
    })
  }

  const handleDragStart = (elementId) => {
    const ids = selectedIds.includes(elementId) ? selectedIds : getIdsForElementSelection(currentPage, elementId)
    const movingElements = currentPage.elements.filter((element) => ids.includes(element.id))
    const bounds = getSelectionBounds(movingElements)
    const anchor = currentPage.elements.find((element) => element.id === elementId)

    setSelectedIds(ids)

    if (!anchor || !bounds) {
      return
    }

    dragSessionRef.current = {
      ids,
      anchorId: elementId,
      anchorStart: { x: anchor.x, y: anchor.y },
      startPositions: Object.fromEntries(movingElements.map((element) => [element.id, { x: element.x, y: element.y }])),
      bounds,
    }
  }

  const handleDragMove = (elementId, event) => {
    const session = dragSessionRef.current

    if (!session || session.anchorId !== elementId) {
      return
    }

    const rawDeltaX = event.target.x() - session.anchorStart.x
    const rawDeltaY = event.target.y() - session.anchorStart.y
    const snapped = snapBoundsToPage(session.bounds, rawDeltaX, rawDeltaY, currentPage)

    setSnapGuides(snapped.guides)

    updateCurrentPage((page) => ({
      ...page,
      elements: page.elements.map((element) => {
        if (!session.ids.includes(element.id)) {
          return element
        }

        const origin = session.startPositions[element.id]
        return sanitizeElement({ ...element, x: origin.x + snapped.deltaX, y: origin.y + snapped.deltaY }, element)
      }),
    }))
  }

  const handleDragEnd = () => {
    dragSessionRef.current = null
    setSnapGuides({ vertical: [], horizontal: [] })
  }

  const handleStageMouseDown = (event) => {
    if (textEditor) {
      commitTextEditor()
    }

    const stage = event.target.getStage()
    const point = getPointerPosition(stage, canvasScale)

    if (!point) {
      return
    }

    const clickedCanvas = getTargetIsCanvas(event.target, stage)

    if (activeTool === 'select' && clickedCanvas) {
      const additive = event.evt.shiftKey || event.evt.ctrlKey || event.evt.metaKey
      setSelectionBox({ start: point, end: point, additive })
      if (!additive) {
        clearSelection()
      }
      return
    }

    if (!clickedCanvas) {
      return
    }

    if (activeTool === 'text') {
      const element = createTextElement(point.x, point.y)
      updateCurrentPage((page) => ({ ...page, elements: [...page.elements, element] }))
      setSelectedIds([element.id])
      setTimeout(() => openTextEditor(element), 0)
      return
    }

    if (activeTool === 'table') {
      const element = createTableElement(point.x, point.y)
      updateCurrentPage((page) => ({ ...page, elements: [...page.elements, element] }))
      setSelectedIds([element.id])
      return
    }

    if (activeTool === 'rect' || activeTool === 'ellipse' || activeTool === 'arrow' || activeTool === 'pen') {
      const draft = createDrawElement(activeTool, point)

      if (!draft) {
        return
      }

      updateCurrentPage((page) => ({ ...page, elements: [...page.elements, draft] }))
      setSelectedIds([draft.id])
      setDrawingState({ id: draft.id, start: point, tool: activeTool })
    }
  }

  const handleStageMouseMove = () => {
    const stage = stageRef.current
    const point = stage ? getPointerPosition(stage, canvasScale) : null

    if (!point) {
      return
    }

    if (selectionBox) {
      setSelectionBox((prev) => (prev ? { ...prev, end: point } : prev))
      return
    }

    if (!drawingState) {
      return
    }

    updateCurrentPage((page) => ({
      ...page,
      elements: page.elements.map((element) => {
        if (element.id !== drawingState.id) {
          return element
        }

        if (element.type === 'pen') {
          return {
            ...element,
            points: [...element.points, point.x - drawingState.start.x, point.y - drawingState.start.y],
          }
        }

        return finalizeDrawElement(element, drawingState.start, point)
      }),
    }))
  }

  const handleStageMouseUp = () => {
    if (selectionBox) {
      const box = normalizeBox(selectionBox.start, selectionBox.end, 1, 1)
      const ids = currentPage.elements
        .filter((element) => intersectsBox(getElementBounds(element), box))
        .flatMap((element) => getIdsForElementSelection(currentPage, element.id))

      setSelectedIds((prev) =>
        selectionBox.additive ? Array.from(new Set([...prev, ...ids])) : Array.from(new Set(ids)),
      )
    }

    setSelectionBox(null)
    setDrawingState(null)
  }

  const handleDuplicateSelected = () => {
    if (selectedElements.length === 0) {
      return
    }

    const groupMap = getDuplicateGroupMap(selectedElements)
    const duplicates = selectedElements.map((element) => {
      const duplicate = duplicateElement(element)
      return element.groupId ? { ...duplicate, groupId: groupMap.get(element.groupId) } : duplicate
    })

    updateCurrentPage((page) => ({ ...page, elements: [...page.elements, ...duplicates] }))
    setSelectedIds(duplicates.map((element) => element.id))
  }

  const handleBringToFront = () => {
    if (selectedIds.length === 0) {
      return
    }

    updateCurrentPage((page) => ({
      ...page,
      elements: [...page.elements.filter((element) => !selectedIds.includes(element.id)), ...page.elements.filter((element) => selectedIds.includes(element.id))],
    }))
  }

  const handleSendToBack = () => {
    if (selectedIds.length === 0) {
      return
    }

    updateCurrentPage((page) => ({
      ...page,
      elements: [...page.elements.filter((element) => selectedIds.includes(element.id)), ...page.elements.filter((element) => !selectedIds.includes(element.id))],
    }))
  }

  const handleAlign = (mode) => {
    if (selectedIds.length < 2) {
      return
    }

    replaceCurrentElements(alignElements(currentPage.elements, selectedIds, mode))
  }

  const handleGroupSelected = () => {
    if (selectedIds.length < 2) {
      return
    }

    const groupId = createId('group')
    patchElementsByIds(selectedIds, (element) => ({ ...element, groupId }))
  }

  const handleUngroupSelected = () => {
    if (selectedIds.length === 0) {
      return
    }

    patchElementsByIds(selectedIds, (element) => ({ ...element, groupId: null }))
  }

  const handleAddPage = () => {
    const page = createBlankPage(`Page ${documentData.pages.length + 1}`)
    updatePages((prev) => ({ ...prev, pages: [...prev.pages, page] }))
    setCurrentPageId(page.id)
    setSelectedIds([])
  }

  const handleDuplicatePage = () => {
    const groupRemap = new Map()
    const pageCopy = {
      ...JSON.parse(JSON.stringify(currentPage)),
      id: createId('page'),
      name: `${currentPage.name} Copy`,
      elements: currentPage.elements.map((element) => {
        const duplicate = duplicateElement(element)

        if (element.groupId) {
          if (!groupRemap.has(element.groupId)) {
            groupRemap.set(element.groupId, createId('group'))
          }

          duplicate.groupId = groupRemap.get(element.groupId)
        }

        return duplicate
      }),
    }

    updatePages((prev) => ({ ...prev, pages: [...prev.pages, pageCopy] }))
    setCurrentPageId(pageCopy.id)
    setSelectedIds([])
  }

  const handleRemovePage = () => {
    if (documentData.pages.length === 1) {
      return
    }

    const remaining = documentData.pages.filter((page) => page.id !== currentPage.id)
    updatePages((prev) => ({ ...prev, pages: remaining }))
    setCurrentPageId(remaining[0].id)
    setSelectedIds([])
  }

  const handleInsertPreset = (preset) => {
    const anchorX = 110 + (currentPage.elements.length % 3) * 48
    const anchorY = 120 + (currentPage.elements.length % 5) * 42

    const element =
      preset === 'text'
        ? createTextElement(anchorX, anchorY)
        : preset === 'rect' || preset === 'ellipse'
          ? createShapeElement(preset, anchorX, anchorY)
          : preset === 'arrow'
            ? createArrowElement(anchorX, anchorY)
            : createTableElement(anchorX, anchorY)

    updateCurrentPage((page) => ({ ...page, elements: [...page.elements, element] }))
    setSelectedIds([element.id])
  }

  const handleImageImportRequest = (targetId = null) => {
    setImageTargetId(targetId)
    imageInputRef.current?.click()
  }

  const handleImageImport = async (event) => {
    const file = event.target.files?.[0]
    event.target.value = ''

    if (!file) {
      return
    }

    const uploaded = await uploadImageToServer(file)
    const src = uploaded.url
    await refreshUploads()

    if (imageTargetId) {
      patchElementsByIds([imageTargetId], (element) => ({ ...element, src }))
      setSelectedIds([imageTargetId])
      setImageTargetId(null)
      return
    }

    const image = createImageElement(src, 120, 140)
    updateCurrentPage((page) => ({ ...page, elements: [...page.elements, image] }))
    setSelectedIds([image.id])
  }

  const handleExportPng = () => {
    const stage = stageRef.current

    if (!stage) {
      return
    }

    const dataUrl = stage.toDataURL({ pixelRatio: 2 / canvasScale, mimeType: 'image/png' })
    const binary = atob(dataUrl.split(',')[1])
    const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0))
    downloadBlob(new Blob([bytes], { type: 'image/png' }), `${toSlug(documentData.title)}-${toSlug(currentPage.name)}.png`)
  }

  const handleExportPdf = async () => {
    await exportDocumentToPdf(documentData)
  }

  const handleRenameLibraryItem = async (item) => {
    const nextTitle = window.prompt('Rename document', item.title)

    if (!nextTitle || nextTitle === item.title) {
      return
    }

    await renameDocumentOnServer(item.id, { title: nextTitle })

    if (documentData.id === item.id) {
      setDocumentData((prev) => ({ ...prev, title: nextTitle }))
    }

    await refreshLibrary()
    await refreshVersions()
  }

  const handleDeleteLibraryItem = async (item) => {
    const confirmed = window.confirm(`Delete "${item.title}" from backend storage?`)

    if (!confirmed) {
      return
    }

    await deleteDocumentFromServer(item.id)

    if (documentData.id === item.id) {
      applyLoadedDocument(createDefaultDocument())
    }

    await refreshLibrary()
    await refreshVersions()
  }

  const handleDeleteUpload = async (item) => {
    const confirmed = window.confirm(`Delete upload file "${item.name}"?`)

    if (!confirmed) {
      return
    }

    await deleteUploadFromServer(item.name)
    await refreshUploads()
  }

  const handleCleanupUploads = async () => {
    const removed = await cleanupUploadsOnServer()
    setUploadStatus({
      type: 'success',
      message: removed.length > 0 ? `${removed.length} unused uploads removed.` : 'No unused uploads found.',
    })
    await refreshUploads()
  }

  const handleRestoreVersion = async (versionId) => {
    if (!documentData.id) {
      return
    }

    const restored = await restoreDocumentVersionApi(documentData.id, versionId)
    applyLoadedDocument(restored)
    setVersionStatus({ type: 'success', message: 'Version restored.' })
    await refreshLibrary()
    await refreshVersions()
  }

  const handleResetDocument = () => {
    const nextDocument = createDefaultDocument()
    applyLoadedDocument(nextDocument)
  }

  const handleLlmRun = async () => {
    if (!llmConfig.apiKey.trim()) {
      setLlmStatus({ type: 'error', message: 'Add an API key for the backend proxy request.' })
      return
    }

    if (llmAction === 'selection' && selectedElements.length === 0) {
      setLlmStatus({ type: 'error', message: 'Select one or more elements first.' })
      return
    }

    try {
      setLlmStatus({ type: 'loading', message: 'Generating canvas content...' })
      const { system, user } = buildPrompts({
        mode: llmAction,
        prompt: llmPrompt,
        documentData,
        currentPage,
        selection: selectedElements,
      })
      const payload = await requestLlmProxy({
        apiKey: llmConfig.apiKey,
        baseUrl: llmConfig.baseUrl,
        model: llmConfig.model,
        system,
        user,
      })
      const result = normalizeLlmPayload(payload)

      if (llmAction === 'pages') {
        const pages = Array.isArray(result.json.pages)
          ? result.json.pages.map((page, index) => ({
              ...createBlankPage(page.name || `AI Page ${index + 1}`),
              ...page,
              id: createId('page'),
              elements: Array.isArray(page.elements)
                ? page.elements.map((element, elementIndex) => sanitizeElement(element, { x: 90 + elementIndex * 12, y: 90 + elementIndex * 16 })).filter(Boolean)
                : [],
            }))
          : []

        if (pages.length === 0) {
          throw new Error('The model did not return any pages.')
        }

        updatePages((prev) => ({ ...prev, pages: [...prev.pages, ...pages] }))
        setCurrentPageId(pages[0].id)
        setSelectedIds([])
      } else if (llmAction === 'selection') {
        const updates = Array.isArray(result.json.elements) ? result.json.elements : []

        if (updates.length === 0) {
          throw new Error('The model did not return any element updates.')
        }

        updateCurrentPage((page) => ({
          ...page,
          elements: page.elements.map((element) => {
            const index = selectedIds.indexOf(element.id)

            if (index === -1) {
              return element
            }

            return sanitizeElement({ ...element, ...updates[Math.min(index, updates.length - 1)] }, element)
          }),
        }))
      } else {
        const elements = Array.isArray(result.json.elements)
          ? result.json.elements.map((element, index) => sanitizeElement(element, { x: 72 + (index % 4) * 52, y: 88 + (index % 5) * 54 })).filter(Boolean)
          : []

        if (elements.length === 0) {
          throw new Error('The model did not return any elements.')
        }

        updateCurrentPage((page) => ({ ...page, elements: [...page.elements, ...elements] }))
        setSelectedIds(elements.map((element) => element.id))
      }

      setLlmStatus({ type: 'success', message: 'AI result applied.' })
    } catch (error) {
      setLlmStatus({
        type: 'error',
        message: error instanceof Error ? error.message : 'AI request failed.',
      })
    }
  }

  const selectedSummary =
    selectedElements.length === 0
      ? 'No selection'
      : selectedElements.length === 1
        ? `${selectedElements[0].type} @ ${Math.round(selectedElements[0].x)}, ${Math.round(selectedElements[0].y)}`
        : `${selectedElements.length} elements selected`

  const alignmentButtons = [
    { id: 'left', label: 'Align left' },
    { id: 'center', label: 'Align center' },
    { id: 'right', label: 'Align right' },
    { id: 'top', label: 'Align top' },
    { id: 'middle', label: 'Align middle' },
    { id: 'bottom', label: 'Align bottom' },
  ]

  return (
    <div className="app-shell">
      <input ref={imageInputRef} type="file" accept="image/*" className="hidden-input" onChange={handleImageImport} />
      <input ref={documentInputRef} type="file" accept="application/json,.json,.poword.json" className="hidden-input" onChange={loadDocumentJson} />

      <header className="topbar">
        <div>
          <p className="eyebrow">PowordPointer</p>
          <input
            className="title-input"
            value={documentData.title}
            onChange={(event) => setDocumentData((prev) => updateTimestamp({ ...prev, title: event.target.value }))}
          />
          <p className="subtle">Canvas documents, inline text editing, PDF export, snap, align, group, and Koa-backed JSON storage.</p>
        </div>
        <div className="header-actions">
          <button className="ghost-button" onClick={() => documentInputRef.current?.click()}>Import JSON</button>
          <button className="ghost-button" onClick={saveDocumentJson}>Export JSON</button>
          <button className="ghost-button" onClick={saveServerDocument}>Save Server</button>
          <button className="ghost-button" onClick={handleExportPng}>Export PNG</button>
          <button className="ghost-button" onClick={handleExportPdf}>Export PDF</button>
          <button className="accent-button" onClick={handleResetDocument}>New document</button>
        </div>
      </header>

      <div className="workspace-grid">
        <aside className="left-panel panel">
          <div className="panel-block">
            <div className="block-header">
              <h2>Pages</h2>
              <span>{documentData.pages.length}</span>
            </div>
            <div className="page-list">
              {documentData.pages.map((page, index) => (
                <button
                  key={page.id}
                  className={`page-card ${page.id === currentPage.id ? 'active' : ''}`}
                  onClick={() => {
                    setCurrentPageId(page.id)
                    setSelectedIds([])
                  }}
                >
                  <span className="page-order">{index + 1}</span>
                  <strong>{page.name}</strong>
                  <small>{page.elements.length} elements</small>
                </button>
              ))}
            </div>
            <div className="row-actions">
              <button className="mini-button" onClick={handleAddPage}>Add page</button>
              <button className="mini-button" onClick={handleDuplicatePage}>Duplicate</button>
              <button className="mini-button danger" onClick={handleRemovePage}>Remove</button>
            </div>
          </div>

          <div className="panel-block">
            <div className="block-header">
              <h2>Backend Library</h2>
              <span>{library.length}</span>
            </div>
            <div className="row-actions">
              <button className="mini-button" onClick={refreshLibrary}>Refresh</button>
              <button className="mini-button" onClick={saveServerDocument}>Save current</button>
            </div>
            <label>
              <span>Search</span>
              <input value={libraryQuery} onChange={(event) => setLibraryQuery(event.target.value)} placeholder="Title or description" />
            </label>
            <p className={`status-pill ${libraryStatus.type}`}>{libraryStatus.message || 'Idle'}</p>
            <div className="page-list server-list">
              {library.map((item) => (
                <div key={item.id} className={`page-card ${item.id === documentData.id ? 'active' : ''}`}>
                  <button className="library-open" onClick={() => loadServerDocument(item.id)}>
                    <strong>{item.title}</strong>
                    <small>{item.pageCount} pages</small>
                    <small>{new Date(item.updatedAt).toLocaleString()}</small>
                  </button>
                  <div className="row-actions">
                    <button className="mini-button" onClick={() => handleRenameLibraryItem(item)}>Rename</button>
                    <button className="mini-button danger" onClick={() => handleDeleteLibraryItem(item)}>Delete</button>
                  </div>
                </div>
              ))}
            </div>
            <p className={`status-pill ${saveStatus.type}`}>{saveStatus.message || 'Not saved yet'}</p>
          </div>

          <div className="panel-block">
            <div className="block-header">
              <h2>Uploads</h2>
              <span>{uploads.length}</span>
            </div>
            <div className="row-actions">
              <button className="mini-button" onClick={refreshUploads}>Refresh</button>
              <button className="mini-button" onClick={handleCleanupUploads}>Cleanup unused</button>
            </div>
            <p className={`status-pill ${uploadStatus.type}`}>{uploadStatus.message || 'Idle'}</p>
            <div className="page-list server-list">
              {uploads.map((item) => (
                <div key={item.name} className="page-card">
                  <button className="library-open" onClick={() => {
                    const image = createImageElement(item.url, 120, 140)
                    updateCurrentPage((page) => ({ ...page, elements: [...page.elements, image] }))
                    setSelectedIds([image.id])
                  }}>
                    <strong>{item.name}</strong>
                    <small>{Math.round(item.size / 1024)} KB</small>
                    <small>Used by {item.usedBy} docs</small>
                  </button>
                  <div className="row-actions">
                    <button className="mini-button danger" onClick={() => handleDeleteUpload(item)}>Delete</button>
                  </div>
                </div>
              ))}
            </div>
          </div>

          <div className="panel-block">
            <div className="block-header">
              <h2>Versions</h2>
              <span>{versions.length}</span>
            </div>
            <div className="row-actions">
              <button className="mini-button" onClick={refreshVersions}>Refresh</button>
            </div>
            <p className={`status-pill ${versionStatus.type}`}>{versionStatus.message || 'Idle'}</p>
            <div className="page-list server-list">
              {versions.map((item) => (
                <div key={item.versionId} className="page-card">
                  <div className="library-open">
                    <strong>{item.title}</strong>
                    <small>{new Date(item.createdAt).toLocaleString()}</small>
                  </div>
                  <div className="row-actions">
                    <button className="mini-button" onClick={() => handleRestoreVersion(item.versionId)}>Restore</button>
                  </div>
                </div>
              ))}
            </div>
          </div>

          <div className="panel-block">
            <div className="block-header">
              <h2>Tools</h2>
              <span>{activeTool}</span>
            </div>
            <div className="tool-grid">
              {TOOL_OPTIONS.map((tool) => (
                <button key={tool.id} className={`tool-button ${activeTool === tool.id ? 'active' : ''}`} onClick={() => setActiveTool(tool.id)}>
                  {tool.label}
                </button>
              ))}
            </div>
            <div className="tool-grid compact">
              <button className="ghost-button" onClick={() => handleInsertPreset('text')}>Quick text</button>
              <button className="ghost-button" onClick={() => handleInsertPreset('rect')}>Panel</button>
              <button className="ghost-button" onClick={() => handleInsertPreset('ellipse')}>Circle</button>
              <button className="ghost-button" onClick={() => handleInsertPreset('arrow')}>Arrow</button>
              <button className="ghost-button" onClick={() => handleInsertPreset('table')}>Table</button>
              <button className="ghost-button" onClick={() => handleImageImportRequest()}>Image</button>
            </div>
          </div>

          <div className="panel-block">
            <div className="block-header">
              <h2>Current page</h2>
              <span>{currentPage.width} x {currentPage.height}</span>
            </div>
            <label>
              <span>Page name</span>
              <input value={currentPage.name} onChange={(event) => updateCurrentPage((page) => ({ ...page, name: event.target.value }))} />
            </label>
            <label>
              <span>Background</span>
              <input value={currentPage.background} onChange={(event) => updateCurrentPage((page) => ({ ...page, background: event.target.value }))} />
            </label>
          </div>
        </aside>

        <main className="canvas-panel panel">
          <div className="canvas-toolbar">
            <div>
              <p className="eyebrow">Canvas</p>
              <h2>{currentPage.name}</h2>
            </div>
            <div className="canvas-meta">
              <span>{selectedSummary}</span>
              <span>{Math.round(canvasScale * 100)}% fit</span>
            </div>
          </div>

          <div className="canvas-shell" ref={canvasShellRef}>
            <div className="canvas-frame" ref={canvasFrameRef} style={{ width: currentPage.width * canvasScale }}>
              <Stage
                ref={stageRef}
                width={currentPage.width * canvasScale}
                height={currentPage.height * canvasScale}
                scaleX={canvasScale}
                scaleY={canvasScale}
                onMouseDown={handleStageMouseDown}
                onTouchStart={handleStageMouseDown}
                onMouseMove={handleStageMouseMove}
                onTouchMove={handleStageMouseMove}
                onMouseUp={handleStageMouseUp}
                onTouchEnd={handleStageMouseUp}
              >
                <Layer>
                  <Rect
                    name="page-background"
                    x={0}
                    y={0}
                    width={currentPage.width}
                    height={currentPage.height}
                    fill={currentPage.background}
                    cornerRadius={28}
                    shadowColor="#10232c"
                    shadowBlur={24}
                    shadowOpacity={0.14}
                  />

                  {snapGuides.vertical.map((x, index) => (
                    <Line key={`guide-v-${x}-${index}`} points={[x, 0, x, currentPage.height]} stroke="#d56f3e" dash={[8, 8]} strokeWidth={1.5} />
                  ))}
                  {snapGuides.horizontal.map((y, index) => (
                    <Line key={`guide-h-${y}-${index}`} points={[0, y, currentPage.width, y]} stroke="#d56f3e" dash={[8, 8]} strokeWidth={1.5} />
                  ))}

                  {currentPage.elements.map((element) => {
                    const selected = selectedIds.includes(element.id)
                    const draggable = activeTool === 'select'
                    const onSelect = (event) => {
                      event.cancelBubble = true
                      setElementSelection(element.id, event.evt.shiftKey || event.evt.ctrlKey || event.evt.metaKey)
                    }

                    const sharedProps = {
                      key: element.id,
                      selected,
                      draggable,
                      registerRef,
                      onSelect,
                      onDragStart: () => handleDragStart(element.id),
                      onDragMove: (event) => handleDragMove(element.id, event),
                      onDragEnd: handleDragEnd,
                      onTransformEnd: () => {
                        if (selectedIds.length > 1) {
                          const nodes = transformerRef.current?.nodes() || []
                          const nextById = new Map(
                            nodes
                              .map((shapeNode) => {
                                const id = shapeNode?.attrs?.id || shapeNode?.attrs?.name
                                const targetElement = currentPage.elements.find((item) => item.id === id)

                                if (!id || !targetElement) {
                                  return null
                                }

                                return [id, extractElementPatchFromNode(shapeNode, targetElement)]
                              })
                              .filter(Boolean),
                          )

                          updateCurrentPage((page) => ({
                            ...page,
                            elements: page.elements.map((item) => {
                              const patch = nextById.get(item.id)?.[1]
                              return patch ? sanitizeElement({ ...item, ...patch }, item) : item
                            }),
                          }))

                          nodes.forEach((shapeNode) => {
                            shapeNode.scaleX(1)
                            shapeNode.scaleY(1)
                          })
                          return
                        }

                        const shapeNode = nodeRefs.current.get(element.id)

                        if (!shapeNode) {
                          return
                        }

                        const patch = extractElementPatchFromNode(shapeNode, element)
                        patchElementsByIds([element.id], (item) => ({ ...item, ...patch }))
                        shapeNode.scaleX(1)
                        shapeNode.scaleY(1)
                      },
                    }

                    const groupProps = {
                      ref: registerRef(element.id),
                      id: element.id,
                      name: element.id,
                      x: element.x,
                      y: element.y,
                      rotation: element.rotation,
                      draggable,
                    }

                    if (element.type === 'rect') {
                      return (
                        <Group {...sharedProps} {...groupProps}>
                          <Rect width={element.width} height={element.height} fill={element.fill} stroke={element.stroke} strokeWidth={element.strokeWidth} opacity={element.opacity} cornerRadius={22} shadowColor={selected ? '#da7c46' : '#10232c'} shadowBlur={selected ? 18 : 8} shadowOpacity={0.15} />
                        </Group>
                      )
                    }

                    if (element.type === 'ellipse') {
                      return (
                        <Group {...sharedProps} {...groupProps}>
                          <Ellipse x={element.width / 2} y={element.height / 2} radiusX={element.width / 2} radiusY={element.height / 2} fill={element.fill} stroke={element.stroke} strokeWidth={element.strokeWidth} opacity={element.opacity} shadowColor={selected ? '#da7c46' : '#10232c'} shadowBlur={selected ? 18 : 8} shadowOpacity={0.15} />
                        </Group>
                      )
                    }

                    if (element.type === 'arrow') {
                      return (
                        <Group {...sharedProps} {...groupProps}>
                          <Arrow points={[0, 0, element.width, element.height]} stroke={element.stroke} fill={element.stroke} strokeWidth={element.strokeWidth} pointerLength={element.pointerLength} pointerWidth={element.pointerWidth} opacity={element.opacity} />
                        </Group>
                      )
                    }

                    if (element.type === 'pen') {
                      return (
                        <Line
                          key={element.id}
                          x={element.x}
                          y={element.y}
                          points={element.points}
                          stroke={element.stroke}
                          strokeWidth={element.strokeWidth}
                          opacity={element.opacity}
                          lineCap="round"
                          lineJoin="round"
                          tension={0.2}
                          draggable={draggable}
                          onClick={onSelect}
                          onTap={onSelect}
                          onDragStart={() => handleDragStart(element.id)}
                          onDragMove={(event) => handleDragMove(element.id, event)}
                          onDragEnd={handleDragEnd}
                          shadowColor={selected ? '#da7c46' : undefined}
                          shadowBlur={selected ? 14 : 0}
                        />
                      )
                    }

                    if (element.type === 'text') {
                      return (
                        <Group
                          {...sharedProps}
                          {...groupProps}
                          onDblClick={() => openTextEditor(element)}
                        >
                          <Text width={element.width} height={element.height} text={element.text} fontSize={element.fontSize} fontFamily={element.fontFamily} fill={element.fill} align={element.align} lineHeight={element.lineHeight} fontStyle={element.fontStyle} padding={element.padding} />
                        </Group>
                      )
                    }

                    if (element.type === 'image') {
                      return <CanvasImageNode {...sharedProps} element={element} />
                    }

                    if (element.type === 'table') {
                      return <TableNode {...sharedProps} element={element} />
                    }

                    return null
                  })}

                  {selectionBox ? (
                    (() => {
                      const box = normalizeBox(selectionBox.start, selectionBox.end, 1, 1)
                      return <Rect x={box.x} y={box.y} width={box.width} height={box.height} stroke="#d56f3e" fill="rgba(213, 111, 62, 0.12)" dash={[10, 6]} />
                    })()
                  ) : null}

                  <Transformer ref={transformerRef} rotateEnabled borderStroke="#d56f3e" borderStrokeWidth={2} anchorFill="#fffaf2" anchorStroke="#d56f3e" anchorSize={10} keepRatio={false} />
                </Layer>
              </Stage>

              {textEditor ? (
                <textarea
                  className="inline-text-editor"
                  value={textEditor.value}
                  autoFocus
                  style={{
                    left: textEditor.x,
                    top: textEditor.y,
                    width: textEditor.width,
                    height: textEditor.height,
                    fontSize: textEditor.fontSize,
                    fontFamily: textEditor.fontFamily,
                    color: textEditor.color,
                    textAlign: textEditor.align,
                    transform: `rotate(${textEditor.rotation}deg)`,
                  }}
                  onChange={(event) => setTextEditor((prev) => ({ ...prev, value: event.target.value }))}
                  onBlur={commitTextEditor}
                  onKeyDown={(event) => {
                    if (event.key === 'Escape') {
                      setTextEditor(null)
                    }

                    if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
                      event.preventDefault()
                      commitTextEditor()
                    }
                  }}
                />
              ) : null}
            </div>
          </div>
        </main>

        <aside className="right-panel panel">
          <div className="panel-block">
            <div className="block-header">
              <h2>Selection</h2>
              <span>{selectedElements.length === 0 ? 'none' : selectedElements.length === 1 ? selectedElements[0].type : `${selectedElements.length} items`}</span>
            </div>

            {selectedElements.length > 0 ? (
              <>
                <div className="row-actions wrap-actions">
                  <button className="mini-button" onClick={handleDuplicateSelected}>Duplicate</button>
                  <button className="mini-button" onClick={handleBringToFront}>Front</button>
                  <button className="mini-button" onClick={handleSendToBack}>Back</button>
                  <button className="mini-button" onClick={handleGroupSelected}>Group</button>
                  <button className="mini-button" onClick={handleUngroupSelected}>Ungroup</button>
                </div>

                {selectedElements.length > 1 ? (
                  <>
                    <div className="block-header compact-header">
                      <h2>Align</h2>
                      <span>Snap to selection bounds</span>
                    </div>
                    <div className="tool-grid compact">
                      {alignmentButtons.map((item) => (
                        <button key={item.id} className="ghost-button" onClick={() => handleAlign(item.id)}>{item.label}</button>
                      ))}
                    </div>
                  </>
                ) : null}

                {primarySelectedElement ? (
                  <>
                    <div className="inspector-grid">
                      <label>
                        <span>X</span>
                        <input type="number" value={Math.round(primarySelectedElement.x)} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, x: Number(event.target.value) }))} />
                      </label>
                      <label>
                        <span>Y</span>
                        <input type="number" value={Math.round(primarySelectedElement.y)} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, y: Number(event.target.value) }))} />
                      </label>
                    </div>

                    {'width' in primarySelectedElement ? (
                      <div className="inspector-grid">
                        <label>
                          <span>Width</span>
                          <input type="number" value={Math.round(primarySelectedElement.width || 0)} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, width: Number(event.target.value) }))} />
                        </label>
                        <label>
                          <span>Height</span>
                          <input type="number" value={Math.round(primarySelectedElement.height || 0)} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, height: Number(event.target.value) }))} />
                        </label>
                      </div>
                    ) : null}

                    {'rotation' in primarySelectedElement ? (
                      <label>
                        <span>Rotation</span>
                        <input type="number" value={Math.round(primarySelectedElement.rotation || 0)} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, rotation: Number(event.target.value) }))} />
                      </label>
                    ) : null}

                    {'fill' in primarySelectedElement ? (
                      <label>
                        <span>Fill</span>
                        <input value={primarySelectedElement.fill || ''} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, fill: event.target.value }))} />
                      </label>
                    ) : null}

                    {'stroke' in primarySelectedElement ? (
                      <div className="inspector-grid">
                        <label>
                          <span>Stroke</span>
                          <input value={primarySelectedElement.stroke || ''} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, stroke: event.target.value }))} />
                        </label>
                        <label>
                          <span>Stroke width</span>
                          <input type="number" value={primarySelectedElement.strokeWidth || 1} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, strokeWidth: Number(event.target.value) }))} />
                        </label>
                      </div>
                    ) : null}

                    {primarySelectedElement.type === 'text' ? (
                      <>
                        <div className="row-actions">
                          <button className="mini-button" onClick={() => openTextEditor(primarySelectedElement)}>Inline edit</button>
                        </div>
                        <label>
                          <span>Text</span>
                          <textarea rows="5" value={primarySelectedElement.text} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, text: event.target.value }))} />
                        </label>
                        <div className="inspector-grid">
                          <label>
                            <span>Font size</span>
                            <input type="number" value={primarySelectedElement.fontSize} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, fontSize: Number(event.target.value) }))} />
                          </label>
                          <label>
                            <span>Align</span>
                            <select value={primarySelectedElement.align} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, align: event.target.value }))}>
                              <option value="left">Left</option>
                              <option value="center">Center</option>
                              <option value="right">Right</option>
                            </select>
                          </label>
                        </div>
                        <div className="inspector-grid">
                          <label>
                            <span>Line height</span>
                            <input type="number" step="0.05" value={primarySelectedElement.lineHeight} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, lineHeight: Number(event.target.value) }))} />
                          </label>
                          <label>
                            <span>Font family</span>
                            <input value={primarySelectedElement.fontFamily} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, fontFamily: event.target.value }))} />
                          </label>
                        </div>
                        <label>
                          <span>Font style</span>
                          <select value={primarySelectedElement.fontStyle} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, fontStyle: event.target.value }))}>
                            <option value="normal">Normal</option>
                            <option value="bold">Bold</option>
                            <option value="italic">Italic</option>
                          </select>
                        </label>
                      </>
                    ) : null}

                    {primarySelectedElement.type === 'table' ? (
                      <>
                        <div className="inspector-grid">
                          <label>
                            <span>Rows</span>
                            <input
                              type="number"
                              value={primarySelectedElement.rows}
                              onChange={(event) => {
                                const rows = Math.max(1, Number(event.target.value) || 1)
                                patchElementsByIds([primarySelectedElement.id], (element) => ({
                                  ...element,
                                  rows,
                                  cells: parseCells(serializeCells(primarySelectedElement.cells), rows, primarySelectedElement.cols),
                                }))
                              }}
                            />
                          </label>
                          <label>
                            <span>Cols</span>
                            <input
                              type="number"
                              value={primarySelectedElement.cols}
                              onChange={(event) => {
                                const cols = Math.max(1, Number(event.target.value) || 1)
                                patchElementsByIds([primarySelectedElement.id], (element) => ({
                                  ...element,
                                  cols,
                                  cells: parseCells(serializeCells(primarySelectedElement.cells), primarySelectedElement.rows, cols),
                                }))
                              }}
                            />
                          </label>
                        </div>
                        <label>
                          <span>Cells</span>
                          <textarea rows="8" value={serializeCells(primarySelectedElement.cells)} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, cells: parseCells(event.target.value, primarySelectedElement.rows, primarySelectedElement.cols) }))} />
                        </label>
                      </>
                    ) : null}

                    {primarySelectedElement.type === 'image' ? (
                      <button className="ghost-button" onClick={() => handleImageImportRequest(primarySelectedElement.id)}>Replace image</button>
                    ) : null}
                  </>
                ) : (
                  <p className="subtle">Multi-selection is active. Use align, group, layer order, duplicate, or delete shortcuts.</p>
                )}
              </>
            ) : (
              <p className="subtle">Click, shift-click, or drag a selection box. Grouped items select together and drag with snap guides.</p>
            )}
          </div>

          <div className="panel-block">
            <div className="block-header">
              <h2>LLM Studio</h2>
              <span>{LLM_ACTIONS.find((item) => item.id === llmAction)?.label}</span>
            </div>
            <label>
              <span>API key</span>
              <input type="password" value={llmConfig.apiKey} placeholder="Stored in localStorage for local use" onChange={(event) => setLlmConfig((prev) => ({ ...prev, apiKey: event.target.value }))} />
            </label>
            <label>
              <span>Base URL</span>
              <input value={llmConfig.baseUrl} onChange={(event) => setLlmConfig((prev) => ({ ...prev, baseUrl: event.target.value }))} />
            </label>
            <label>
              <span>Model</span>
              <input value={llmConfig.model} onChange={(event) => setLlmConfig((prev) => ({ ...prev, model: event.target.value }))} />
            </label>
            <label>
              <span>Action</span>
              <select value={llmAction} onChange={(event) => setLlmAction(event.target.value)}>
                {LLM_ACTIONS.map((item) => (
                  <option key={item.id} value={item.id}>{item.label}</option>
                ))}
              </select>
            </label>
            <label>
              <span>Prompt</span>
              <textarea rows="8" value={llmPrompt} onChange={(event) => setLlmPrompt(event.target.value)} />
            </label>
            <button className="accent-button full" onClick={handleLlmRun}>Run AI action</button>
            <p className={`status-pill ${llmStatus.type}`}>{llmStatus.message || 'Idle'}</p>
            <p className="subtle tiny">The React client now calls the Koa backend proxy, which forwards the OpenAI-compatible request and returns parsed JSON.</p>
          </div>
        </aside>
      </div>
    </div>
  )
}

export default App
