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
  distributeElements,
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
  createCommentOnServer,
  deleteCommentFromServer,
  deleteDocumentFromServer,
  deleteUploadFromServer,
  fetchComments,
  fetchCurrentUser,
  fetchDocumentVersions,
  fetchVersionDiff,
  fetchDocumentById,
  fetchDocumentLibrary,
  fetchTemplates,
  fetchUploads,
  getAuthToken,
  loginUser,
  renameDocumentOnServer,
  replyToCommentOnServer,
  registerUser,
  resolveCommentOnServer,
  requestLlmProxy,
  restoreDocumentVersion as restoreDocumentVersionApi,
  saveDocumentToServer,
  saveTemplateOnServer,
  seedTemplatesOnServer,
  setAuthToken,
  uploadImageToServer,
  deleteTemplateFromServer,
} from './lib/api'
import { normalizeLlmPayload, buildPrompts } from './lib/llm'
import { exportDocumentToPdf } from './lib/pdf'

const defaultLlmState = {
  apiKey: '',
  baseUrl: 'https://api.openai.com/v1',
  model: 'gpt-4.1-mini',
}

const MAX_HISTORY = 80
const AUTOSAVE_MS = 2500
const LOCAL_RECOVERY_KEY = 'powordpointer.recovery'

const extractCommentMeta = (text) => ({
  mentions: Array.from(new Set((text.match(/@[a-zA-Z0-9_\-.]+/g) || []).map((item) => item.slice(1)))),
  tags: Array.from(new Set((text.match(/#[a-zA-Z0-9_\-.]+/g) || []).map((item) => item.slice(1)))),
})

const renderCommentText = (text) =>
  `${text || ''}`
    .split(/(\s+)/)
    .map((token, index) => {
      if (/^@[a-zA-Z0-9_\-.]+$/.test(token)) {
        return <mark key={`${token}-${index}`} className="mention-token">{token}</mark>
      }

      if (/^#[a-zA-Z0-9_\-.]+$/.test(token)) {
        return <mark key={`${token}-${index}`} className="tag-token">{token}</mark>
      }

      return token
    })

const TemplatePreview = ({ template }) => {
  const page = template?.document?.pages?.[0]

  if (!page) {
    return null
  }

  const scale = 180 / page.width

  return (
    <div className="template-preview-canvas" style={{ width: page.width * scale, height: page.height * scale }}>
      <div className="template-preview-surface" style={{ background: page.background }}>
        {page.elements.slice(0, 8).map((element) => {
          const commonStyle = {
            left: element.x * scale,
            top: element.y * scale,
            width: (element.width || 120) * scale,
            height: (element.height || 60) * scale,
            transform: `rotate(${element.rotation || 0}deg)`,
          }

          if (element.type === 'text') {
            return (
              <div key={element.id} className="template-shape text" style={{ ...commonStyle, color: element.fill, fontSize: Math.max(8, element.fontSize * scale * 0.55) }}>
                {element.text}
              </div>
            )
          }

          if (element.type === 'ellipse') {
            return <div key={element.id} className="template-shape ellipse" style={{ ...commonStyle, background: element.fill, borderColor: element.stroke }} />
          }

          if (element.type === 'arrow') {
            return <div key={element.id} className="template-shape arrow" style={{ left: element.x * scale, top: element.y * scale, width: Math.max(20, Math.abs(element.width) * scale), borderTopColor: element.stroke }} />
          }

          return <div key={element.id} className={`template-shape ${element.type}`} style={{ ...commonStyle, background: element.fill || '#dce7ea', borderColor: element.stroke || '#23404d' }} />
        })}
      </div>
    </div>
  )
}

const invertHistoryEntry = (entry) => ({
  ...entry,
  nextDocument: entry.previousDocument,
  previousDocument: entry.nextDocument,
})

const createTemplates = () => [
  {
    id: 'proposal',
    title: 'Proposal deck',
    description: 'Title, summary, three-part structure, KPI table',
    document: sanitizeDocument({
      id: createId('doc'),
      title: 'Proposal Template',
      description: 'Proposal document template',
      createdAt: new Date().toISOString(),
      pages: [
        {
          ...createBlankPage('Overview'),
          elements: [
            createTextElement(88, 88, { text: 'Project Proposal', fontSize: 44, width: 700, height: 80 }),
            createTextElement(92, 182, { text: 'Executive summary\nKey opportunity\nExpected outcome', fontSize: 24, width: 540, height: 180 }),
            createShapeElement('rect', 640, 170, { width: 190, height: 200, fill: '#dfeef2' }),
            createTableElement(88, 460, { width: 760, height: 260 }),
          ],
        },
      ],
    }),
  },
  {
    id: 'meeting',
    title: 'Meeting notes',
    description: 'Agenda, decisions, owners, next steps',
    document: sanitizeDocument({
      id: createId('doc'),
      title: 'Meeting Notes Template',
      description: 'Structured meeting note template',
      createdAt: new Date().toISOString(),
      pages: [
        {
          ...createBlankPage('Notes'),
          elements: [
            createTextElement(88, 88, { text: 'Meeting Notes', fontSize: 42, width: 620, height: 70 }),
            createTextElement(88, 180, { text: 'Agenda\n-\n\nDecisions\n-\n\nAction items\n-', fontSize: 24, width: 520, height: 420, fontFamily: 'Avenir Next, Segoe UI Variable, Noto Sans KR, sans-serif' }),
            createTableElement(570, 180, { width: 290, height: 260, cols: 2, rows: 4, cells: [['Owner', 'Due'], ['Name', 'Date'], ['Name', 'Date'], ['Name', 'Date']] }),
          ],
        },
      ],
    }),
  },
  {
    id: 'report',
    title: 'Status report',
    description: 'Status, risks, metrics, timeline',
    document: sanitizeDocument({
      id: createId('doc'),
      title: 'Status Report Template',
      description: 'Operational report template',
      createdAt: new Date().toISOString(),
      pages: [
        {
          ...createBlankPage('Status'),
          elements: [
            createTextElement(88, 88, { text: 'Weekly Status Report', fontSize: 40, width: 700, height: 80 }),
            createShapeElement('rect', 88, 190, { width: 240, height: 180, fill: '#dcefe5' }),
            createShapeElement('rect', 360, 190, { width: 240, height: 180, fill: '#fff0d4' }),
            createShapeElement('rect', 632, 190, { width: 240, height: 180, fill: '#f8d8cf' }),
            createTableElement(88, 430, { width: 784, height: 280 }),
          ],
        },
      ],
    }),
  },
]

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

const snapBoundsToGuides = (bounds, deltaX, deltaY, page, guidePool) => {
  const moved = {
    x: bounds.x + deltaX,
    y: bounds.y + deltaY,
    width: bounds.width,
    height: bounds.height,
  }

  let bestDeltaX = deltaX
  let bestDeltaY = deltaY
  const guides = { vertical: [], horizontal: [] }
  const movingVerticals = [moved.x, moved.x + moved.width / 2, moved.x + moved.width]
  const movingHorizontals = [moved.y, moved.y + moved.height / 2, moved.y + moved.height]

  guidePool.vertical.forEach((guide) => {
    movingVerticals.forEach((position) => {
      if (Math.abs(position - guide) <= snapTolerance) {
        bestDeltaX += guide - position
        guides.vertical = [guide]
      }
    })
  })

  guidePool.horizontal.forEach((guide) => {
    movingHorizontals.forEach((position) => {
      if (Math.abs(position - guide) <= snapTolerance) {
        bestDeltaY += guide - position
        guides.horizontal = [guide]
      }
    })
  })

  const pageSnap = snapBoundsToPage(bounds, bestDeltaX, bestDeltaY, page)
  return {
    deltaX: pageSnap.deltaX,
    deltaY: pageSnap.deltaY,
    guides: {
      vertical: [...new Set([...guides.vertical, ...pageSnap.guides.vertical])],
      horizontal: [...new Set([...guides.horizontal, ...pageSnap.guides.horizontal])],
    },
  }
}

const collectSnapGuides = (page, movingIds) => {
  const vertical = [0, page.width / 2, page.width]
  const horizontal = [0, page.height / 2, page.height]

  page.elements.forEach((element) => {
    if (movingIds.includes(element.id)) {
      return
    }

    const bounds = getElementBounds(element)
    if (!bounds) {
      return
    }

    vertical.push(bounds.x, bounds.x + bounds.width / 2, bounds.x + bounds.width)
    horizontal.push(bounds.y, bounds.y + bounds.height / 2, bounds.y + bounds.height)
  })

  return { vertical, horizontal }
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
  const [autosaveEnabled] = useState(true)
  const [lastSavedDocument, setLastSavedDocument] = useState(null)
  const [recoveryDraft, setRecoveryDraft] = useState(null)
  const [imageTargetId, setImageTargetId] = useState(null)
  const [textEditor, setTextEditor] = useState(null)
  const [uploads, setUploads] = useState([])
  const [uploadStatus, setUploadStatus] = useState({ type: 'idle', message: '' })
  const [versions, setVersions] = useState([])
  const [versionStatus, setVersionStatus] = useState({ type: 'idle', message: '' })
  const [versionDiff, setVersionDiff] = useState(null)
  const [authMode, setAuthMode] = useState('login')
  const [authForm, setAuthForm] = useState({ email: '', password: '', name: '' })
  const [authStatus, setAuthStatus] = useState({ type: 'idle', message: '' })
  const [currentUser, setCurrentUser] = useState(null)
  const [comments, setComments] = useState([])
  const [commentStatus, setCommentStatus] = useState({ type: 'idle', message: '' })
  const [commentDraft, setCommentDraft] = useState('')
  const [replyDrafts, setReplyDrafts] = useState({})
  const [commentFilter, setCommentFilter] = useState('all')
  const [tagFilter, setTagFilter] = useState('')
  const [mentionSuggestions, setMentionSuggestions] = useState([])
  const [historyPast, setHistoryPast] = useState([])
  const [historyFuture, setHistoryFuture] = useState([])
  const [activeTemplateId, setActiveTemplateId] = useState('proposal')
  const [templates, setTemplates] = useState([])
  const [templateEditor, setTemplateEditor] = useState({ title: '', description: '' })
  const [templatePreviewId, setTemplatePreviewId] = useState('')
  const [layerExpanded, setLayerExpanded] = useState(() => {
    try {
      return JSON.parse(window.localStorage.getItem('powordpointer.layerExpanded') || '{}')
    } catch {
      return {}
    }
  })
  const [recoveryHistory, setRecoveryHistory] = useState(() => {
    try {
      return JSON.parse(window.localStorage.getItem('powordpointer.recovery.history') || '[]')
    } catch {
      return []
    }
  })

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
    window.localStorage.setItem('powordpointer.layerExpanded', JSON.stringify(layerExpanded))
  }, [layerExpanded])

  useEffect(() => {
    window.localStorage.setItem('powordpointer.recovery.history', JSON.stringify(recoveryHistory))
  }, [recoveryHistory])

  useEffect(() => {
    const raw = window.localStorage.getItem(LOCAL_RECOVERY_KEY)

    if (!raw) {
      return
    }

    try {
      const parsed = JSON.parse(raw)
      if (parsed?.document) {
        setRecoveryDraft(parsed)
      }
    } catch {
      return
    }
  }, [])

  useEffect(() => {
    if (!currentUser || !documentData.id) {
      setComments([])
      return
    }

    const loadComments = async () => {
      try {
        setCommentStatus({ type: 'loading', message: 'Loading comments...' })
        const nextComments = await fetchComments(documentData.id)
        setComments(nextComments)
        setCommentStatus({ type: 'success', message: 'Comments ready.' })
      } catch (error) {
        setCommentStatus({
          type: 'error',
          message: error instanceof Error ? error.message : 'Could not load comments.',
        })
      }
    }

    loadComments()
  }, [currentUser, documentData.id])

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
    const loadSession = async () => {
      if (!getAuthToken()) {
        return
      }

      try {
        const user = await fetchCurrentUser()
        setCurrentUser(user)
        setAuthStatus({ type: 'success', message: `Signed in as ${user.name}.` })
      } catch {
        setAuthToken('')
      }
    }

    loadSession()
  }, [])

  useEffect(() => {
    const loadLibrary = async () => {
      if (!currentUser) {
        setLibrary([])
        return
      }

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
  }, [currentUser, libraryQuery])

  useEffect(() => {
    if (!autosaveEnabled || !currentUser || !documentData.id) {
      return undefined
    }

    const timeout = window.setTimeout(async () => {
      const snapshot = JSON.stringify(documentData)

      if (snapshot === lastSavedDocument) {
        return
      }

      try {
        const saved = await saveDocumentToServer(documentData)
        setDocumentData(saved)
        setLastSavedDocument(JSON.stringify(saved))
        const recoveryEntry = { updatedAt: new Date().toISOString(), document: saved }
        window.localStorage.setItem(LOCAL_RECOVERY_KEY, JSON.stringify(recoveryEntry))
        setRecoveryHistory((prev) => [recoveryEntry, ...prev].slice(0, 8))
        setSaveStatus({ type: 'success', message: 'Autosaved to server.' })
      } catch (error) {
        setSaveStatus({
          type: 'error',
          message: error instanceof Error ? error.message : 'Autosave failed.',
        })
      }
    }, AUTOSAVE_MS)

    return () => window.clearTimeout(timeout)
  }, [autosaveEnabled, currentUser, documentData, lastSavedDocument])

  useEffect(() => {
    const loadTemplates = async () => {
      if (!currentUser) {
        setTemplates([])
        return
      }

      try {
        let nextTemplates = await fetchTemplates()

        if (nextTemplates.length === 0) {
          const seeded = createTemplates()
          await seedTemplatesOnServer(seeded)
          nextTemplates = await fetchTemplates()
        }

        setTemplates(nextTemplates)
        if (nextTemplates[0]) {
          setActiveTemplateId(nextTemplates[0].id)
          setTemplatePreviewId(nextTemplates[0].id)
        }
      } catch {
        setTemplates(createTemplates())
      }
    }

    loadTemplates()
  }, [currentUser])

  useEffect(() => {
    const loadUploads = async () => {
      if (!currentUser) {
        setUploads([])
        return
      }

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
  }, [currentUser])

  useEffect(() => {
    const loadVersions = async () => {
      if (!currentUser || !documentData.id) {
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
  }, [currentUser, documentData.id, documentData.updatedAt])

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
    setDocumentData((prev) => {
      const next = updateTimestamp(updater(prev))
      setHistoryPast((history) => [
        ...history.slice(-MAX_HISTORY + 1),
        {
          previousDocument: prev,
          nextDocument: next,
        },
      ])
      setHistoryFuture([])
      return next
    })
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

  const handleAuthSubmit = async () => {
    try {
      setAuthStatus({ type: 'loading', message: authMode === 'login' ? 'Signing in...' : 'Creating account...' })
      const payload =
        authMode === 'login'
          ? await loginUser({ email: authForm.email, password: authForm.password })
          : await registerUser(authForm)

      setAuthToken(payload.token)
      setCurrentUser(payload.user)
      setAuthStatus({ type: 'success', message: `Welcome, ${payload.user.name}.` })
    } catch (error) {
      setAuthStatus({
        type: 'error',
        message: error instanceof Error ? error.message : 'Authentication failed.',
      })
    }
  }

  const handleLogout = () => {
    setAuthToken('')
    setCurrentUser(null)
    setLibrary([])
    setUploads([])
    setVersions([])
    setVersionDiff(null)
    setComments([])
    setLastSavedDocument(null)
    setAuthStatus({ type: 'idle', message: 'Signed out.' })
  }

  const handleUndo = () => {
    setHistoryPast((past) => {
      if (past.length === 0) {
        return past
      }

      const previous = past[past.length - 1]
      setHistoryFuture((future) => [invertHistoryEntry(previous), ...future].slice(0, MAX_HISTORY))
      setDocumentData(previous.previousDocument)
      return past.slice(0, -1)
    })
  }

  const handleRedo = () => {
    setHistoryFuture((future) => {
      if (future.length === 0) {
        return future
      }

      const [next, ...rest] = future
      setHistoryPast((past) => [...past.slice(-MAX_HISTORY + 1), invertHistoryEntry(next)])
      setDocumentData(next.previousDocument)
      return rest
    })
  }

  const applyTemplate = (templateId) => {
    const template = templates.find((item) => item.id === templateId)

    if (!template) {
      return
    }

    const nextDocument = sanitizeDocument({
      ...template.document,
      id: createId('doc'),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    })

    setHistoryPast([])
    setHistoryFuture([])
    applyLoadedDocument(nextDocument)
  }

  const handleSaveTemplate = async () => {
    const templateId = activeTemplateId || createId('template')
    await saveTemplateOnServer({
      id: templateId,
      title: templateEditor.title || documentData.title,
      description: templateEditor.description || documentData.description,
      document: documentData,
    })
    const nextTemplates = await fetchTemplates()
    setTemplates(nextTemplates)
    setActiveTemplateId(templateId)
  }

  const handleDeleteTemplate = async () => {
    if (!activeTemplateId) {
      return
    }

    await deleteTemplateFromServer(activeTemplateId)
    const nextTemplates = await fetchTemplates()
    setTemplates(nextTemplates)
    setActiveTemplateId(nextTemplates[0]?.id || '')
  }

  const handleRecoverDraft = () => {
    if (!recoveryDraft?.document) {
      return
    }

    applyLoadedDocument(recoveryDraft.document)
    setRecoveryDraft(null)
    window.localStorage.removeItem(LOCAL_RECOVERY_KEY)
  }

  const handleRecoverHistoryEntry = (entry) => {
    if (!entry?.document) {
      return
    }

    applyLoadedDocument(entry.document)
  }

  const dismissRecoveryDraft = () => {
    setRecoveryDraft(null)
    window.localStorage.removeItem(LOCAL_RECOVERY_KEY)
  }

  const toggleLayerGroup = (groupId) => {
    setLayerExpanded((prev) => ({
      ...prev,
      [groupId]: prev[groupId] === false ? true : false,
    }))
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

  const refreshComments = async () => {
    if (!documentData.id) {
      setComments([])
      return
    }

    try {
      const nextComments = await fetchComments(documentData.id)
      setComments(nextComments)
      setCommentStatus({ type: 'success', message: 'Comments refreshed.' })
    } catch (error) {
      setCommentStatus({
        type: 'error',
        message: error instanceof Error ? error.message : 'Could not refresh comments.',
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
      setLastSavedDocument(JSON.stringify(saved))
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
      setLastSavedDocument(JSON.stringify(nextDocument))
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
    const guidePool = collectSnapGuides(currentPage, session.ids)
    const snapped = snapBoundsToGuides(session.bounds, rawDeltaX, rawDeltaY, currentPage, guidePool)

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

  const handleDistribute = (direction) => {
    if (selectedIds.length < 3) {
      return
    }

    replaceCurrentElements(distributeElements(currentPage.elements, selectedIds, direction))
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

  const handleLayerOrder = (direction) => {
    if (selectedIds.length === 0) {
      return
    }

    updateCurrentPage((page) => {
      const elements = [...page.elements]
      const targetIndex = elements.findIndex((element) => element.id === selectedIds[0])

      if (targetIndex === -1) {
        return page
      }

      const step = direction === 'up' ? 1 : -1
      const nextIndex = Math.max(0, Math.min(elements.length - 1, targetIndex + step))
      const [item] = elements.splice(targetIndex, 1)
      elements.splice(nextIndex, 0, item)
      return { ...page, elements }
    })
  }

  const handleLayerDragStart = (event, elementId) => {
    event.dataTransfer.setData('text/plain', elementId)
  }

  const handleLayerDrop = (event, targetId) => {
    event.preventDefault()
    const sourceId = event.dataTransfer.getData('text/plain')

    if (!sourceId || sourceId === targetId) {
      return
    }

    updateCurrentPage((page) => {
      const elements = [...page.elements]
      const sourceIndex = elements.findIndex((element) => element.id === sourceId)
      const targetIndex = elements.findIndex((element) => element.id === targetId)

      if (sourceIndex === -1 || targetIndex === -1) {
        return page
      }

      const [item] = elements.splice(sourceIndex, 1)
      elements.splice(targetIndex, 0, item)
      return { ...page, elements }
    })
  }

  const toggleLayerFlag = (elementId, key) => {
    patchElementsByIds([elementId], (element) => ({ ...element, [key]: !element[key] }))
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

  const handlePreviewVersionDiff = async (versionId) => {
    if (!documentData.id) {
      return
    }

    const diff = await fetchVersionDiff(documentData.id, versionId)
    setVersionDiff(diff)
  }

  const handleAddComment = async () => {
    if (!documentData.id || !commentDraft.trim()) {
      return
    }

    const comment = await createCommentOnServer(documentData.id, {
      body: commentDraft,
      pageId: currentPage.id,
      elementId: primarySelectedElement?.id || null,
    })

    setComments((prev) => [comment, ...prev])
    setCommentDraft('')
    setMentionSuggestions([])
    setCommentStatus({ type: 'success', message: 'Comment added.' })
  }

  const handleDeleteComment = async (commentId) => {
    await deleteCommentFromServer(commentId)
    setComments((prev) => prev.filter((comment) => comment.id !== commentId))
  }

  const filteredComments = comments.filter((comment) => {
    const meta = extractCommentMeta(comment.body)

    if (commentFilter === 'open') {
      return !comment.resolved
    }

    if (commentFilter === 'page') {
      return comment.pageId === currentPage.id
    }

    if (commentFilter === 'selection') {
      return primarySelectedElement ? comment.elementId === primarySelectedElement.id : false
    }

    if (commentFilter === 'tag') {
      return tagFilter ? meta.tags.includes(tagFilter) : meta.tags.length > 0
    }

    return true
  })

  const groupedLayerEntries = useMemo(() => {
    const entries = []
    const seenGroups = new Set()

    ;[...currentPage.elements].reverse().forEach((element) => {
      if (element.groupId) {
        if (seenGroups.has(element.groupId)) {
          return
        }

        seenGroups.add(element.groupId)
        entries.push({
          type: 'group',
          id: element.groupId,
          items: [...currentPage.elements].filter((item) => item.groupId === element.groupId).reverse(),
        })
        return
      }

      entries.push({ type: 'element', id: element.id, item: element })
    })

    return entries
  }, [currentPage.elements])

  const handleReplyToComment = async (commentId) => {
    const body = replyDrafts[commentId]?.trim()

    if (!documentData.id || !body) {
      return
    }

    const comment = await replyToCommentOnServer(documentData.id, commentId, body)
    setComments((prev) => [comment, ...prev])
    setReplyDrafts((prev) => ({ ...prev, [commentId]: '' }))
  }

  const handleResolveComment = async (commentId, resolved) => {
    const comment = await resolveCommentOnServer(commentId, resolved)
    setComments((prev) => prev.map((item) => (item.id === comment.id ? comment : item)))
  }

  const handleCommentDraftChange = (value) => {
    setCommentDraft(value)
    const mentionMatch = value.match(/@([a-zA-Z0-9_\-.]*)$/)

    if (!mentionMatch) {
      setMentionSuggestions([])
      return
    }

    const query = mentionMatch[1].toLowerCase()
    const candidates = Array.from(new Set(comments.map((comment) => comment.authorName))).filter((name) =>
      name.toLowerCase().includes(query),
    )
    setMentionSuggestions(candidates.slice(0, 5))
  }

  const handleInsertMention = (name) => {
    setCommentDraft((prev) => prev.replace(/@([a-zA-Z0-9_\-.]*)$/, `@${name} `))
    setMentionSuggestions([])
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
          <button className="ghost-button" onClick={handleUndo} disabled={historyPast.length === 0}>Undo</button>
          <button className="ghost-button" onClick={handleRedo} disabled={historyFuture.length === 0}>Redo</button>
          {currentUser ? <span className="user-badge">{currentUser.name}</span> : null}
          <button className="ghost-button" onClick={() => documentInputRef.current?.click()}>Import JSON</button>
          <button className="ghost-button" onClick={saveDocumentJson}>Export JSON</button>
          <button className="ghost-button" onClick={saveServerDocument} disabled={!currentUser}>Save Server</button>
          <button className="ghost-button" onClick={handleExportPng}>Export PNG</button>
          <button className="ghost-button" onClick={handleExportPdf}>Export PDF</button>
          <button className="accent-button" onClick={handleResetDocument}>New document</button>
          {currentUser ? (
            <button className="ghost-button" onClick={handleLogout}>Logout</button>
          ) : null}
        </div>
      </header>

      {!currentUser ? (
        <section className="auth-panel panel">
          <div className="panel-block auth-block">
            <div className="block-header">
              <h2>{authMode === 'login' ? 'Sign In' : 'Register'}</h2>
              <span>Auth required for server sync</span>
            </div>
            <label>
              <span>Email</span>
              <input value={authForm.email} onChange={(event) => setAuthForm((prev) => ({ ...prev, email: event.target.value }))} />
            </label>
            {authMode === 'register' ? (
              <label>
                <span>Name</span>
                <input value={authForm.name} onChange={(event) => setAuthForm((prev) => ({ ...prev, name: event.target.value }))} />
              </label>
            ) : null}
            <label>
              <span>Password</span>
              <input type="password" value={authForm.password} onChange={(event) => setAuthForm((prev) => ({ ...prev, password: event.target.value }))} />
            </label>
            <div className="row-actions">
              <button className="accent-button" onClick={handleAuthSubmit}>{authMode === 'login' ? 'Login' : 'Create account'}</button>
              <button className="ghost-button" onClick={() => setAuthMode((prev) => (prev === 'login' ? 'register' : 'login'))}>
                {authMode === 'login' ? 'Need account?' : 'Have account?'}
              </button>
            </div>
            <p className={`status-pill ${authStatus.type}`}>{authStatus.message || 'Enter credentials to continue.'}</p>
          </div>
        </section>
      ) : null}

      {recoveryDraft ? (
        <section className="recovery-banner panel">
          <div className="panel-block recovery-block">
            <div className="block-header">
              <h2>Recovery Available</h2>
              <span>{new Date(recoveryDraft.updatedAt).toLocaleString()}</span>
            </div>
            <p className="subtle">A recent autosaved draft is available to restore.</p>
            <div className="row-actions">
              <button className="accent-button" onClick={handleRecoverDraft}>Restore draft</button>
              <button className="ghost-button" onClick={dismissRecoveryDraft}>Dismiss</button>
            </div>
          </div>
        </section>
      ) : null}

      {recoveryHistory.length > 0 ? (
        <section className="recovery-banner panel">
          <div className="panel-block recovery-block">
            <div className="block-header">
              <h2>Recovery History</h2>
              <span>{recoveryHistory.length}</span>
            </div>
            <div className="page-list server-list">
              {recoveryHistory.map((entry, index) => (
                <div key={`${entry.updatedAt}-${index}`} className="page-card">
                  <div className="library-open">
                    <strong>{entry.document?.title || 'Untitled draft'}</strong>
                    <small>{new Date(entry.updatedAt).toLocaleString()}</small>
                  </div>
                  <div className="row-actions">
                    <button className="mini-button" onClick={() => handleRecoverHistoryEntry(entry)}>Restore</button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </section>
      ) : null}

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
              <h2>Templates</h2>
              <span>{templates.length}</span>
            </div>
            <label>
              <span>Template</span>
              <select value={activeTemplateId} onChange={(event) => {
                setActiveTemplateId(event.target.value)
                setTemplatePreviewId(event.target.value)
              }}>
                {templates.map((template) => (
                  <option key={template.id} value={template.id}>{template.title}</option>
                ))}
              </select>
            </label>
            <p className="subtle">{templates.find((template) => template.id === activeTemplateId)?.description}</p>
            {templates.find((template) => template.id === templatePreviewId)?.document?.pages?.[0] ? (
              <div className="template-preview-card">
                <strong>Preview</strong>
                <TemplatePreview template={templates.find((template) => template.id === templatePreviewId)} />
                <small>{templates.find((template) => template.id === templatePreviewId)?.document.pages[0].name}</small>
                <small>
                  {templates.find((template) => template.id === templatePreviewId)?.document.pages[0].elements.length || 0} elements on first page
                </small>
              </div>
            ) : null}
            <label>
              <span>Template title</span>
              <input value={templateEditor.title} onChange={(event) => setTemplateEditor((prev) => ({ ...prev, title: event.target.value }))} placeholder="Use current document title if empty" />
            </label>
            <label>
              <span>Template description</span>
              <textarea rows="3" value={templateEditor.description} onChange={(event) => setTemplateEditor((prev) => ({ ...prev, description: event.target.value }))} />
            </label>
            <div className="row-actions">
              <button className="ghost-button" onClick={() => applyTemplate(activeTemplateId)}>Apply template</button>
              <button className="ghost-button" onClick={handleSaveTemplate}>Save template</button>
              <button className="mini-button danger" onClick={handleDeleteTemplate} disabled={!activeTemplateId}>Delete</button>
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
                    <img className="upload-thumb" src={item.url} alt="" />
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
                    <small>{item.summary.pageCount} pages, {item.summary.elementCount} elements</small>
                  </div>
                  <div className="row-actions">
                    <button className="mini-button" onClick={() => handlePreviewVersionDiff(item.versionId)}>Diff</button>
                    <button className="mini-button" onClick={() => handleRestoreVersion(item.versionId)}>Restore</button>
                  </div>
                </div>
              ))}
            </div>
            {versionDiff ? (
              <div className="version-diff-card">
                <strong>Diff Preview</strong>
                <small>Version: {versionDiff.versionTitle}</small>
                <small>Current: {versionDiff.currentTitle}</small>
                <small>Page delta: {versionDiff.pageDelta}</small>
                <small>Element delta: {versionDiff.elementDelta}</small>
                <small>Added pages: {versionDiff.addedPageNames.join(', ') || 'None'}</small>
                <small>Removed pages: {versionDiff.removedPageNames.join(', ') || 'None'}</small>
                <div className="diff-grid">
                  <div>
                    <strong>Version Pages</strong>
                    {versionDiff.versionPages?.map((page) => (
                      <small key={`v-${page.id || page.name}`}>{page.name} ({page.elements?.length || 0})</small>
                    ))}
                  </div>
                  <div>
                    <strong>Current Pages</strong>
                    {versionDiff.currentPages?.map((page) => (
                      <small key={`c-${page.id || page.name}`}>{page.name} ({page.elements?.length || 0})</small>
                    ))}
                  </div>
                </div>
              </div>
            ) : null}
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

          <div className="panel-block">
            <div className="block-header">
              <h2>Layers</h2>
              <span>{currentPage.elements.length}</span>
            </div>
            <div className="page-list server-list">
              {groupedLayerEntries.map((entry) =>
                entry.type === 'group' ? (
                  <div key={entry.id} className="layer-group-card">
                    <button className="layer-group-toggle" onClick={() => toggleLayerGroup(entry.id)}>
                      <strong className="layer-group-title">{layerExpanded[entry.id] === false ? '[+]' : '[-]'} Group {entry.id}</strong>
                    </button>
                    {layerExpanded[entry.id] === false ? null : entry.items.map((element) => (
                      <div
                        key={element.id}
                        className={`page-card nested-layer ${selectedIds.includes(element.id) ? 'active' : ''}`}
                        draggable
                        onDragStart={(event) => handleLayerDragStart(event, element.id)}
                        onDragOver={(event) => event.preventDefault()}
                        onDrop={(event) => handleLayerDrop(event, element.id)}
                      >
                        <button className="library-open" onClick={() => setSelectedIds((prev) => prev.includes(element.id) ? prev.filter((id) => id !== element.id) : [...prev, element.id])}>
                          <strong>{element.type}</strong>
                          <small>{element.id}</small>
                        </button>
                        <div className="row-actions">
                          <button className="mini-button" onClick={() => toggleLayerFlag(element.id, 'hidden')}>{element.hidden ? 'Show' : 'Hide'}</button>
                          <button className="mini-button" onClick={() => toggleLayerFlag(element.id, 'locked')}>{element.locked ? 'Unlock' : 'Lock'}</button>
                        </div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div
                    key={entry.id}
                    className={`page-card ${selectedIds.includes(entry.item.id) ? 'active' : ''}`}
                    draggable
                    onDragStart={(event) => handleLayerDragStart(event, entry.item.id)}
                    onDragOver={(event) => event.preventDefault()}
                    onDrop={(event) => handleLayerDrop(event, entry.item.id)}
                  >
                    <button className="library-open" onClick={() => setSelectedIds((prev) => prev.includes(entry.item.id) ? prev.filter((id) => id !== entry.item.id) : [...prev, entry.item.id])}>
                      <strong>{entry.item.type}</strong>
                      <small>{entry.item.id}</small>
                    </button>
                    <div className="row-actions">
                      <button className="mini-button" onClick={() => toggleLayerFlag(entry.item.id, 'hidden')}>{entry.item.hidden ? 'Show' : 'Hide'}</button>
                      <button className="mini-button" onClick={() => toggleLayerFlag(entry.item.id, 'locked')}>{entry.item.locked ? 'Unlock' : 'Lock'}</button>
                    </div>
                  </div>
                ),
              )}
            </div>
            <div className="row-actions">
              <button className="mini-button" onClick={() => handleLayerOrder('up')}>Move up</button>
              <button className="mini-button" onClick={() => handleLayerOrder('down')}>Move down</button>
            </div>
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
                    const draggable = activeTool === 'select' && !element.locked
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
                          {element.hidden ? null : (
                          <Rect width={element.width} height={element.height} fill={element.fill} stroke={element.stroke} strokeWidth={element.strokeWidth} opacity={element.opacity} cornerRadius={22} shadowColor={selected ? '#da7c46' : '#10232c'} shadowBlur={selected ? 18 : 8} shadowOpacity={0.15} />
                          )}
                        </Group>
                      )
                    }

                    if (element.type === 'ellipse') {
                      return (
                        <Group {...sharedProps} {...groupProps}>
                          {element.hidden ? null : (
                          <Ellipse x={element.width / 2} y={element.height / 2} radiusX={element.width / 2} radiusY={element.height / 2} fill={element.fill} stroke={element.stroke} strokeWidth={element.strokeWidth} opacity={element.opacity} shadowColor={selected ? '#da7c46' : '#10232c'} shadowBlur={selected ? 18 : 8} shadowOpacity={0.15} />
                          )}
                        </Group>
                      )
                    }

                    if (element.type === 'arrow') {
                      return (
                        <Group {...sharedProps} {...groupProps}>
                          {element.hidden ? null : (
                          <Arrow points={[0, 0, element.width, element.height]} stroke={element.stroke} fill={element.stroke} strokeWidth={element.strokeWidth} pointerLength={element.pointerLength} pointerWidth={element.pointerWidth} opacity={element.opacity} />
                          )}
                        </Group>
                      )
                    }

                    if (element.type === 'pen') {
                      return element.hidden ? null : (
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
                      return element.hidden ? null : (
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
                      return element.hidden ? null : <CanvasImageNode {...sharedProps} element={element} />
                    }

                    if (element.type === 'table') {
                      return element.hidden ? null : <TableNode {...sharedProps} element={element} />
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
                      <button className="ghost-button" onClick={() => handleDistribute('horizontal')}>Distribute H</button>
                      <button className="ghost-button" onClick={() => handleDistribute('vertical')}>Distribute V</button>
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
                        <div className="row-actions wrap-actions">
                          <button
                            className="mini-button"
                            onClick={() => patchElementsByIds([primarySelectedElement.id], (element) => ({
                              ...element,
                              rows: element.rows + 1,
                              cells: [...element.cells, Array.from({ length: element.cols }, () => 'Value')],
                            }))}
                          >
                            Add row
                          </button>
                          <button
                            className="mini-button"
                            onClick={() => patchElementsByIds([primarySelectedElement.id], (element) => ({
                              ...element,
                              rows: Math.max(1, element.rows - 1),
                              cells: element.cells.slice(0, Math.max(1, element.rows - 1)),
                            }))}
                          >
                            Remove row
                          </button>
                          <button
                            className="mini-button"
                            onClick={() => patchElementsByIds([primarySelectedElement.id], (element) => ({
                              ...element,
                              cols: element.cols + 1,
                              cells: element.cells.map((row, index) => [...row, index === 0 ? `Header ${row.length + 1}` : 'Value']),
                            }))}
                          >
                            Add col
                          </button>
                          <button
                            className="mini-button"
                            onClick={() => patchElementsByIds([primarySelectedElement.id], (element) => ({
                              ...element,
                              cols: Math.max(1, element.cols - 1),
                              cells: element.cells.map((row) => row.slice(0, Math.max(1, element.cols - 1))),
                            }))}
                          >
                            Remove col
                          </button>
                        </div>
                        <label>
                          <span>Header fill</span>
                          <input value={primarySelectedElement.headerFill} onChange={(event) => patchElementsByIds([primarySelectedElement.id], (element) => ({ ...element, headerFill: event.target.value }))} />
                        </label>
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

          <div className="panel-block">
            <div className="block-header">
              <h2>Comments</h2>
              <span>{comments.length}</span>
            </div>
            <p className={`status-pill ${commentStatus.type}`}>{commentStatus.message || 'Idle'}</p>
            <label>
              <span>Filter</span>
              <select value={commentFilter} onChange={(event) => setCommentFilter(event.target.value)}>
                <option value="all">All</option>
                <option value="open">Open only</option>
                <option value="page">Current page</option>
                <option value="selection">Current selection</option>
                <option value="tag">By tag</option>
              </select>
            </label>
            {commentFilter === 'tag' ? (
              <label>
                <span>Tag</span>
                <input value={tagFilter} onChange={(event) => setTagFilter(event.target.value)} placeholder="tag name without #" />
              </label>
            ) : null}
            <label>
              <span>New comment</span>
              <textarea rows="4" value={commentDraft} onChange={(event) => handleCommentDraftChange(event.target.value)} placeholder="Comment on the page or selected element" />
            </label>
            {mentionSuggestions.length > 0 ? (
              <div className="mention-suggestions">
                {mentionSuggestions.map((name) => (
                  <button key={name} className="mini-button" onClick={() => handleInsertMention(name)}>@{name}</button>
                ))}
              </div>
            ) : null}
            <div className="row-actions">
              <button className="accent-button" onClick={handleAddComment} disabled={!currentUser || !documentData.id}>Add comment</button>
              <button className="ghost-button" onClick={refreshComments} disabled={!currentUser || !documentData.id}>Refresh</button>
            </div>
            <div className="page-list server-list">
              {filteredComments.filter((comment) => !comment.parentId).map((comment) => (
                <div key={comment.id} className={`page-card ${comment.resolved ? 'resolved-comment' : ''}`}>
                  <div className="library-open">
                    <strong>{comment.authorName}</strong>
                    <small>{new Date(comment.createdAt).toLocaleString()}</small>
                    <small>{comment.elementId ? `Element: ${comment.elementId}` : `Page: ${comment.pageId || currentPage.id}`}</small>
                    <small>{renderCommentText(comment.body)}</small>
                    <small>
                      {(() => {
                        const meta = extractCommentMeta(comment.body)
                        return [
                          meta.mentions.length > 0 ? `Mentions: ${meta.mentions.join(', ')}` : '',
                          meta.tags.length > 0 ? `Tags: ${meta.tags.join(', ')}` : '',
                        ].filter(Boolean).join(' | ') || 'No mentions or tags'
                      })()}
                    </small>
                    <small>{comment.resolved ? 'Resolved' : 'Open'}</small>
                  </div>
                  <div className="row-actions">
                    <button className="mini-button" onClick={() => handleResolveComment(comment.id, !comment.resolved)}>{comment.resolved ? 'Reopen' : 'Resolve'}</button>
                    <button className="mini-button danger" onClick={() => handleDeleteComment(comment.id)}>Delete</button>
                  </div>
                  <label>
                    <span>Reply</span>
                    <textarea rows="2" value={replyDrafts[comment.id] || ''} onChange={(event) => setReplyDrafts((prev) => ({ ...prev, [comment.id]: event.target.value }))} />
                  </label>
                  <button className="mini-button" onClick={() => handleReplyToComment(comment.id)}>Add reply</button>
                  {comments.filter((reply) => reply.parentId === comment.id).map((reply) => (
                    <div key={reply.id} className="comment-reply">
                      <strong>{reply.authorName}</strong>
                      <small>{new Date(reply.createdAt).toLocaleString()}</small>
                      <small>{renderCommentText(reply.body)}</small>
                    </div>
                  ))}
                </div>
              ))}
            </div>
          </div>
        </aside>
      </div>
    </div>
  )
}

export default App
