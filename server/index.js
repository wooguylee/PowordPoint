import fs from 'node:fs/promises'
import path from 'node:path'
import { randomUUID } from 'node:crypto'
import bcrypt from 'bcryptjs'
import jwt from 'jsonwebtoken'
import Koa from 'koa'
import Router from '@koa/router'
import multer from '@koa/multer'
import bodyParser from 'koa-bodyparser'
import serve from 'koa-static'
import { initializeDatabase } from './db.js'
import {
  UPLOAD_ROOT,
  createUpload,
  createUser,
  cleanupUnusedUploads,
  createComment,
  listTemplates,
  replyToComment,
  resolveComment,
  upsertTemplate,
  deleteDocument,
  deleteComment,
  deleteUpload,
  ensureUploadRoot,
  findUserByEmail,
  findUserById,
  getVersionDiff,
  listDocuments,
  listDocumentVersions,
  listComments,
  listUploads,
  patchDocumentMeta,
  readDocument,
  restoreDocumentVersion,
  writeDocument,
} from './storage.js'

const app = new Koa()
const router = new Router({ prefix: '/api' })
const PORT = Number(process.env.PORT || 8787)
const DIST_ROOT = path.resolve('dist')
const uploadStatic = serve(UPLOAD_ROOT)
const JWT_SECRET = process.env.JWT_SECRET

await initializeDatabase()
await ensureUploadRoot()

const upload = multer({
  storage: multer.diskStorage({
    destination: (_request, _file, callback) => callback(null, UPLOAD_ROOT),
    filename: (_request, file, callback) => {
      const extension = path.extname(file.originalname) || '.bin'
      callback(null, `${Date.now()}-${randomUUID()}${extension}`)
    },
  }),
})

const sendError = (ctx, status, message) => {
  ctx.status = status
  ctx.body = { error: message }
}

const issueToken = (user) =>
  jwt.sign(
    { sub: user.id, email: user.email, name: user.name },
    JWT_SECRET,
    { expiresIn: '7d' },
  )

const auth = async (ctx, next) => {
  const header = ctx.headers.authorization || ''
  const token = header.startsWith('Bearer ') ? header.slice(7) : ''

  if (!token) {
    sendError(ctx, 401, 'Authentication required.')
    return
  }

  try {
    const payload = jwt.verify(token, JWT_SECRET)
    const user = await findUserById(payload.sub)

    if (!user) {
      sendError(ctx, 401, 'Invalid token.')
      return
    }

    ctx.state.user = user
    await next()
  } catch {
    sendError(ctx, 401, 'Invalid token.')
  }
}

const extractJson = (rawText) => {
  const text = `${rawText || ''}`.trim()

  if (!text) {
    throw new Error('LLM returned an empty response.')
  }

  const fenced = text.match(/```(?:json)?\s*([\s\S]*?)```/i)
  const candidate = fenced?.[1]?.trim() || text
  const start = candidate.indexOf('{')
  const end = candidate.lastIndexOf('}')

  if (start === -1 || end === -1) {
    throw new Error('Could not find JSON in the LLM response.')
  }

  return JSON.parse(candidate.slice(start, end + 1))
}

const getTextFromContent = (content) => {
  if (typeof content === 'string') {
    return content
  }

  if (Array.isArray(content)) {
    return content
      .map((part) => {
        if (typeof part === 'string') {
          return part
        }

        if (part && typeof part === 'object' && typeof part.text === 'string') {
          return part.text
        }

        return ''
      })
      .join('')
  }

  return ''
}

app.use(bodyParser({ jsonLimit: '30mb' }))

router.get('/health', (ctx) => {
  ctx.body = { ok: true }
})

router.post('/auth/register', async (ctx) => {
  const { email, name, password } = ctx.request.body || {}

  if (!email || !name || !password) {
    sendError(ctx, 400, 'email, name, and password are required.')
    return
  }

  const existing = await findUserByEmail(email)
  if (existing) {
    sendError(ctx, 409, 'Email already exists.')
    return
  }

  const user = await createUser({
    id: randomUUID(),
    email,
    name,
    passwordHash: await bcrypt.hash(password, 10),
  })

  ctx.status = 201
  ctx.body = { user, token: issueToken(user) }
})

router.post('/auth/login', async (ctx) => {
  const { email, password } = ctx.request.body || {}

  if (!email || !password) {
    sendError(ctx, 400, 'email and password are required.')
    return
  }

  const user = await findUserByEmail(email)
  if (!user || !(await bcrypt.compare(password, user.passwordHash))) {
    sendError(ctx, 401, 'Invalid credentials.')
    return
  }

  ctx.body = {
    user: { id: user.id, email: user.email, name: user.name, createdAt: user.createdAt },
    token: issueToken(user),
  }
})

router.get('/auth/me', auth, async (ctx) => {
  ctx.body = { user: ctx.state.user }
})

router.get('/documents', auth, async (ctx) => {
  ctx.body = { documents: await listDocuments(ctx.state.user.id, ctx.query.q || '') }
})

router.get('/documents/:id', auth, async (ctx) => {
  try {
    ctx.body = { document: await readDocument(ctx.state.user.id, ctx.params.id) }
  } catch {
    sendError(ctx, 404, 'Document not found.')
  }
})

router.post('/documents', auth, async (ctx) => {
  const documentData = ctx.request.body?.document

  if (!documentData || typeof documentData !== 'object') {
    sendError(ctx, 400, 'A document payload is required.')
    return
  }

  const documentId = documentData.id || randomUUID()
  const nextDocument = {
    ...documentData,
    id: documentId,
    updatedAt: new Date().toISOString(),
  }

  if (!nextDocument.createdAt) {
    nextDocument.createdAt = nextDocument.updatedAt
  }

  await writeDocument(ctx.state.user.id, documentId, nextDocument)
  ctx.status = 201
  ctx.body = { document: nextDocument }
})

router.put('/documents/:id', auth, async (ctx) => {
  const documentData = ctx.request.body?.document

  if (!documentData || typeof documentData !== 'object') {
    sendError(ctx, 400, 'A document payload is required.')
    return
  }

  const nextDocument = {
    ...documentData,
    id: ctx.params.id,
    updatedAt: new Date().toISOString(),
  }

  if (!nextDocument.createdAt) {
    nextDocument.createdAt = nextDocument.updatedAt
  }

  await writeDocument(ctx.state.user.id, ctx.params.id, nextDocument)
  ctx.body = { document: nextDocument }
})

router.patch('/documents/:id', auth, async (ctx) => {
  const patch = ctx.request.body || {}

  if (typeof patch !== 'object') {
    sendError(ctx, 400, 'A metadata patch payload is required.')
    return
  }

  try {
    ctx.body = { document: await patchDocumentMeta(ctx.state.user.id, ctx.params.id, patch) }
  } catch {
    sendError(ctx, 404, 'Document not found.')
  }
})

router.delete('/documents/:id', auth, async (ctx) => {
  try {
    await deleteDocument(ctx.state.user.id, ctx.params.id)
    ctx.status = 204
  } catch {
    sendError(ctx, 404, 'Document not found.')
  }
})

router.get('/documents/:id/versions', auth, async (ctx) => {
  ctx.body = { versions: await listDocumentVersions(ctx.state.user.id, ctx.params.id) }
})

router.get('/documents/:id/comments', auth, async (ctx) => {
  ctx.body = { comments: await listComments(ctx.state.user.id, ctx.params.id) }
})

router.post('/documents/:id/comments', auth, async (ctx) => {
  const payload = ctx.request.body || {}

  if (!payload.body) {
    sendError(ctx, 400, 'Comment body is required.')
    return
  }

  ctx.status = 201
  ctx.body = {
    comment: await createComment(ctx.state.user.id, ctx.params.id, payload),
  }
})

router.post('/documents/:id/comments/:commentId/replies', auth, async (ctx) => {
  const payload = ctx.request.body || {}

  if (!payload.body) {
    sendError(ctx, 400, 'Reply body is required.')
    return
  }

  ctx.status = 201
  ctx.body = {
    comment: await replyToComment(ctx.state.user.id, ctx.params.id, ctx.params.commentId, payload.body),
  }
})

router.patch('/comments/:id', auth, async (ctx) => {
  try {
    ctx.body = {
      comment: await resolveComment(ctx.state.user.id, ctx.params.id, Boolean(ctx.request.body?.resolved)),
    }
  } catch {
    sendError(ctx, 404, 'Comment not found.')
  }
})

router.delete('/comments/:id', auth, async (ctx) => {
  try {
    await deleteComment(ctx.state.user.id, ctx.params.id)
    ctx.status = 204
  } catch {
    sendError(ctx, 404, 'Comment not found.')
  }
})

router.get('/documents/:id/versions/:versionId/diff', auth, async (ctx) => {
  try {
    ctx.body = { diff: await getVersionDiff(ctx.state.user.id, ctx.params.id, ctx.params.versionId) }
  } catch {
    sendError(ctx, 404, 'Document version not found.')
  }
})

router.get('/templates', auth, async (ctx) => {
  ctx.body = { templates: await listTemplates() }
})

router.post('/templates/seed', auth, async (ctx) => {
  const templates = Array.isArray(ctx.request.body?.templates) ? ctx.request.body.templates : []

  for (const template of templates) {
    await upsertTemplate(template)
  }

  ctx.status = 204
})

router.post('/documents/:id/restore/:versionId', auth, async (ctx) => {
  try {
    ctx.body = {
      document: await restoreDocumentVersion(ctx.state.user.id, ctx.params.id, ctx.params.versionId),
    }
  } catch {
    sendError(ctx, 404, 'Document version not found.')
  }
})

router.get('/uploads', auth, async (ctx) => {
  ctx.body = { uploads: await listUploads(ctx.state.user.id) }
})

router.delete('/uploads/:fileName', auth, async (ctx) => {
  try {
    await deleteUpload(ctx.state.user.id, ctx.params.fileName)
    ctx.status = 204
  } catch {
    sendError(ctx, 404, 'Upload not found.')
  }
})

router.post('/uploads/cleanup', auth, async (ctx) => {
  ctx.body = { removed: await cleanupUnusedUploads(ctx.state.user.id) }
})

router.post('/uploads/image', auth, upload.single('image'), async (ctx) => {
  if (!ctx.file) {
    sendError(ctx, 400, 'An image file is required.')
    return
  }

  ctx.status = 201
  await createUpload(ctx.state.user.id, {
    fileName: ctx.file.filename,
    originalName: ctx.file.originalname,
    url: `/uploads/${ctx.file.filename}`,
    size: ctx.file.size,
    type: ctx.file.mimetype,
  })
  ctx.body = {
    image: {
      name: ctx.file.originalname,
      url: `/uploads/${ctx.file.filename}`,
      size: ctx.file.size,
      type: ctx.file.mimetype,
    },
  }
})

router.post('/llm', auth, async (ctx) => {
  const { apiKey, baseUrl, model, system, user } = ctx.request.body || {}

  if (!apiKey || !baseUrl || !model || !system || !user) {
    sendError(ctx, 400, 'apiKey, baseUrl, model, system, and user are required.')
    return
  }

  const response = await fetch(`${`${baseUrl}`.replace(/\/$/, '')}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model,
      temperature: 0.7,
      messages: [
        { role: 'system', content: system },
        { role: 'user', content: user },
      ],
    }),
  })

  if (!response.ok) {
    sendError(ctx, response.status, (await response.text()) || 'LLM request failed.')
    return
  }

  const payload = await response.json()
  const message = payload?.choices?.[0]?.message?.content
  const text = getTextFromContent(message)

  if (!text) {
    sendError(ctx, 502, 'LLM response did not include message content.')
    return
  }

  ctx.body = {
    text,
    json: extractJson(text),
  }
})

app.use(async (ctx, next) => {
  try {
    await next()
  } catch (error) {
    sendError(ctx, error.status || 500, error.message || 'Internal server error.')
  }
})

app.use(router.routes())
app.use(router.allowedMethods())

app.use(async (ctx, next) => {
  if (!ctx.path.startsWith('/uploads/')) {
    await next()
    return
  }

  ctx.path = ctx.path.replace(/^\/uploads/, '') || '/'
  await uploadStatic(ctx, next)
})

app.use(serve(DIST_ROOT))

app.use(async (ctx) => {
  const indexPath = path.join(DIST_ROOT, 'index.html')

  try {
    ctx.type = 'html'
    ctx.body = await fs.readFile(indexPath, 'utf8')
  } catch {
    ctx.status = 404
    ctx.body = 'Build the client first or run Vite dev server.'
  }
})

app.listen(PORT, () => {
  console.log(`PowordPointer server running on http://localhost:${PORT}`)
})
