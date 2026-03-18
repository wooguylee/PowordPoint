import fs from 'node:fs/promises'
import path from 'node:path'
import { pool } from './db.js'

export const UPLOAD_ROOT = path.resolve('server', 'data', 'uploads')

export const ensureUploadRoot = async () => {
  await fs.mkdir(UPLOAD_ROOT, { recursive: true })
}

const collectUploadRefs = (value, refs = new Set()) => {
  if (typeof value === 'string' && value.startsWith('/uploads/')) {
    refs.add(value)
    return refs
  }

  if (Array.isArray(value)) {
    value.forEach((item) => collectUploadRefs(item, refs))
    return refs
  }

  if (value && typeof value === 'object') {
    Object.values(value).forEach((item) => collectUploadRefs(item, refs))
  }

  return refs
}

const ensureOwner = (row, userId) => {
  if (!row || row.owner_id !== userId) {
    throw new Error('Forbidden')
  }
}

export const createUser = async ({ id, email, name, passwordHash }) => {
  const { rows } = await pool.query(
    `
      INSERT INTO users (id, email, name, password_hash)
      VALUES ($1, $2, $3, $4)
      RETURNING id, email, name, created_at AS "createdAt"
    `,
    [id, email.toLowerCase(), name, passwordHash],
  )

  return rows[0]
}

export const findUserByEmail = async (email) => {
  const { rows } = await pool.query(
    `SELECT id, email, name, password_hash AS "passwordHash", created_at AS "createdAt" FROM users WHERE email = $1`,
    [email.toLowerCase()],
  )

  return rows[0] || null
}

export const findUserById = async (userId) => {
  const { rows } = await pool.query(
    `SELECT id, email, name, created_at AS "createdAt" FROM users WHERE id = $1`,
    [userId],
  )

  return rows[0] || null
}

export const listDocuments = async (userId, query = '') => {
  const normalized = `%${`${query}`.trim().toLowerCase()}%`
  const { rows } = await pool.query(
    `
      SELECT
        id,
        title,
        description,
        updated_at AS "updatedAt",
        jsonb_array_length(COALESCE(payload->'pages', '[]'::jsonb)) AS "pageCount"
      FROM documents
      WHERE owner_id = $1
        AND ($2 = '%%' OR lower(coalesce(title, '') || ' ' || coalesce(description, '')) LIKE $2)
      ORDER BY updated_at DESC
    `,
    [userId, normalized],
  )

  return rows
}

export const readDocument = async (userId, documentId) => {
  const { rows } = await pool.query(
    `SELECT owner_id, payload FROM documents WHERE id = $1`,
    [documentId],
  )

  if (rows.length === 0) {
    throw new Error('Document not found.')
  }

  ensureOwner(rows[0], userId)
  return rows[0].payload
}

export const writeDocument = async (userId, documentId, documentData) => {
  const client = await pool.connect()

  try {
    await client.query('BEGIN')

    const currentResult = await client.query(
      `SELECT owner_id, payload FROM documents WHERE id = $1`,
      [documentId],
    )

    if (currentResult.rows.length > 0) {
      ensureOwner(currentResult.rows[0], userId)
      await client.query(
        `
          INSERT INTO document_versions (document_id, owner_id, title, payload, created_at)
          VALUES ($1, $2, $3, $4::jsonb, NOW())
        `,
        [
          documentId,
          userId,
          currentResult.rows[0].payload.title || 'Untitled document',
          JSON.stringify(currentResult.rows[0].payload),
        ],
      )
    }

    await client.query(
      `
        INSERT INTO documents (id, owner_id, title, description, payload, created_at, updated_at)
        VALUES ($1, $2, $3, $4, $5::jsonb, $6, $7)
        ON CONFLICT (id)
        DO UPDATE SET
          title = EXCLUDED.title,
          description = EXCLUDED.description,
          payload = EXCLUDED.payload,
          updated_at = EXCLUDED.updated_at
      `,
      [
        documentId,
        userId,
        documentData.title || 'Untitled document',
        documentData.description || '',
        JSON.stringify(documentData),
        documentData.createdAt,
        documentData.updatedAt,
      ],
    )

    await client.query('COMMIT')
    return documentData
  } catch (error) {
    await client.query('ROLLBACK')
    throw error
  } finally {
    client.release()
  }
}

export const patchDocumentMeta = async (userId, documentId, patch) => {
  const current = await readDocument(userId, documentId)
  const nextDocument = {
    ...current,
    title: typeof patch.title === 'string' ? patch.title : current.title,
    description: typeof patch.description === 'string' ? patch.description : current.description,
    updatedAt: new Date().toISOString(),
  }

  await writeDocument(userId, documentId, nextDocument)
  return nextDocument
}

export const deleteDocument = async (userId, documentId) => {
  const { rows } = await pool.query(`SELECT owner_id FROM documents WHERE id = $1`, [documentId])

  if (rows.length === 0) {
    throw new Error('Document not found.')
  }

  ensureOwner(rows[0], userId)
  await pool.query(`DELETE FROM document_versions WHERE document_id = $1`, [documentId])
  await pool.query(`DELETE FROM documents WHERE id = $1`, [documentId])
}

export const listDocumentVersions = async (userId, documentId) => {
  const current = await pool.query(`SELECT owner_id FROM documents WHERE id = $1`, [documentId])
  if (current.rows.length === 0) {
    return []
  }
  ensureOwner(current.rows[0], userId)

  const { rows } = await pool.query(
    `
      SELECT id::text AS "versionId", title, created_at AS "createdAt", payload
      FROM document_versions
      WHERE document_id = $1 AND owner_id = $2
      ORDER BY created_at DESC
    `,
    [documentId, userId],
  )

  return rows.map((row, index) => ({
    versionId: row.versionId,
    createdAt: row.createdAt,
    title: row.title,
    summary: {
      pageCount: Array.isArray(row.payload?.pages) ? row.payload.pages.length : 0,
      elementCount: Array.isArray(row.payload?.pages)
        ? row.payload.pages.reduce((count, page) => count + (Array.isArray(page.elements) ? page.elements.length : 0), 0)
        : 0,
      index,
    },
  }))
}

export const readDocumentVersion = async (userId, documentId, versionId) => {
  const { rows } = await pool.query(
    `
      SELECT id::text AS "versionId", created_at AS "createdAt", title, payload, owner_id
      FROM document_versions
      WHERE document_id = $1 AND id::text = $2
    `,
    [documentId, versionId],
  )

  if (rows.length === 0) {
    throw new Error('Document version not found.')
  }

  ensureOwner(rows[0], userId)
  return rows[0]
}

export const restoreDocumentVersion = async (userId, documentId, versionId) => {
  const currentDocument = await readDocument(userId, documentId)
  const version = await readDocumentVersion(userId, documentId, versionId)
  const restored = {
    ...version.payload,
    id: documentId,
    updatedAt: new Date().toISOString(),
  }

  await writeDocument(userId, documentId, {
    ...currentDocument,
    ...restored,
  })

  return restored
}

export const listUploads = async (userId) => {
  await ensureUploadRoot()

  const usedUploads = new Map()
  const { rows: documents } = await pool.query(`SELECT payload FROM documents WHERE owner_id = $1`, [userId])

  documents.forEach((row) => {
    const refs = [...collectUploadRefs(row.payload)]
    refs.forEach((ref) => {
      usedUploads.set(ref, (usedUploads.get(ref) || 0) + 1)
    })
  })

  const { rows } = await pool.query(
    `
      SELECT file_name AS "name", original_name AS "originalName", url, size, mime_type AS "type", created_at AS "updatedAt"
      FROM uploads
      WHERE owner_id = $1
      ORDER BY created_at DESC
    `,
    [userId],
  )

  return rows.map((row) => ({
    ...row,
    usedBy: usedUploads.get(row.url) || 0,
  }))
}

export const createUpload = async (userId, upload) => {
  await pool.query(
    `
      INSERT INTO uploads (owner_id, file_name, original_name, url, size, mime_type)
      VALUES ($1, $2, $3, $4, $5, $6)
    `,
    [userId, upload.fileName, upload.originalName, upload.url, upload.size, upload.type],
  )
}

export const deleteUpload = async (userId, fileName) => {
  const { rows } = await pool.query(`SELECT owner_id FROM uploads WHERE file_name = $1`, [fileName])

  if (rows.length === 0) {
    throw new Error('Upload not found.')
  }

  ensureOwner(rows[0], userId)
  await pool.query(`DELETE FROM uploads WHERE file_name = $1`, [fileName])
  await ensureUploadRoot()
  await fs.unlink(path.join(UPLOAD_ROOT, fileName))
}

export const cleanupUnusedUploads = async (userId) => {
  const uploads = await listUploads(userId)
  const removed = []

  for (const upload of uploads) {
    if (upload.usedBy > 0) {
      continue
    }

    await deleteUpload(userId, upload.name)
    removed.push(upload.name)
  }

  return removed
}

export const getVersionDiff = async (userId, documentId, versionId) => {
  const version = await readDocumentVersion(userId, documentId, versionId)
  const current = await readDocument(userId, documentId)

  const versionPages = Array.isArray(version.payload?.pages) ? version.payload.pages : []
  const currentPages = Array.isArray(current?.pages) ? current.pages : []
  const versionElementCount = versionPages.reduce((count, page) => count + (Array.isArray(page.elements) ? page.elements.length : 0), 0)
  const currentElementCount = currentPages.reduce((count, page) => count + (Array.isArray(page.elements) ? page.elements.length : 0), 0)

  return {
    versionId,
    currentTitle: current.title,
    versionTitle: version.title,
    currentUpdatedAt: current.updatedAt,
    versionCreatedAt: version.createdAt,
    pageDelta: currentPages.length - versionPages.length,
    elementDelta: currentElementCount - versionElementCount,
    addedPageNames: currentPages.map((page) => page.name).filter((name) => !versionPages.some((page) => page.name === name)),
    removedPageNames: versionPages.map((page) => page.name).filter((name) => !currentPages.some((page) => page.name === name)),
    currentPages,
    versionPages,
  }
}

export const listComments = async (userId, documentId) => {
  const { rows: documentRows } = await pool.query(`SELECT owner_id FROM documents WHERE id = $1`, [documentId])

  if (documentRows.length === 0) {
    return []
  }

  ensureOwner(documentRows[0], userId)

  const { rows } = await pool.query(
    `
      SELECT
        id::text AS id,
        parent_id::text AS "parentId",
        page_id AS "pageId",
        element_id AS "elementId",
        body,
        resolved,
        author_name AS "authorName",
        created_at AS "createdAt"
      FROM comments
      WHERE document_id = $1
      ORDER BY created_at DESC
    `,
    [documentId],
  )

  return rows
}

export const createComment = async (userId, documentId, comment) => {
  const { rows: documentRows } = await pool.query(`SELECT owner_id FROM documents WHERE id = $1`, [documentId])

  if (documentRows.length === 0) {
    throw new Error('Document not found.')
  }

  ensureOwner(documentRows[0], userId)

  const user = await findUserById(userId)
  const { rows } = await pool.query(
    `
      INSERT INTO comments (document_id, owner_id, author_name, page_id, element_id, body)
      VALUES ($1, $2, $3, $4, $5, $6)
      RETURNING id::text AS id, parent_id::text AS "parentId", page_id AS "pageId", element_id AS "elementId", body, resolved, author_name AS "authorName", created_at AS "createdAt"
    `,
    [documentId, userId, user?.name || 'User', comment.pageId || null, comment.elementId || null, comment.body],
  )

  return rows[0]
}

export const replyToComment = async (userId, documentId, parentId, body) => {
  const { rows: documentRows } = await pool.query(`SELECT owner_id FROM documents WHERE id = $1`, [documentId])

  if (documentRows.length === 0) {
    throw new Error('Document not found.')
  }

  ensureOwner(documentRows[0], userId)
  const user = await findUserById(userId)

  const { rows } = await pool.query(
    `
      INSERT INTO comments (document_id, owner_id, author_name, parent_id, body)
      VALUES ($1, $2, $3, $4, $5)
      RETURNING id::text AS id, parent_id::text AS "parentId", page_id AS "pageId", element_id AS "elementId", body, resolved, author_name AS "authorName", created_at AS "createdAt"
    `,
    [documentId, userId, user?.name || 'User', Number(parentId), body],
  )

  return rows[0]
}

export const resolveComment = async (userId, commentId, resolved) => {
  const { rows } = await pool.query(
    `
      UPDATE comments
      SET resolved = $2
      WHERE id::text = $1 AND owner_id = $3
      RETURNING id::text AS id, parent_id::text AS "parentId", page_id AS "pageId", element_id AS "elementId", body, resolved, author_name AS "authorName", created_at AS "createdAt"
    `,
    [commentId, resolved, userId],
  )

  if (rows.length === 0) {
    throw new Error('Comment not found.')
  }

  return rows[0]
}

export const deleteComment = async (userId, commentId) => {
  const { rows } = await pool.query(`SELECT owner_id FROM comments WHERE id::text = $1`, [commentId])

  if (rows.length === 0) {
    throw new Error('Comment not found.')
  }

  ensureOwner(rows[0], userId)
  await pool.query(`DELETE FROM comments WHERE id::text = $1`, [commentId])
}

export const listTemplates = async () => {
  const { rows } = await pool.query(
    `SELECT id, title, description, payload FROM templates ORDER BY created_at ASC`,
  )

  return rows.map((row) => ({
    id: row.id,
    title: row.title,
    description: row.description,
    document: row.payload,
  }))
}

export const upsertTemplate = async (template) => {
  await pool.query(
    `
      INSERT INTO templates (id, title, description, payload)
      VALUES ($1, $2, $3, $4::jsonb)
      ON CONFLICT (id)
      DO UPDATE SET title = EXCLUDED.title, description = EXCLUDED.description, payload = EXCLUDED.payload
    `,
    [template.id, template.title, template.description, JSON.stringify(template.document)],
  )
}
