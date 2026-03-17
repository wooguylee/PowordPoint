import fs from 'node:fs/promises'
import path from 'node:path'

const DATA_ROOT = path.resolve('server', 'data', 'documents')
export const UPLOAD_ROOT = path.resolve('server', 'data', 'uploads')
const VERSION_ROOT = path.resolve('server', 'data', 'versions')

const ensureDataRoot = async () => {
  await fs.mkdir(DATA_ROOT, { recursive: true })
}

export const ensureUploadRoot = async () => {
  await fs.mkdir(UPLOAD_ROOT, { recursive: true })
}

const ensureVersionRoot = async () => {
  await fs.mkdir(VERSION_ROOT, { recursive: true })
}

const filePathFor = (documentId) => path.join(DATA_ROOT, `${documentId}.json`)
const versionDirFor = (documentId) => path.join(VERSION_ROOT, documentId)

const tryReadCurrentDocument = async (documentId) => {
  try {
    return await readDocument(documentId)
  } catch {
    return null
  }
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

const createVersionSnapshot = async (documentId, documentData) => {
  await ensureVersionRoot()
  await fs.mkdir(versionDirFor(documentId), { recursive: true })

  const versionId = `${Date.now()}`
  const versionData = {
    versionId,
    documentId,
    createdAt: new Date().toISOString(),
    document: documentData,
  }

  await fs.writeFile(
    path.join(versionDirFor(documentId), `${versionId}.json`),
    JSON.stringify(versionData, null, 2),
    'utf8',
  )
}

export const listDocuments = async (query = '') => {
  await ensureDataRoot()

  const entries = await fs.readdir(DATA_ROOT, { withFileTypes: true })
  const items = []
  const normalizedQuery = `${query}`.trim().toLowerCase()

  for (const entry of entries) {
    if (!entry.isFile() || !entry.name.endsWith('.json')) {
      continue
    }

    const raw = await fs.readFile(path.join(DATA_ROOT, entry.name), 'utf8')
    const documentData = JSON.parse(raw)

    const summary = {
      id: documentData.id,
      title: documentData.title,
      description: documentData.description,
      updatedAt: documentData.updatedAt,
      pageCount: Array.isArray(documentData.pages) ? documentData.pages.length : 0,
    }

    if (
      normalizedQuery &&
      !`${summary.title || ''} ${summary.description || ''}`.toLowerCase().includes(normalizedQuery)
    ) {
      continue
    }

    items.push(summary)
  }

  return items.sort((left, right) => `${right.updatedAt}`.localeCompare(`${left.updatedAt}`))
}

export const readDocument = async (documentId) => {
  await ensureDataRoot()
  const raw = await fs.readFile(filePathFor(documentId), 'utf8')
  return JSON.parse(raw)
}

export const writeDocument = async (documentId, documentData) => {
  await ensureDataRoot()
  const currentDocument = await tryReadCurrentDocument(documentId)

  if (currentDocument) {
    await createVersionSnapshot(documentId, currentDocument)
  }

  await fs.writeFile(filePathFor(documentId), JSON.stringify(documentData, null, 2), 'utf8')
  return documentData
}

export const patchDocumentMeta = async (documentId, patch) => {
  const current = await readDocument(documentId)
  const nextDocument = {
    ...current,
    title: typeof patch.title === 'string' ? patch.title : current.title,
    description: typeof patch.description === 'string' ? patch.description : current.description,
    updatedAt: new Date().toISOString(),
  }

  await writeDocument(documentId, nextDocument)
  return nextDocument
}

export const deleteDocument = async (documentId) => {
  await ensureDataRoot()
  await fs.unlink(filePathFor(documentId))

  try {
    await fs.rm(versionDirFor(documentId), { recursive: true, force: true })
  } catch {
    return undefined
  }
}

export const listDocumentVersions = async (documentId) => {
  await ensureVersionRoot()

  try {
    const entries = await fs.readdir(versionDirFor(documentId), { withFileTypes: true })
    const versions = []

    for (const entry of entries) {
      if (!entry.isFile() || !entry.name.endsWith('.json')) {
        continue
      }

      const raw = await fs.readFile(path.join(versionDirFor(documentId), entry.name), 'utf8')
      const payload = JSON.parse(raw)
      versions.push({
        versionId: payload.versionId,
        createdAt: payload.createdAt,
        title: payload.document?.title || 'Untitled document',
      })
    }

    return versions.sort((left, right) => `${right.createdAt}`.localeCompare(`${left.createdAt}`))
  } catch {
    return []
  }
}

export const restoreDocumentVersion = async (documentId, versionId) => {
  const currentDocument = await readDocument(documentId)
  await createVersionSnapshot(documentId, currentDocument)

  const raw = await fs.readFile(path.join(versionDirFor(documentId), `${versionId}.json`), 'utf8')
  const payload = JSON.parse(raw)
  const restored = {
    ...payload.document,
    id: documentId,
    updatedAt: new Date().toISOString(),
  }

  await fs.writeFile(filePathFor(documentId), JSON.stringify(restored, null, 2), 'utf8')
  return restored
}

export const listUploads = async () => {
  await ensureUploadRoot()
  await ensureDataRoot()

  const usedUploads = new Map()
  const documentEntries = await fs.readdir(DATA_ROOT, { withFileTypes: true })

  for (const entry of documentEntries) {
    if (!entry.isFile() || !entry.name.endsWith('.json')) {
      continue
    }

    const raw = await fs.readFile(path.join(DATA_ROOT, entry.name), 'utf8')
    const documentData = JSON.parse(raw)
    const refs = [...collectUploadRefs(documentData)]

    refs.forEach((ref) => {
      usedUploads.set(ref, (usedUploads.get(ref) || 0) + 1)
    })
  }

  const uploads = await fs.readdir(UPLOAD_ROOT, { withFileTypes: true })
  const items = []

  for (const entry of uploads) {
    if (!entry.isFile()) {
      continue
    }

    const filePath = path.join(UPLOAD_ROOT, entry.name)
    const stat = await fs.stat(filePath)
    const url = `/uploads/${entry.name}`
    items.push({
      name: entry.name,
      url,
      size: stat.size,
      updatedAt: stat.mtime.toISOString(),
      usedBy: usedUploads.get(url) || 0,
    })
  }

  return items.sort((left, right) => `${right.updatedAt}`.localeCompare(`${left.updatedAt}`))
}

export const deleteUpload = async (fileName) => {
  await ensureUploadRoot()
  await fs.unlink(path.join(UPLOAD_ROOT, fileName))
}

export const cleanupUnusedUploads = async () => {
  const uploads = await listUploads()
  const removed = []

  for (const upload of uploads) {
    if (upload.usedBy > 0) {
      continue
    }

    await deleteUpload(upload.name)
    removed.push(upload.name)
  }

  return removed
}
