const tokenKey = 'powordpointer.auth.token'

export const setAuthToken = (token) => {
  if (token) {
    window.localStorage.setItem(tokenKey, token)
  } else {
    window.localStorage.removeItem(tokenKey)
  }
}

export const getAuthToken = () => window.localStorage.getItem(tokenKey) || ''

const request = async (path, options = {}) => {
  const headers = { ...(options.headers || {}) }
  const token = getAuthToken()

  if (token) {
    headers.Authorization = `Bearer ${token}`
  }

  if (!(options.body instanceof FormData)) {
    headers['Content-Type'] = 'application/json'
  }

  const response = await fetch(path, {
    headers,
    ...options,
  })

  const isJson = response.headers.get('content-type')?.includes('application/json')
  const payload = isJson ? await response.json() : null

  if (!response.ok) {
    throw new Error(payload?.error || 'Request failed.')
  }

  return payload
}

export const fetchDocumentLibrary = async (query = '') => {
  const suffix = query ? `?q=${encodeURIComponent(query)}` : ''
  const payload = await request(`/api/documents${suffix}`)
  return Array.isArray(payload.documents) ? payload.documents : []
}

export const registerUser = async (payload) => {
  return request('/api/auth/register', {
    method: 'POST',
    body: JSON.stringify(payload),
  })
}

export const loginUser = async (payload) => {
  return request('/api/auth/login', {
    method: 'POST',
    body: JSON.stringify(payload),
  })
}

export const fetchCurrentUser = async () => {
  const payload = await request('/api/auth/me')
  return payload.user
}

export const fetchDocumentById = async (documentId) => {
  const payload = await request(`/api/documents/${documentId}`)
  return payload.document
}

export const saveDocumentToServer = async (documentData) => {
  const method = documentData.id ? 'PUT' : 'POST'
  const path = documentData.id ? `/api/documents/${documentData.id}` : '/api/documents'
  const payload = await request(path, {
    method,
    body: JSON.stringify({ document: documentData }),
  })

  return payload.document
}

export const renameDocumentOnServer = async (documentId, patch) => {
  const payload = await request(`/api/documents/${documentId}`, {
    method: 'PATCH',
    body: JSON.stringify(patch),
  })

  return payload.document
}

export const deleteDocumentFromServer = async (documentId) => {
  await request(`/api/documents/${documentId}`, {
    method: 'DELETE',
  })
}

export const fetchDocumentVersions = async (documentId) => {
  const payload = await request(`/api/documents/${documentId}/versions`)
  return Array.isArray(payload.versions) ? payload.versions : []
}

export const restoreDocumentVersion = async (documentId, versionId) => {
  const payload = await request(`/api/documents/${documentId}/restore/${versionId}`, {
    method: 'POST',
  })

  return payload.document
}

export const fetchVersionDiff = async (documentId, versionId) => {
  const payload = await request(`/api/documents/${documentId}/versions/${versionId}/diff`)
  return payload.diff
}

export const uploadImageToServer = async (file) => {
  const body = new FormData()
  body.append('image', file)

  const payload = await request('/api/uploads/image', {
    method: 'POST',
    body,
  })

  return payload.image
}

export const fetchUploads = async () => {
  const payload = await request('/api/uploads')
  return Array.isArray(payload.uploads) ? payload.uploads : []
}

export const deleteUploadFromServer = async (fileName) => {
  await request(`/api/uploads/${encodeURIComponent(fileName)}`, {
    method: 'DELETE',
  })
}

export const cleanupUploadsOnServer = async () => {
  const payload = await request('/api/uploads/cleanup', {
    method: 'POST',
  })

  return Array.isArray(payload.removed) ? payload.removed : []
}

export const requestLlmProxy = async (payload) => {
  return request('/api/llm', {
    method: 'POST',
    body: JSON.stringify(payload),
  })
}
