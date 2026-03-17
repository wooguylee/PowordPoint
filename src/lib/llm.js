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

export const extractJson = (rawText) => {
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

const schemaGuide = `
Allowed element types: text, rect, ellipse, arrow, pen, table.

Text element:
{"type":"text","x":80,"y":120,"width":420,"height":140,"text":"Heading","fontSize":32,"fontFamily":"Avenir Next, Segoe UI Variable, Noto Sans KR, sans-serif","fill":"#142b36","align":"left","lineHeight":1.2,"fontStyle":"normal"}

Rect or ellipse:
{"type":"rect","x":80,"y":120,"width":240,"height":140,"fill":"#c7e7ef","stroke":"#23404d","strokeWidth":2,"opacity":1}

Arrow:
{"type":"arrow","x":80,"y":120,"width":280,"height":0,"stroke":"#d56f3e","strokeWidth":4,"pointerLength":18,"pointerWidth":18,"opacity":1}

Pen:
{"type":"pen","x":80,"y":120,"points":[0,0,160,20,220,90],"stroke":"#142b36","strokeWidth":3,"opacity":1}

Table:
{"type":"table","x":80,"y":120,"width":520,"height":240,"rows":4,"cols":3,"stroke":"#23404d","strokeWidth":1,"fill":"#fffdf8","headerFill":"#ddecef","textColor":"#17323d","fontSize":18,"cells":[["H1","H2","H3"],["A","B","C"],["D","E","F"],["G","H","I"]]}
`

export const buildPrompts = ({ mode, prompt, documentData, currentPage, selection }) => {
  const documentSummary = {
    title: documentData.title,
    description: documentData.description,
    pageCount: documentData.pages.length,
    currentPage: {
      name: currentPage.name,
      width: currentPage.width,
      height: currentPage.height,
      background: currentPage.background,
      elements: currentPage.elements.map((element) => ({
        id: element.id,
        type: element.type,
        x: element.x,
        y: element.y,
        width: element.width,
        height: element.height,
        text: element.text,
      })),
    },
    selection,
  }

  const sharedSystem = [
    'You are generating layout JSON for a page-based canvas editor named PowordPointer.',
    'Return strict JSON only. Do not include markdown, prose, or comments.',
    'Keep all objects inside the current page size unless the task is page generation.',
    schemaGuide,
  ].join('\n\n')

  if (mode === 'draft') {
    return {
      system: `${sharedSystem}\n\nReturn {"elements": [...]} with polished text blocks and optional supporting shapes.`,
      user: `Create document draft content for this prompt:\n${prompt}\n\nEditor context:\n${JSON.stringify(documentSummary, null, 2)}`,
    }
  }

  if (mode === 'layout') {
    return {
      system: `${sharedSystem}\n\nReturn {"elements": [...]} focused on diagrams, callouts, panels, arrows, and supporting text.`,
      user: `Create a layout composition for this prompt:\n${prompt}\n\nEditor context:\n${JSON.stringify(documentSummary, null, 2)}`,
    }
  }

  if (mode === 'pages') {
    return {
      system: `${sharedSystem}\n\nReturn {"pages": [{"name":"Page name","background":"#fffaf2","elements":[...]}]}. Generate 2-5 pages unless the prompt asks otherwise.`,
      user: `Generate multiple document pages for this prompt:\n${prompt}\n\nEditor context:\n${JSON.stringify(documentSummary, null, 2)}`,
    }
  }

  return {
    system: `${sharedSystem}\n\nReturn {"elements": [...]} where each element corresponds to one selected element in the same order, preserving intent while applying requested edits.`,
    user: `Edit the selected elements according to this prompt:\n${prompt}\n\nSelected elements:\n${JSON.stringify(selection, null, 2)}`,
  }
}

export const normalizeLlmPayload = (payload) => {
  const message = payload?.text || getTextFromContent(payload?.choices?.[0]?.message?.content)

  if (!message) {
    throw new Error('LLM response did not include message content.')
  }

  return {
    text: message,
    json: payload?.json || extractJson(message),
  }
}
