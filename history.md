# PowordPointer Development History

## Phase 1 - Initial Editor Foundation
- Scaffolded a React + Vite application using JavaScript only.
- Built a canvas-based page document editor with `react-konva`.
- Added page management, text, rectangle, ellipse, arrow, pen, image, and table elements.
- Implemented local JSON save/load and PNG export.
- Added OpenAI-compatible LLM integration UI and prompt pipeline.

## Phase 2 - Backend Introduction
- Added a Koa backend for document persistence and LLM proxying.
- Implemented backend JSON document save/load endpoints.
- Added server-backed document library and OpenAI-compatible proxy route.
- Added PDF export support and improved inline text editing UX.

## Phase 3 - Editing Enhancements
- Added multi-selection, drag selection, grouping, ungrouping, align tools, and snap guides.
- Implemented local image import and page-based canvas editing improvements.
- Added page duplication/removal and richer inspector controls.

## Phase 4 - Storage and Export Upgrades
- Added server-side image upload support and upload browsing.
- Added document search, rename, delete, and version history endpoints.
- Added version restore and diff preview support.
- Added improved PDF export with image handling, colors, rotation, and Korean font embedding.

## Phase 5 - PostgreSQL Migration and Auth
- Migrated document and version persistence from JSON files to PostgreSQL.
- Added `.env`-driven PostgreSQL and JWT configuration.
- Added authentication with register/login/me endpoints using JWT.
- Added ownership checks for documents, uploads, versions, and comments.
- Moved upload metadata into PostgreSQL while keeping file binaries on disk.

## Phase 6 - Productivity Features
- Added undo/redo with document history tracking.
- Added layer panel with lock/hide controls and reorder actions.
- Added server-backed template system and seeded default templates.
- Added richer table editing with row/column insert/delete and header styling.
- Added comment system for page/element annotations.

## Phase 7 - Workflow and Review Features
- Extended comments with reply threads and resolved state.
- Added visual version diff preview cards.
- Added upload thumbnails and richer asset browsing.
- Added group-aware layer display and drag-and-drop reordering.
- Added autosave to server with local recovery draft support.
- Added comment filtering for all/open/current page/current selection.
- Added template manager UI with save/delete and preview summary.
- Added mention/tag parsing and highlighting inside comments.

## Phase 8 - Recovery and Advanced Review UX
- Added recovery history list for multiple autosaved drafts and point-in-time restore.
- Added a recovery banner to restore the latest draft or dismiss it.
- Upgraded template previews from text summary to a lightweight visual thumbnail renderer.
- Added comment mention suggestions while typing `@` names.
- Added tag-based comment filtering and explicit extracted mention/tag metadata display.
- Improved layer panel with collapsible group sections and persisted expand/collapse state.

## Current Architecture Summary
- Frontend: React + Vite + Konva canvas editor.
- Backend: Koa + PostgreSQL + JWT auth.
- Persistence: PostgreSQL for users, documents, versions, uploads metadata, templates, comments.
- File storage: local disk for uploaded files in `server/data/uploads`.
- Export: JSON, PNG, PDF.
- AI: OpenAI-compatible LLM proxy through backend.

## Key Files
- `src/App.jsx`
- `src/App.css`
- `src/lib/api.js`
- `src/lib/editor.js`
- `src/lib/llm.js`
- `src/lib/pdf.js`
- `server/index.js`
- `server/db.js`
- `server/storage.js`
- `.env.example`

## Validation Performed
- Repeatedly ran `npm run lint` successfully after feature milestones.
- Repeatedly ran `npm run build` successfully after feature milestones.
- Validated environment-based backend config loading with Node import checks.

## Notes
- Realtime collaborative editing was assessed but intentionally not implemented.
- Current autosave stores a recovery draft locally and syncs to the authenticated server document.
- Bundle size warnings remain for the frontend build, but builds succeed.
