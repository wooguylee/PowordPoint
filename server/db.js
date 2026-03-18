import 'dotenv/config'
import { Pool } from 'pg'

const requiredEnv = [
  'PGHOST',
  'PGPORT',
  'PGUSER',
  'PGPASSWORD',
  'PGDATABASE',
  'JWT_SECRET',
]

const missing = requiredEnv.filter((key) => !process.env[key])

if (missing.length > 0) {
  throw new Error(`Missing PostgreSQL environment variables: ${missing.join(', ')}`)
}

export const pool = new Pool({
  host: process.env.PGHOST,
  port: Number(process.env.PGPORT),
  user: process.env.PGUSER,
  password: process.env.PGPASSWORD,
  database: process.env.PGDATABASE,
  ssl: process.env.PGSSL === 'true' ? { rejectUnauthorized: false } : false,
})

export const initializeDatabase = async () => {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS users (
      id TEXT PRIMARY KEY,
      email TEXT UNIQUE NOT NULL,
      name TEXT NOT NULL,
      password_hash TEXT NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `)

  await pool.query(`
    CREATE TABLE IF NOT EXISTS documents (
      id TEXT PRIMARY KEY,
      owner_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
      title TEXT NOT NULL,
      description TEXT,
      payload JSONB NOT NULL,
      created_at TIMESTAMPTZ NOT NULL,
      updated_at TIMESTAMPTZ NOT NULL
    );
  `)

  await pool.query(`
    CREATE TABLE IF NOT EXISTS document_versions (
      id BIGSERIAL PRIMARY KEY,
      document_id TEXT NOT NULL,
      owner_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
      title TEXT NOT NULL,
      payload JSONB NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `)

  await pool.query(`
    CREATE INDEX IF NOT EXISTS idx_document_versions_document_id
    ON document_versions (document_id, created_at DESC);
  `)

  await pool.query(`
    CREATE TABLE IF NOT EXISTS uploads (
      id BIGSERIAL PRIMARY KEY,
      owner_id TEXT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
      file_name TEXT UNIQUE NOT NULL,
      original_name TEXT NOT NULL,
      url TEXT UNIQUE NOT NULL,
      size BIGINT NOT NULL,
      mime_type TEXT NOT NULL,
      created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
    );
  `)

  await pool.query(`
    CREATE INDEX IF NOT EXISTS idx_uploads_owner_id
    ON uploads (owner_id, created_at DESC);
  `)
}
