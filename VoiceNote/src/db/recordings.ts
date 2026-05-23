import { getDb } from './database';
import { Recording } from '../types';

function rowToRecording(row: any): Recording {
  return {
    id: row.id,
    title: row.title,
    audioUri: row.audio_uri,
    duration: row.duration,
    transcript: row.transcript,
    categoryId: row.category_id,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
  };
}

export async function listRecordings(categoryId?: string): Promise<Recording[]> {
  const db = await getDb();
  const rows = categoryId
    ? await db.getAllAsync('SELECT * FROM recordings WHERE category_id = ? ORDER BY created_at DESC', [categoryId])
    : await db.getAllAsync('SELECT * FROM recordings ORDER BY created_at DESC');
  return (rows as any[]).map(rowToRecording);
}

export async function getRecording(id: string): Promise<Recording | null> {
  const db = await getDb();
  const row = await db.getFirstAsync('SELECT * FROM recordings WHERE id = ?', [id]);
  return row ? rowToRecording(row) : null;
}

export async function insertRecording(rec: Recording): Promise<void> {
  const db = await getDb();
  await db.runAsync(
    `INSERT INTO recordings (id, title, audio_uri, duration, transcript, category_id, created_at, updated_at)
     VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
    [rec.id, rec.title, rec.audioUri, rec.duration, rec.transcript, rec.categoryId, rec.createdAt, rec.updatedAt]
  );
}

export async function updateRecording(id: string, patch: Partial<Pick<Recording, 'title' | 'transcript' | 'categoryId'>>): Promise<void> {
  const db = await getDb();
  const updatedAt = new Date().toISOString();
  const fields: string[] = [];
  const values: any[] = [];

  if (patch.title !== undefined) { fields.push('title = ?'); values.push(patch.title); }
  if (patch.transcript !== undefined) { fields.push('transcript = ?'); values.push(patch.transcript); }
  if (patch.categoryId !== undefined) { fields.push('category_id = ?'); values.push(patch.categoryId); }

  fields.push('updated_at = ?');
  values.push(updatedAt);
  values.push(id);

  await db.runAsync(`UPDATE recordings SET ${fields.join(', ')} WHERE id = ?`, values);
}

export async function deleteRecording(id: string): Promise<void> {
  const db = await getDb();
  await db.runAsync('DELETE FROM recordings WHERE id = ?', [id]);
}
