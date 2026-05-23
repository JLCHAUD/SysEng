import { getDb } from './database';
import { Category } from '../types';

function rowToCategory(row: any): Category {
  return { id: row.id, name: row.name, color: row.color, createdAt: row.created_at };
}

export async function listCategories(): Promise<Category[]> {
  const db = await getDb();
  const rows = await db.getAllAsync('SELECT * FROM categories ORDER BY name ASC');
  return (rows as any[]).map(rowToCategory);
}

export async function insertCategory(cat: Category): Promise<void> {
  const db = await getDb();
  await db.runAsync(
    'INSERT INTO categories (id, name, color, created_at) VALUES (?, ?, ?, ?)',
    [cat.id, cat.name, cat.color, cat.createdAt]
  );
}

export async function deleteCategory(id: string): Promise<void> {
  const db = await getDb();
  await db.runAsync('DELETE FROM categories WHERE id = ?', [id]);
}
