import { create } from 'zustand';
import { Recording, Category } from '../types';
import * as RecordingsDB from '../db/recordings';
import * as CategoriesDB from '../db/categories';

interface AppState {
  recordings: Recording[];
  categories: Category[];
  loadRecordings: (categoryId?: string) => Promise<void>;
  loadCategories: () => Promise<void>;
  addRecording: (rec: Recording) => Promise<void>;
  updateRecording: (id: string, patch: Partial<Pick<Recording, 'title' | 'transcript' | 'categoryId'>>) => Promise<void>;
  deleteRecording: (id: string) => Promise<void>;
  addCategory: (cat: Category) => Promise<void>;
  deleteCategory: (id: string) => Promise<void>;
}

export const useStore = create<AppState>((set, get) => ({
  recordings: [],
  categories: [],

  loadRecordings: async (categoryId?) => {
    const recordings = await RecordingsDB.listRecordings(categoryId);
    set({ recordings });
  },

  loadCategories: async () => {
    const categories = await CategoriesDB.listCategories();
    set({ categories });
  },

  addRecording: async (rec) => {
    await RecordingsDB.insertRecording(rec);
    await get().loadRecordings();
  },

  updateRecording: async (id, patch) => {
    await RecordingsDB.updateRecording(id, patch);
    set((state) => ({
      recordings: state.recordings.map((r) =>
        r.id === id ? { ...r, ...patch, updatedAt: new Date().toISOString() } : r
      ),
    }));
  },

  deleteRecording: async (id) => {
    await RecordingsDB.deleteRecording(id);
    set((state) => ({ recordings: state.recordings.filter((r) => r.id !== id) }));
  },

  addCategory: async (cat) => {
    await CategoriesDB.insertCategory(cat);
    set((state) => ({ categories: [...state.categories, cat] }));
  },

  deleteCategory: async (id) => {
    await CategoriesDB.deleteCategory(id);
    set((state) => ({ categories: state.categories.filter((c) => c.id !== id) }));
  },
}));
