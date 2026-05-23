export interface Recording {
  id: string;
  title: string;
  audioUri: string;
  duration: number; // seconds
  transcript: string;
  categoryId: string | null;
  createdAt: string; // ISO
  updatedAt: string; // ISO
}

export interface Category {
  id: string;
  name: string;
  color: string;
  createdAt: string;
}

// Add-on plugin interface — every add-on implements this
export interface TextAddon {
  id: string;
  name: string;
  description: string;
  icon: string;
  run: (text: string) => Promise<string>;
}
