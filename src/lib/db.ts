const DB_NAME = "StudentGrowthDashboardDB";
const DB_VERSION = 1;

export interface SettingData {
  id: "main";
  apiKey: string;
  model: string;
  rememberKey: boolean;
  rubric1Mode: "file" | "text";
  rubric1Text: string;
  rubric1Files: Array<{ id: string; name: string; type: string; size: number; base64: string }>;
  rubric2Mode: "file" | "text";
  rubric2Text: string;
  rubric2Files: Array<{ id: string; name: string; type: string; size: number; base64: string }>;
}

export interface StudentData {
  id: string;
  name: string;
  className: string;
  number: string;
  firstFiles: Array<{ id: string; name: string; type: string; size: number; base64: string }>;
  secondFiles: Array<{ id: string; name: string; type: string; size: number; base64: string }>;
  analysis: null | {
    status: "running" | "done" | "error";
    result?: any; // The schema-compliant JSON output from Gemini
    error?: string;
    analyzedAt: number;
  };
  createdAt: number;
}

function openDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onerror = () => reject(request.error);
    request.onsuccess = () => resolve(request.result);

    request.onupgradeneeded = (event) => {
      const db = request.result;
      if (!db.objectStoreNames.contains("settings")) {
        db.createObjectStore("settings", { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains("students")) {
        db.createObjectStore("students", { keyPath: "id" });
      }
    };
  });
}

export const dbService = {
  // --- Settings Store ---
  async getSettings(): Promise<SettingData> {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction("settings", "readonly");
      const store = transaction.objectStore("settings");
      const request = store.get("main");

      request.onerror = () => reject(request.error);
      request.onsuccess = () => {
        const defaultSettings: SettingData = {
          id: "main",
          apiKey: "",
          model: "gemini-2.5-flash",
          rememberKey: false,
          rubric1Mode: "file",
          rubric1Text: "",
          rubric1Files: [],
          rubric2Mode: "file",
          rubric2Text: "",
          rubric2Files: [],
        };
        resolve(request.result || defaultSettings);
      };
    });
  },

  async saveSettings(settings: SettingData): Promise<void> {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction("settings", "readwrite");
      const store = transaction.objectStore("settings");
      const request = store.put(settings);

      request.onerror = () => reject(request.error);
      request.onsuccess = () => resolve();
    });
  },

  // --- Students Store ---
  async getAllStudents(): Promise<StudentData[]> {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction("students", "readonly");
      const store = transaction.objectStore("students");
      const request = store.getAll();

      request.onerror = () => reject(request.error);
      request.onsuccess = () => {
        const list = request.result || [];
        // Sort by number, then className, then createdAt
        list.sort((a, b) => {
          const numA = parseInt(a.number) || 0;
          const numB = parseInt(b.number) || 0;
          if (numA !== numB) return numA - numB;
          return a.createdAt - b.createdAt;
        });
        resolve(list);
      };
    });
  },

  async getStudent(id: string): Promise<StudentData | undefined> {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction("students", "readonly");
      const store = transaction.objectStore("students");
      const request = store.get(id);

      request.onerror = () => reject(request.error);
      request.onsuccess = () => resolve(request.result);
    });
  },

  async saveStudent(student: StudentData): Promise<void> {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction("students", "readwrite");
      const store = transaction.objectStore("students");
      const request = store.put(student);

      request.onerror = () => reject(request.error);
      request.onsuccess = () => resolve();
    });
  },

  async deleteStudent(id: string): Promise<void> {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction("students", "readwrite");
      const store = transaction.objectStore("students");
      const request = store.delete(id);

      request.onerror = () => reject(request.error);
      request.onsuccess = () => resolve();
    });
  },

  async clearAllData(): Promise<void> {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(["settings", "students"], "readwrite");
      const settingsStore = transaction.objectStore("settings");
      const studentsStore = transaction.objectStore("students");

      settingsStore.clear();
      studentsStore.clear();

      transaction.onerror = () => reject(transaction.error);
      transaction.oncomplete = () => resolve();
    });
  }
};
