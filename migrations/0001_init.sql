CREATE TABLE IF NOT EXISTS users (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  email TEXT NOT NULL UNIQUE,
  display_name TEXT NOT NULL,
  phone TEXT,
  role TEXT NOT NULL CHECK (role IN ('admin', 'pulizie', 'proprietario')),
  is_active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS apartment_permissions (
  user_id INTEGER NOT NULL,
  apartment_id TEXT NOT NULL,
  PRIMARY KEY (user_id, apartment_id),
  FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS module_permissions (
  user_id INTEGER NOT NULL,
  module_key TEXT NOT NULL,
  PRIMARY KEY (user_id, module_key),
  FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS app_settings (
  key TEXT PRIMARY KEY,
  value TEXT NOT NULL,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS workbook_uploads (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  object_key TEXT NOT NULL,
  original_name TEXT NOT NULL,
  uploaded_by TEXT,
  uploaded_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  is_current INTEGER NOT NULL DEFAULT 0
);
