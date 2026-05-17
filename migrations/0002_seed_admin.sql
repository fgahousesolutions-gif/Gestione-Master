INSERT INTO users (email, display_name, role, is_active)
VALUES ('fgahousesolutions@gmail.com', 'FGA House Solutions', 'admin', 1)
ON CONFLICT(email) DO UPDATE SET
  display_name = excluded.display_name,
  role = excluded.role,
  is_active = excluded.is_active;
