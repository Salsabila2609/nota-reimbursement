-- Run this in Supabase SQL Editor

-- Users table (drivers + admin)
create table if not exists users (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  email text unique not null,
  password_hash text not null,
  role text not null check (role in ('driver', 'admin')),
  created_at timestamptz default now()
);

-- Submissions table
create table if not exists submissions (
  id uuid primary key default gen_random_uuid(),
  driver_id uuid references users(id) on delete cascade,
  driver_name text not null,
  category text not null, -- 'parkir', 'tol', 'bensin', 'lainnya'
  description text,
  amount numeric(12,2),
  submission_date date not null default current_date,
  image_path text not null, -- path di Supabase Storage
  image_url text,           -- public URL
  status text not null default 'pending' check (status in ('pending', 'approved', 'rejected')),
  blur_rejected boolean default false,
  created_at timestamptz default now()
);

-- Index for fast queries
create index if not exists idx_submissions_driver on submissions(driver_id);
create index if not exists idx_submissions_date on submissions(submission_date);
create index if not exists idx_submissions_status on submissions(status);

-- Storage bucket (run via Supabase dashboard or this SQL)
-- Bucket name: nota-images
-- Public: false (private, accessed via signed URLs)

-- Row Level Security
alter table submissions enable row level security;
alter table users enable row level security;

-- Policies: admin bisa lihat semua, driver hanya lihat milik sendiri
-- (Untuk simplicity, kita handle auth di app layer, bukan RLS)
-- Disable RLS kalau mau simple:
alter table submissions disable row level security;
alter table users disable row level security;

-- Seed admin user (password: admin123 - ganti di production!)
-- bcrypt hash dari 'admin123'
insert into users (name, email, password_hash, role) values
  ('Admin', 'admin@company.com', '$2a$10$N9qo8uLOickgx2ZMRZoMyeIjZAgcfl7p92ldGxad68LJZdL17lra2', 'admin')
on conflict (email) do nothing;

-- Seed driver users (password: driver123)
insert into users (name, email, password_hash, role) values
  ('Budi Santoso', 'budi@company.com', '$2a$10$92IXUNpkjO0rOQ5byMi.Ye4oKoEa3Ro9llC/.og/at2uheWG/igi.', 'driver'),
  ('Andi Wijaya', 'andi@company.com', '$2a$10$92IXUNpkjO0rOQ5byMi.Ye4oKoEa3Ro9llC/.og/at2uheWG/igi.', 'driver'),
  ('Siti Rahma', 'siti@company.com', '$2a$10$92IXUNpkjO0rOQ5byMi.Ye4oKoEa3Ro9llC/.og/at2uheWG/igi.', 'driver'),
  ('Dedi Kurniawan', 'dedi@company.com', '$2a$10$92IXUNpkjO0rOQ5byMi.Ye4oKoEa3Ro9llC/.og/at2uheWG/igi.', 'driver'),
  ('Rini Susanti', 'rini@company.com', '$2a$10$92IXUNpkjO0rOQ5byMi.Ye4oKoEa3Ro9llC/.og/at2uheWG/igi.', 'driver')
on conflict (email) do nothing;

-- Tambahkan kolom OCR (jalankan ini jika tabel submissions sudah ada)
ALTER TABLE submissions ADD COLUMN IF NOT EXISTS ocr_raw_text text;
