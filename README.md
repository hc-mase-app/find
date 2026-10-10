# Find Cell Online

Versi online aplikasi penjualan Find Cell.

## Komponen

- Aplikasi: Python HTTP handler
- Hosting: Vercel
- Database: Neon Postgres

## Variabel aplikasi

- `DATABASE_URL`: koneksi Postgres berpola pooled untuk aplikasi.
- `DATABASE_URL_UNPOOLED`: koneksi langsung untuk migrasi data awal.

## Menjalankan secara lokal

Tanpa `DATABASE_URL`, aplikasi memakai SQLite lokal pada folder `data`.

```powershell
python app.py
```

Dengan `DATABASE_URL`, aplikasi memakai database Neon.

## Migrasi data awal

Jalankan `migrate_sqlite_to_postgres.py` setelah koneksi Neon tersedia. Skrip mempertahankan ID dan memeriksa jumlah baris setiap tabel.
