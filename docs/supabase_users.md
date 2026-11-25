# Menambah pengguna cloud (v1.9.7)

Login memakai **username + password**. Di Supabase, username `dea` disimpan sebagai email `dea@flowbpa.local` (tidak ada email yang dikirim).

## Tambah pengguna
1. Supabase → **Authentication → Users → Add user → Create new user**.
2. Email: `<username>@flowbpa.local` (huruf kecil, tanpa spasi), isi password awal, **centang Auto Confirm User**.
3. Daftarkan sebagai anggota tim (SQL Editor):
   ```sql
   insert into public.members (email, name) values ('<username>@flowbpa.local', '<Nama>')
   on conflict (email) do nothing;
   ```
4. Beri tahu username + password awal ke orangnya; mereka bisa ganti lewat ☁ Cloud → **Ganti password**.

## Lupa password / set password oleh admin
Tidak ada email nyata, jadi "Send password recovery" tidak bisa dipakai. Admin set password baru lewat SQL Editor:
```sql
update auth.users set encrypted_password = crypt('<password baru>', gen_salt('bf'))
where email = '<username>@flowbpa.local';
```

## Akun email lama
Akun yang dulu login lewat link email bisa tetap dipakai: beri password dengan SQL di atas (pakai alamat email aslinya), lalu di layar login ketik alamat email lengkap sebagai username.
