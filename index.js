import express from 'express';
import cors from 'cors';
import multer from 'multer';
import dotenv from 'dotenv';
import readline from 'readline';
import fs from 'fs';
import xlsx from 'xlsx';
import { createClient } from '@supabase/supabase-js';

dotenv.config();

const app = express();

// Middleware harus diletakkan di atas semua route
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Inisialisasi Supabase
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY
);

const upload = multer({ dest: 'uploads/' });

// --- API KUIS ---
app.post('/api/kuis/import', upload.single('file'), async (req, res) => {
  try {
    const { guru_id, judul, kelas, deskripsi, kkm, deadline, is_published } = req.body;
    
    if (!req.file) {
      return res.status(400).json({ success: false, error: "File tidak ditemukan" });
    }

    const filePath = req.file.path;
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    fs.unlinkSync(filePath); // Hapus file temporary setelah dibaca

    // 1. Simpan data utama ke tabel 'kuis'
    const { data: kuisData, error: kuisError } = await supabase
      .from('kuis')
      .insert([{
        guru_id,
        judul,
        kelas,
        deskripsi: deskripsi || null,
        kkm: Number(kkm) || 70,
        deadline: deadline ? new Date(deadline).toISOString() : null,
        mode_kuis: req.body.mode_kuis || 'import',
        is_published: is_published === 'true'
      }])
      .select()
      .single();

    if (kuisError) throw new Error(kuisError.message);

    // 2. Mapping soal dari file Excel
    const soalPayload = sheetData.map((row, index) => {
      const pertanyaan = row.pertanyaan || row.Pertanyaan || row.question || '';
      const opsi_a = row.opsi_a || row.A || '';
      const opsi_b = row.opsi_b || row.B || '';
      const opsi_c = row.opsi_c || row.C || '';
      const opsi_d = row.opsi_d || row.D || '';

      // Tentukan tipe soal
      const isEssay = !opsi_a && !opsi_b && !opsi_c && !opsi_d;

      return {
        kuis_id: kuisData.id,
        pertanyaan,
        tipe_soal: isEssay ? 'essay' : 'pg',
        opsi_a: isEssay ? null : opsi_a,
        opsi_b: isEssay ? null : opsi_b,
        opsi_c: isEssay ? null : opsi_c,
        opsi_d: isEssay ? null : opsi_d,
        jawaban_benar: row.jawaban_benar || row.jawaban || row.Kunci || '',
        bobot: 10,
        urutan: index + 1
      };
    }).filter(q => q.pertanyaan); // Pastikan pertanyaan tidak kosong

    // 3. Simpan soal ke tabel 'kuis_soal'
    const { error: soalError } = await supabase
      .from('kuis_soal')
      .insert(soalPayload);

    if (soalError) throw new Error(soalError.message);

    res.json({ success: true, message: "Kuis berhasil diimport!" });

  } catch (err) {
    console.error("Import Error:", err);
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
    res.status(500).json({ success: false, error: err.message });
  }
});

app.post('/api/kuis/preview', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: "File tidak ditemukan" });
    }

    const filePath = req.file.path;

    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

    fs.unlinkSync(filePath);

    if (sheetData.length === 0) {
      return res.status(400).json({ success: false, error: "File kosong" });
    }

    let parsedSoal = [];
    sheetData.forEach((row, index) => {
      const pertanyaan = row.pertanyaan || row.Pertanyaan || row.question || '';
      if (!pertanyaan) return;

      const opsi_a = row.opsi_a || row.A || '';
      const opsi_b = row.opsi_b || row.B || '';
      const opsi_c = row.opsi_c || row.C || '';
      const opsi_d = row.opsi_d || row.D || '';

      // Tentukan tipe soal: jika semua opsi kosong, maka dianggap 'essay'
      const isEssay = !opsi_a && !opsi_b && !opsi_c && !opsi_d;

      parsedSoal.push({
        urutan: parsedSoal.length + 1,
        tipe: isEssay ? 'essay' : 'pg',
        pertanyaan,
        opsi_a,
        opsi_b,
        opsi_c,
        opsi_d,
        jawaban_benar: row.jawaban_benar || row.jawaban || row.Kunci || ''
      });
    });

    res.json({ 
      success: true, 
      soal: parsedSoal 
    });

  } catch (err) {
    console.error("Preview Error:", err);
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
    res.status(500).json({ success: false, error: err.message });
  }
});

// --- API ADMIN USERS ---
app.post('/api/admin/users', async (req, res) => {
  try {
    const payload = req.body;

    if (!payload.email || !payload.nama || !payload.role) {
      return res.status(400).json({ success: false, error: "Data wajib diisi lengkap" });
    }

    // 1. Buat User di Supabase Auth (Masuk ke auth.users)
    const { data: authData, error: authError } = await supabase.auth.admin.createUser({
      email: payload.email,
      password: payload.password || 'Password123',
      email_confirm: true,
      user_metadata: { 
        nama: payload.nama, 
        role: payload.role 
      }
    });

    if (authError) {
      console.error("Auth Error:", authError);
      throw authError;
    }

    const userId = authData.user.id;

    // 2. Simpan juga ke tabel 'profiles' (Agar terbaca di Manajemen User)
    const { error: profileError } = await supabase.from('profiles').insert({
      id: userId,
      nama: payload.nama,
      email: payload.email,
      role: payload.role
    });

    if (profileError) {
      console.error("Profile Error:", profileError);
    }

    // 3. JIKA ROLE-NYA SISWA, masukkan data ke tabel 'students'
    if (payload.role === 'siswa') {
      const studentData = {
        id: userId, // Menghubungkan relasi ID Auth
        nama: payload.nama,
      };

      if (payload.nisn) studentData.nisn = payload.nisn;
      if (payload.nis) studentData.nis = payload.nis;
      if (payload.kelas) studentData.kelas = payload.kelas;
      if (payload.tempat_lahir) studentData.tempat_lahir = payload.tempat_lahir;
      if (payload.tanggal_lahir) studentData.tanggal_lahir = payload.tanggal_lahir;
      if (payload.nama_ibu) studentData.nama_ibu = payload.nama_ibu;
      if (payload.nama_ayah) studentData.nama_ayah = payload.nama_ayah;
      if (payload.no_hp_ortu) studentData.no_hp_ortu = payload.no_hp_ortu;
      if (payload.pekerjaan_ortu) studentData.pekerjaan_ortu = payload.pekerjaan_ortu;

      const { error: studentError } = await supabase.from('students').insert(studentData);

      if (studentError) {
        console.error("Student Error:", studentError);
        throw studentError;
      }
    }

    res.status(201).json({ success: true, message: "User, Profile, dan Data Siswa berhasil dibuat" });
  } catch (err) {
    console.error("Backend Error Detail:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// --- API IMPORT CSV SISWA ---
app.post('/api/admin/users/import-csv', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, error: "File CSV tidak ditemukan" });
    }

    const filePath = req.file.path;
    const fileStream = fs.createReadStream(filePath);

    const rl = readline.createInterface({
      input: fileStream,
      crlfDelay: Infinity
    });

    let lines = [];
    for await (const line of rl) {
      if (line.trim()) lines.push(line);
    }

    // Hapus file temporary setelah dibaca
    fs.unlinkSync(filePath);

    if (lines.length === 0) {
      return res.status(400).json({ success: false, error: "File CSV kosong" });
    }

    let successCount = 0;
    let errors = [];

    // Lewati baris pertama jika itu adalah header (NISN, Nama, Kelas)
    const startIndex = lines[0].toLowerCase().includes('nisn') ? 1 : 0;

    for (let i = startIndex; i < lines.length; i++) {
      const parts = lines[i].split(',').map(item => item.trim());
      const [nisn, nama, kelas] = parts;

      if (!nisn || !nama) continue;

      // Masukkan langsung ke tabel students
      const { error: dbError } = await supabase.from('students').insert({
        nisn: nisn,
        nama: nama,
        kelas: kelas || null
      });

      if (dbError) {
        errors.push(`Baris ${i + 1} (${nisn}): ${dbError.message}`);
      } else {
        successCount++;
      }
    }

    res.json({ 
      success: true, 
      message: `Berhasil mengimpor ${successCount} data siswa!`,
      errors: errors.length > 0 ? errors : undefined
    });

  } catch (err) {
    console.error("CSV Import Error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// --- API STUDENTS ---
app.post('/api/students/create', async (req, res) => {
  try {
    const { nisn, nama, kelas } = req.body;
    
    const { error: dbError } = await supabase.from('students').insert({ 
      nisn, nama, kelas 
    });
    if (dbError) throw dbError;

    res.json({ success: true, message: "Siswa berhasil didaftarkan" });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// --- API VALIDASI MANDIRI (UNTUK PANEL GURU) ---
app.get('/api/teacher/attendance-validations', async (req, res) => {
  try {
    const { data, error } = await supabase
      .from('absensi')
      .select('*');

    if (error) throw error;

    res.json({ success: true, data: data || [] });
  } catch (err) {
    console.error("Error fetching validations:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// --- API UPDATE STATUS VALIDASI ABSENSI MANDIRI (OTOMATIS SYNC STATUS) ---
app.patch('/api/teacher/attendance-validations/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { status_validasi, catatan_guru, status_pengajuan_siswa } = req.body; 

    if (!['valid', 'ditolak'].includes(status_validasi)) {
      return res.status(400).json({ success: false, error: "Status validasi tidak valid" });
    }

    // Jika disetujui ('valid'), ikuti status pengajuan siswa (misal: 'izin', 'sakit'). 
    // Jika ditolak ('ditolak'), otomatis ubah status harian menjadi 'alpha'.
    const statusHarianBaru = status_validasi === 'valid' ? (status_pengajuan_siswa || 'izin') : 'alpha';

    const { error: updateError } = await supabase
      .from('absensi')
      .update({
        status_validasi: status_validasi,
        catatan_guru: catatan_guru || null,
        status: statusHarianBaru // Sinkronisasi otomatis ke absensi harian
      })
      .eq('id', id);

    if (updateError) throw updateError;

    res.json({ 
      success: true, 
      message: status_validasi === 'valid' ? "Absensi berhasil divalidasi!" : "Absensi ditolak dan otomatis diubah menjadi Alpha!" 
    });

  } catch (err) {
    console.error("Update Validation Error:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// --- API STATISTIK DISTRIBUSI SISWA PER KELAS ---
app.get('/api/dashboard/distribusi-kelas', async (req, res) => {
  try {
    const { data, error } = await supabase
      .from('students')
      .select('kelas');

    if (error) throw error;

    // Hitung jumlah siswa per kelas secara otomatis
    const distribusi = {};
    (data || []).forEach(item => {
      const kelas = item.kelas || 'Belum Diatur';
      distribusi[kelas] = (distribusi[kelas] || 0) + 1;
    });

    // Ubah ke format array agar mudah di-mapping di frontend
    const result = Object.keys(distribusi).map(kelas => ({
      kelas: kelas,
      jumlah: distribusi[kelas]
    }));

    res.json({ success: true, data: result });
  } catch (err) {
    console.error("Error distribusi kelas:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});

import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 1. Agar Express membaca file statis hasil build React (pastikan folder 'public' berisi file build dist Anda)
app.use(express.static(path.join(__dirname, 'public')));

// 2. SPA Fallback: Semua rute selain /api akan diarahkan ke index.html supaya React Router tidak 404
app.get(/^(?!\/api).*/, (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Penanganan route tidak ditemukan agar tetap JSON
app.use((req, res) => {
  res.status(404).json({ success: false, error: "Endpoint tidak ditemukan" });
});

app.listen(process.env.PORT || 3000, () => {
  console.log('Server running on port 3000');
});