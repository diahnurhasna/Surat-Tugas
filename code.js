function processSection(section, data, hasTeknisi2) {
  if (!section) {
    return;
  }

  const sectionType = section.getType(); // Untuk logging

  // pattern
  const patternNomorSurat = "\\<\\<\\s*" + "Nomor Surat" + "\\s*\\>\\>";
  const patternTanggalSurat = "\\<\\<\\s*" + "Tanggal Surat" + "\\s*\\>\\>";
  const patternNamaHari = "\\<\\<\\s*" + "Nama Hari" + "\\s*\\>\\>";
  const patternNamaManajer = "\\<\\<\\s*" + "Nama Manajer" + "\\s*\\>\\>";
  const patternTujuanTugas = "\\<\\<\\s*" + "Tujuan Tugas" + "\\s*\\>\\>";
  const patternTanggalMulai = "\\<\\<\\s*" + "Tanggal Mulai" + "\\s*\\>\\>";
  const patternTanggalSelesai = "\\<\\<\\s*" + "Tanggal Selesai" + "\\s*\\>\\>";
  const patternLokasiTugas = "\\<\\<\\s*" + "Lokasi Tugas" + "\\s*\\>\\>";

  section.replaceText(patternNomorSurat, data.nomorSurat);
  section.replaceText(patternTanggalSurat, data.tanggalSurat);
  section.replaceText(patternNamaHari, data.namaHari);
  section.replaceText(patternNamaManajer, data.namamanajer || "");
  section.replaceText(patternTujuanTugas, data.tujuantugas || "");
  section.replaceText(patternTanggalMulai, data.tanggalmulai || "");
  section.replaceText(patternTanggalSelesai, data.tanggalselesai || "");
  section.replaceText(patternLokasiTugas, data.lokasitugas || "");
  let nikmanajer = "";
  let jabatanmanajer = "";
  if (data.namamanajer == "FX. SIGIT EKO PRAYOGO") {
    nikmanajer = "906535";
    jabatanmanajer = "Mgr Shared Service Pekanbaru";
  }
  if (data.namamanajer == "HAFITRA HARIANSYAH") {
    nikmanajer = "865829";
    jabatanmanajer = "Mgr Wilayah Pekanbaru";
  }
  if (data.namamanajer == "BUDI SANTOSO") {
    nikmanajer = "885773";
    jabatanmanajer = "Project Mgr Area Pekanbaru";
  }
  if (data.namamanajer == "EKA BANGKIT PRASTYA") {
    nikmanajer = "905962";
    jabatanmanajer = "Project Mgr Area Pekanbaru";
  }
  const patternNikManajer = "\\<\\<\\s*" + "NIK Manajer" + "\\s*\\>\\>";
  const patternJabatanManajer = "\\<\\<\\s*" + "Jabatan Manajer" + "\\s*\\>\\>";

  const textNikManajer = "Nik. " + nikmanajer;

  section.replaceText(patternNikManajer, textNikManajer);
  section.replaceText(patternJabatanManajer, jabatanmanajer);

  // 2. Cari tabel HANYA di section ini
  const tables = section.getTables();

  if (tables.length > 0) {
    // Asumsi kita hanya peduli pada tabel pertama di section ini
    const table = tables[0];
    Logger.log("Table found in section: " + sectionType);

    // 3. Logika menambah baris Teknisi 2 - PENDEKATAN BARU
    if (hasTeknisi2) {
      Logger.log("Teknisi 2 data exists. Attempting to add row in " + sectionType);
      try {
        // Ambil baris template (Teknisi 1) di index 1
        const templateRow = table.getRow(1);

        if (templateRow) {
          
          // --- INI SOLUSI BARUNYA ---
          // 1. Tentukan indeks untuk menyisipkan baris baru (di bawah baris template)
          const insertIndex = table.getChildIndex(templateRow) + 1;

          // 2. Buat SALINAN (copy) dari baris template.
          //    Ini akan menyalin SEMUA format (border, font, bg color, dll)
          const newRow = table.insertTableRow(insertIndex, templateRow.copy());
          // --- SELESAI SOLUSI ---

          // 3. 'newRow' sekarang adalah salinan persis dari baris Teknisi 1.
          //    Kita hanya perlu mengganti placeholder-nya dengan data Teknisi 2.
          
          // Definisikan pattern untuk placeholder ASLI (Teknisi 1)
          const patternNamaTeknisi = "\\<\\<\\s*" + "Nama Teknisi 1" + "\\s*\\>\\>";
          const patternNIKTeknisi = "\\<\\<\\s*" + "NIK Teknisi 1" + "\\s*\\>\\>";
          const patternPosisiTeknisi = "\\<\\<\\s*" + "Posisi Teknisi 1" + "\\s*\\>\\>";
          const patternPSATeknisi = "\\<\\<\\s*" + "PSA Teknisi" + "\\s*\\>\\>";

          // Ganti placeholder di baris BARU (newRow) dengan data Teknisi 2
          newRow.replaceText(patternNamaTeknisi, data.namateknisi2 || "");
          newRow.replaceText(patternNIKTeknisi, data.nikteknisi2 || "");
          newRow.replaceText(patternPosisiTeknisi, data.posisiteknisi2 || "");
          newRow.replaceText(patternPSATeknisi, data.psateknisi2 || "");
          
          // Set nomor urut secara manual
          newRow.getCell(0).setText("2");

          Logger.log("Successfully added AND formatted row for Teknisi 2 in " + sectionType);
        } else {
          Logger.log("Table found, but template row (Row 1) is missing in " + sectionType);
        }
      } catch (tableErr) {
        Logger.log("ERROR saat mencoba menambah baris tabel di " + sectionType + ": " + tableErr);
        
        // FALLBACK: (Biarkan ini sebagai cadangan jika metode copy gagal)
        try {
          Logger.log("Trying ultra-simple method...");
          const newRow = table.appendTableRow();
          
          // Langsung isi data tanpa copy template
          newRow.appendTableCell("2"); // No urut
          newRow.appendTableCell(data.namateknisi2 || ""); // Nama
          newRow.appendTableCell(data.nikteknisi2 || ""); // NIK
          newRow.appendTableCell(data.posisiteknisi2 || ""); // Posisi
          newRow.appendTableCell(data.psateknisi2 || ""); // PSA
          
          Logger.log("Successfully added row for Teknisi 2 using ultra-simple method in " + sectionType);
        } catch (simpleErr) {
          Logger.log("Ultra-simple method also failed: " + simpleErr);
        }
      }
    }

    // 4. Ganti placeholder Teknisi 1 di section ini
    // Ini akan mengisi baris template asli
    section.replaceText("<<Nama Teknisi 1>>", data.namateknisi1 || "");
    section.replaceText("<<NIK Teknisi 1>>", data.nikteknisi1 || "");
    section.replaceText("<<Posisi Teknisi 1>>", data.posisiteknisi1 || "");
    section.replaceText("<<PSA Teknisi>>", data.psateknisi1 || "");
  }
}
/**
 * Fungsi validasi email sederhana
 */
function isValidEmail(email) {
  if (!email || email.trim() === "") return false;
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email.trim());
}

/**
 * Fungsi utama yang dipicu saat Google Form disubmit.
 */
function onFormSubmit(e) {
  try {
    if (!e || !e.range) {
      Logger.log("Event object missing. Untuk testing jalankan testGenerate().");
      return;
    }

    const sheet = e.source.getActiveSheet();
    const row = e.range.getRow();
    const lastCol = sheet.getLastColumn();

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const data = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
    const timeZone = Session.getScriptTimeZone();

    // Buat objek 'obj' dari data form
    const obj = {};
    for (let i = 0; i < headers.length; i++) {
      const rawHeader = (headers[i] || "").toString().trim();
      const key = rawHeader.toLowerCase().replace(/\s+/g, "").replace(/[^a-z0-9]/g, "");
      let value = data[i];

      if (value instanceof Date) {
        value = Utilities.formatDate(value, timeZone, "dd MMMM yyyy");
      } else if (value === null || value === undefined) {
        value = "";
      } else {
        value = value.toString();
      }
      obj[key] = value;
    }

    // Tambahkan data non-form ke 'obj' agar bisa dilempar ke helper
    const today = new Date();
    obj.nomorSurat = "353/UM.000/TA." + Utilities.formatDate(today, timeZone, "MMdd") + "/2025";
    obj.tanggalSurat = Utilities.formatDate(today, timeZone, "dd MMMM yyyy");

    // --- PERBAIKAN NAMA HARI (BAHASA INDONESIA) ---
    // Buat array lookup untuk nama hari
    const namaHariIndonesia = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
    // Dapatkan indeks hari (0 = Minggu, 1 = Senin, dst.)
    const dayIndex = today.getDay();
    // Ambil nama hari dari array
    obj.namaHari = namaHariIndonesia[dayIndex];
    // --- SELESAI PERBAIKAN ---

    Logger.log("Mapping: " + JSON.stringify(obj));

    // --- Persiapan Dokumen ---
    const templateId = "1mIOrMHiCEVWccUOlCP5N6EPqUEq2w9W7e48drgf96qc";
    const folderId = "14MWia7aR8Og20886EFJEgXa4evo9YLN6";

    const copyName = "Surat Tugas - " + (obj.namateknisi1 || "NoName");
    const docFileCopied = DriveApp.getFileById(templateId).makeCopy(copyName, DriveApp.getFolderById(folderId));
    const doc = DocumentApp.openById(docFileCopied.getId());

    // --- LOGIKA UTAMA (YANG BARU) ---
    const hasTeknisi2 = (obj.namateknisi2 && obj.namateknisi2.trim() !== "");

    // 1. Proses Body Dokumen
    processSection(doc.getBody(), obj, hasTeknisi2);
    // 2. Proses Header Dokumen
    processSection(doc.getHeader(), obj, hasTeknisi2);
    // 3. Proses Footer Dokumen
    processSection(doc.getFooter(), obj, hasTeknisi2);

    // --- Selesai & Kirim Email ---
    doc.saveAndClose();
    Logger.log("Dokumen berhasil dibuat: " + doc.getUrl());

    const pdf = DriveApp.getFileById(doc.getId()).getAs("application/pdf");
    const pdfFile = DriveApp.getFolderById(folderId).createFile(pdf);

    const emailTujuan = obj.emailteknisi;
    Logger.log("Email teknisi dari form: '" + emailTujuan + "'. Mengirim email...");

    // Validasi email
    if (emailTujuan && emailTujuan.trim() !== "" && isValidEmail(emailTujuan.trim())) {
      MailApp.sendEmail({
        to: emailTujuan.trim(),
        subject: "Approval Surat Tugas - " + (obj.namateknisi1 || "NoName"),
        body: "Halo,\n\nBerikut terlampir surat tugas, silahkan print dan minta ttd approval.\n\nSalam.",
        attachments: [pdfFile]
      });
      Logger.log("Selesai: file dibuat dan email dikirim.");
    } else {
      Logger.log("Email tujuan kosong atau tidak valid, email tidak dikirim: " + emailTujuan);
    }

  } catch (err) {
    Logger.log("FATAL ERROR in onFormSubmit: " + err.toString() + "\nStack: " + err.stack);
  }
}

/**
 * Fungsi helper untuk testing.
 */
function testGenerate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const lastRow = sheet.getLastRow();
  const e = {
    source: ss,
    range: sheet.getRange(lastRow, 1)
  };
  onFormSubmit(e);
}

/**
 * Fungsi helper (tidak terpakai saat ini).
 */
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
