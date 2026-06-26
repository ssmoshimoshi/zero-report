const CONFIG = {
  SHEET_ID: '197JbA2mdWC-H1oaQvwDT51e6UwgwUe5Jt2yb6064awc',
  ROOT_FOLDER_ID: '1cpwnFb5lh4OVJxbFpezBA48iLEgISZyj'
};

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Zero Cafe Workspace')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getOrCreateSubFolder(parentFolder, folderName) {
  const folders = parentFolder.searchFolders("title = '" + folderName + "'");
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

function getStructuredFolder(reportType, dateObj) {
  const rootFolder = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  const year = dateObj.getFullYear().toString();
  const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  const month = monthNames[dateObj.getMonth()];
  const yearFolder = getOrCreateSubFolder(rootFolder, year);
  const monthFolder = getOrCreateSubFolder(yearFolder, month);
  let categoryName = "Daily Report";
  if (reportType === 'weekly') categoryName = "Weekly Report";
  if (reportType === 'monthly') categoryName = "Monthly Report";
  return getOrCreateSubFolder(monthFolder, categoryName);
}

function submitFullReport(payloadStr) {
  try {
    const data = JSON.parse(payloadStr);
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const reportType = data.type || 'daily';
    
    if (reportType === 'daily') {
      const sheet = ss.getSheetByName('Daily');
      if (!sheet) throw new Error("Tab 'Daily' tidak ditemukan di Spreadsheet.");
      
      const tanggal = data.tanggal || "-";
      const supervisor = data.supervisor || "-";
      const shift = data.shift || "-";
      const outlet = data.outlet || "Perintis";
      
      const omsetShift1 = Number(data.penjualan?.shift1) || 0;
      const omsetShift2 = Number(data.penjualan?.shift2) || 0;
      const totalOmset = omsetShift1 + omsetShift2;
      const transaksi = data.penjualan?.transaksi || 0;
      const komplain = data.feedback?.totalKomplain || 0;
      const totalRemake = data.feedback?.totalRemake || 0;
      const kendala = data.penutup?.kendala || "-";
      
      let auditTableRows = '';
      if (data.kas?.audit && data.kas.audit.length > 0) {
        data.kas.audit.forEach(a => {
          auditTableRows += `<tr>
            <td style="border:1px solid #ccc; padding:5px;">${a.jam || '-'}</td>
            <td style="border:1px solid #ccc; padding:5px;">Rp ${(a.aktual || 0).toLocaleString('id-ID')}</td>
            <td style="border:1px solid #ccc; padding:5px;">Rp ${(a.qris || 0).toLocaleString('id-ID')}</td>
            <td style="border:1px solid #ccc; padding:5px;">Rp ${(a.tunai || 0).toLocaleString('id-ID')}</td>
            <td style="border:1px solid #ccc; padding:5px;">${a.keterangan || '-'}</td>
          </tr>`;
        });
      } else {
        auditTableRows = '<tr><td colspan="5" style="border:1px solid #ccc; padding:5px; text-align:center;">Belum ada data audit kas</td></tr>';
      }
      
      let staffTableRows = '';
      if (data.staff && data.staff.length > 0) {
        data.staff.forEach(s => {
          let statusText = s.status === 'Telat' ? '<strong style="color:red;">TELAT</strong>' : 'HADIR';
          let keramahanText = s.keramahan === 'Ya' ? '<strong style="color:red;">YA</strong>' : 'Tidak';
          staffTableRows += `<tr>
            <td style="border:1px solid #ccc; padding:5px;">${s.nama} (${s.posisi})</td>
            <td style="border:1px solid #ccc; padding:5px; text-align:center;">${statusText}</td>
            <td style="border:1px solid #ccc; padding:5px; text-align:center;">${keramahanText}</td>
            <td style="border:1px solid #ccc; padding:5px;">${s.keterangan || '-'}</td>
          </tr>`;
        });
      } else {
        staffTableRows = '<tr><td colspan="4" style="border:1px solid #ccc; padding:5px; text-align:center;">Tidak ada data staff</td></tr>';
      }
      
      let qcTableRows = `<tr>
        <td style="border:1px solid #ccc; padding:5px;">${data.qc?.espresso?.jam || '-'}</td>
        <td style="border:1px solid #ccc; padding:5px;"><strong>Espresso</strong></td>
        <td style="border:1px solid #ccc; padding:5px;">${data.qc?.espresso?.status || '-'}</td>
        <td style="border:1px solid #ccc; padding:5px;">${data.qc?.espresso?.keterangan || '-'}</td>
      </tr>`;
      if (data.qc?.items && data.qc.items.length > 0) {
        data.qc.items.forEach(q => {
          qcTableRows += `<tr>
            <td style="border:1px solid #ccc; padding:5px;">${q.jam || '-'}</td>
            <td style="border:1px solid #ccc; padding:5px;">${q.nama || '-'}</td>
            <td style="border:1px solid #ccc; padding:5px;">${q.status || '-'}</td>
            <td style="border:1px solid #ccc; padding:5px;">${q.keterangan || '-'}</td>
          </tr>`;
        });
      }
      
      let feedbackTableRows = '';
      if (data.feedback?.items && data.feedback.items.length > 0) {
        data.feedback.items.forEach(fb => {
          feedbackTableRows += `<tr>
            <td style="border:1px solid #ccc; padding:5px;">${fb.jam || '-'}</td>
            <td style="border:1px solid #ccc; padding:5px;">${fb.inisial || '-'}</td>
            <td style="border:1px solid #ccc; padding:5px;">${fb.isi || '-'}</td>
            <td style="border:1px solid #ccc; padding:5px;">${fb.remake || '-'}</td>
            <td style="border:1px solid #ccc; padding:5px;">${fb.eskalasi || 'Tidak'}</td>
            <td style="border:1px solid #ccc; padding:5px;">${fb.respon || '-'}</td>
          </tr>`;
        });
      } else {
        feedbackTableRows = '<tr><td colspan="6" style="border:1px solid #ccc; padding:5px; text-align:center;">Tidak ada insiden komplain</td></tr>';
      }
      
      let fasilitasTableRows = '';
      if (data.fasilitas && data.fasilitas.length > 0) {
        data.fasilitas.forEach(f => {
          fasilitasTableRows += `<tr>
            <td style="border:1px solid #ccc; padding:5px;">${f.item || '-'}</td>
            <td style="border:1px solid #ccc; padding:5px;">${f.status || '-'}</td>
            <td style="border:1px solid #ccc; padding:5px;">${f.eskalasi || '-'}</td>
            <td style="border:1px solid #ccc; padding:5px;">${f.keterangan || '-'}</td>
          </tr>`;
        });
      } else {
        fasilitasTableRows = '<tr><td colspan="4" style="border:1px solid #ccc; padding:5px; text-align:center;">Tidak ada kerusakan fasilitas</td></tr>';
      }
      
      let bahanTableRows = '';
      if (data.bahan && data.bahan.length > 0) {
        data.bahan.forEach(b => {
          bahanTableRows += `<tr>
            <td style="border:1px solid #ccc; padding:5px;">${b.nama || '-'}</td>
            <td style="border:1px solid #ccc; padding:5px;">${b.ketersediaan || '-'}</td>
            <td style="border:1px solid #ccc; padding:5px;">Rp ${(b.harga || 0).toLocaleString('id-ID')}</td>
          </tr>`;
        });
      } else {
        bahanTableRows = '<tr><td colspan="3" style="border:1px solid #ccc; padding:5px; text-align:center;">Aman, tidak ada pengajuan belanja</td></tr>';
      }
      
      const folder = getStructuredFolder('daily', new Date(tanggal !== "-" ? tanggal : new Date()));
      const fileName = `DailyReport_${tanggal}_${supervisor}.pdf`;
     
      const html = `<div style="font-family:sans-serif; padding: 20px;">
        <h2 style="margin-bottom:0;">HARIAN - SUPERVISOR OUTLET ZERO ${outlet.toUpperCase()}</h2>
        <p style="margin-top:5px; color:#555;">Tanggal: ${tanggal} | Nama Supervisor: ${supervisor} | Shift: ${shift}</p>
        <hr>
       
        <h3 style="background:#eee; padding:5px;">A. RINGKASAN OPERASIONAL</h3>
        <table style="width:100%; border-collapse:collapse; margin-bottom:20px; font-size:14px;">
          <tr><td style="padding:5px; border:1px solid #ccc; width:50%;"><strong>Total Penjualan 08:00 - 17:00</strong></td><td style="padding:5px; border:1px solid #ccc;">Rp ${omsetShift1.toLocaleString('id-ID')}</td></tr>
          <tr><td style="padding:5px; border:1px solid #ccc;"><strong>Total Penjualan 17:00 - 20:30</strong></td><td style="padding:5px; border:1px solid #ccc;">Rp ${omsetShift2.toLocaleString('id-ID')}</td></tr>
          <tr><td style="padding:5px; border:1px solid #ccc; background:#f9f9f9;"><strong>Total Omset Aktual</strong></td><td style="padding:5px; border:1px solid #ccc; background:#f9f9f9;"><strong>Rp ${totalOmset.toLocaleString('id-ID')}</strong></td></tr>
          <tr><td style="padding:5px; border:1px solid #ccc;"><strong>Jumlah Transaksi</strong></td><td style="padding:5px; border:1px solid #ccc;">${transaksi} Struk</td></tr>
        </table>
       
        <h3 style="background:#eee; padding:5px;">TOP 3 PRODUK</h3>
        <table style="width:100%; border-collapse:collapse; font-size:14px;">
          <tr>
            <td style="width:50%; vertical-align:top; padding-right:10px;">
              <strong>Makanan:</strong>
              <ol style="margin-top:5px; padding-left:20px;">
                <li>${data.produk?.makanan?.[0] || '-'}</li>
                <li>${data.produk?.makanan?.[1] || '-'}</li>
                <li>${data.produk?.makanan?.[2] || '-'}</li>
              </ol>
            </td>
            <td style="width:50%; vertical-align:top;">
              <strong>Minuman:</strong>
              <ol style="margin-top:5px; padding-left:20px;">
                <li>${data.produk?.minuman?.[0] || '-'}</li>
                <li>${data.produk?.minuman?.[1] || '-'}</li>
                <li>${data.produk?.minuman?.[2] || '-'}</li>
              </ol>
            </td>
          </tr>
        </table>
        
        <h3 style="background:#eee; padding:5px; margin-top:20px;">B. SINKRONISASI KAS KASIR</h3>
        <table style="width:100%; border-collapse:collapse; font-size:12px; text-align:left;">
          <tr style="background:#f9f9f9;">
            <th style="border:1px solid #ccc; padding:5px;">Jam</th>
            <th style="border:1px solid #ccc; padding:5px;">Kas Setelah Audit</th>
            <th style="border:1px solid #ccc; padding:5px;">Qris</th>
            <th style="border:1px solid #ccc; padding:5px;">Tunai</th>
            <th style="border:1px solid #ccc; padding:5px;">Ket</th>
          </tr>
          ${auditTableRows}
        </table>
        
        <h3 style="background:#eee; padding:5px; margin-top:20px;">C. KEPATUHAN STAFF</h3>
        <table style="width:100%; border-collapse:collapse; font-size:12px; text-align:left;">
          <tr style="background:#f9f9f9;">
            <th style="border:1px solid #ccc; padding:5px; width:35%;">Nama & Posisi</th>
            <th style="border:1px solid #ccc; padding:5px; width:15%; text-align:center;">Telat?</th>
            <th style="border:1px solid #ccc; padding:5px; width:20%; text-align:center;">Keramahan Terlewat?</th>
            <th style="border:1px solid #ccc; padding:5px;">Catatan Khusus</th>
          </tr>
          ${staffTableRows}
        </table>
        
        <h3 style="background:#eee; padding:5px; margin-top:20px;">D. POIN BRIEFING SHIFT</h3>
        <table style="width:100%; border-collapse:collapse; font-size:12px;">
          <tr>
            <td style="border:1px solid #ccc; padding:8px; vertical-align:top; width:50%;">
              <strong>Masalah Shift Sebelumnya:</strong><br>
              <span style="white-space:pre-wrap;">${data.briefing?.evaluasi || '-'}</span>
            </td>
            <td style="border:1px solid #ccc; padding:8px; vertical-align:top;">
              <strong>Fokus Perilaku Hari Ini:</strong><br>
              <span style="white-space:pre-wrap;">${data.briefing?.fokus || '-'}</span>
            </td>
          </tr>
        </table>
        
        <h3 style="background:#eee; padding:5px; margin-top:20px;">E. QUALITY CONTROL (QC)</h3>
        <table style="width:100%; border-collapse:collapse; font-size:12px; text-align:left;">
          <tr style="background:#f9f9f9;">
            <th style="border:1px solid #ccc; padding:5px; width:15%;">Waktu QC</th>
            <th style="border:1px solid #ccc; padding:5px; width:25%;">Jenis QC</th>
            <th style="border:1px solid #ccc; padding:5px; width:20%;">Hasil</th>
            <th style="border:1px solid #ccc; padding:5px;">Masalah & Koreksi</th>
          </tr>
          ${qcTableRows}
        </table>
        
        <h3 style="background:#eee; padding:5px; margin-top:20px;">F & G. KOMPLAIN & PENGGANTIAN PRODUK</h3>
        <table style="width:100%; border-collapse:collapse; font-size:12px; text-align:left;">
          <tr style="background:#f9f9f9;">
            <th style="border:1px solid #ccc; padding:5px;">Jam</th>
            <th style="border:1px solid #ccc; padding:5px;">Inisial</th>
            <th style="border:1px solid #ccc; padding:5px;">Isi Komplain</th>
            <th style="border:1px solid #ccc; padding:5px;">Remake</th>
            <th style="border:1px solid #ccc; padding:5px;">Eskalasi?</th>
            <th style="border:1px solid #ccc; padding:5px;">Tindakan SPV</th>
          </tr>
          ${feedbackTableRows}
        </table>
        <p style="font-size:12px; margin-top:8px;"><strong>Total Komplain Hari Ini:</strong> ${komplain} Kasus</p>
        <p style="font-size:12px; margin-top:2px;"><strong>Total Penggantian (Remake):</strong> ${totalRemake} Kasus</p>
        ${totalRemake >= 3 ? `<p style="font-size:12px; margin-top:2px; color:red;"><strong>Analisis Remake (>=3):</strong> ${data.feedback?.analisisRemake || '-'}</p>` : ''}
        
        <h3 style="background:#eee; padding:5px; margin-top:20px;">H. INVENTARIS & FASILITAS</h3>
        <table style="width:100%; border-collapse:collapse; font-size:12px; text-align:left;">
          <tr style="background:#f9f9f9;">
            <th style="border:1px solid #ccc; padding:5px;">Jenis Fasilitas</th>
            <th style="border:1px solid #ccc; padding:5px; width:15%;">Kondisi</th>
            <th style="border:1px solid #ccc; padding:5px; width:20%;">Tindakan</th>
            <th style="border:1px solid #ccc; padding:5px;">Keterangan</th>
          </tr>
          ${fasilitasTableRows}
        </table>
        
        <h3 style="background:#eee; padding:5px; margin-top:20px;">I. BAHAN PENUNJANG (DIBUTUHKAN)</h3>
        <table style="width:100%; border-collapse:collapse; font-size:12px; text-align:left;">
          <tr style="background:#f9f9f9;">
            <th style="border:1px solid #ccc; padding:5px;">Nama Bahan</th>
            <th style="border:1px solid #ccc; padding:5px; width:25%;">Status</th>
            <th style="border:1px solid #ccc; padding:5px; width:25%;">Biaya (Jika Beli)</th>
          </tr>
          ${bahanTableRows}
        </table>
       
        <h3 style="background:#eee; padding:5px; margin-top:20px;">J. KEJADIAN PENTING LAINNYA</h3>
        <p style="font-size:13px;"><strong>Kendala Utama:</strong><br><span style="white-space:pre-wrap;">${kendala}</span></p>
        <p style="font-size:13px; margin-top:10px;"><strong>Rekomendasi Besok:</strong><br><span style="white-space:pre-wrap;">${data.penutup?.rekomendasi || '-'}</span></p>
        
        <h3 style="background:#eee; padding:5px; margin-top:20px;">K. FOTO PENDUKUNG</h3>
        <p style="font-size:12px;">Link folder GDrive Dokumentasi Harian dapat diakses melalui sistem dashboard pusat. Bukti kebersihan, briefing, dan nota telah diarsipkan secara terpisah.</p>
      </div>`;
      
      const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF).setName(fileName);
      const file = folder.createFile(blob);
      
      // Indeks kolom 4 (Total Omset) dan 5 (Komplain) dijaga agar Dashboard GM tetap jalan
      sheet.appendRow([tanggal, supervisor, outlet, shift, totalOmset, komplain, kendala, file.getUrl()]);
    }
    
    else if (reportType === 'weekly') {
      const sheet = ss.getSheetByName('Weekly');
      if (!sheet) throw new Error("Tab 'Weekly' tidak ditemukan di Spreadsheet.");
      
      const periode = data.periode || "-";
      const supervisor = data.supervisor || "-";
      const outlet = data.outlet || "Perintis";
      
      let totalRealSales = 0;
      let salesHarianRows = '';
      if (data.weekly?.salesHarian && data.weekly.salesHarian.length > 0) {
        data.weekly.salesHarian.forEach(s => {
          let tr = Number(s.target) || 0;
          let rl = Number(s.real) || 0;
          let prc = tr > 0 ? ((rl / tr) * 100).toFixed(1) : 0;
          totalRealSales += rl;
          salesHarianRows += `<tr>
            <td style="border:1px solid #ccc; padding:5px;">${s.hari}</td>
            <td style="border:1px solid #ccc; padding:5px;">Rp ${tr.toLocaleString('id-ID')}</td>
            <td style="border:1px solid #ccc; padding:5px;">Rp ${rl.toLocaleString('id-ID')}</td>
            <td style="border:1px solid #ccc; padding:5px; text-align:center;">${prc}%</td>
          </tr>`;
        });
      }
      
      let staffPerformaRows = '';
      if (data.weekly?.staff && data.weekly.staff.length > 0) {
        data.weekly.staff.forEach(st => {
          let icon = st.status === 'Menurun' ? '🔴' : (st.status === 'Stagnan' ? '🟡' : '🟢');
          staffPerformaRows += `<tr>
            <td style="border:1px solid #ccc; padding:5px;">${st.nama}</td>
            <td style="border:1px solid #ccc; padding:5px; text-align:center;">${icon} ${st.status}</td>
            <td style="border:1px solid #ccc; padding:5px;">${st.alasan || '-'}</td>
          </tr>`;
        });
      } else {
        staffPerformaRows = '<tr><td colspan="3" style="border:1px solid #ccc; padding:5px; text-align:center;">Tidak ada evaluasi staff minggu ini</td></tr>';
      }

      const komplain = data.weekly?.komplain?.total || 0;
      const kendalaUtama = data.weekly?.kendalaUtama || "-";
      
      const folder = getStructuredFolder('weekly', new Date());
      const fileName = `WeeklyReport_${periode}_${supervisor}.pdf`;
     
      const html = `<div style="font-family:sans-serif; padding: 20px;">
        <h2 style="border-bottom: 2px solid #000; padding-bottom:10px;">LAPORAN MINGGUAN - SUPERVISOR OUTLET ZERO ${outlet.toUpperCase()}</h2>
        <p><strong>Periode:</strong> ${periode}<br><strong>Supervisor:</strong> ${supervisor}</p>
        
        <h3 style="margin-top:30px; background:#eee; padding:5px;">A. REKAP TARGET SALES HARIAN</h3>
        <table style="width:100%; border-collapse:collapse; font-size:12px; text-align:left;">
          <tr style="background:#f9f9f9;">
            <th style="border:1px solid #ccc; padding:5px;">Hari</th>
            <th style="border:1px solid #ccc; padding:5px;">Target (Rp)</th>
            <th style="border:1px solid #ccc; padding:5px;">Real (Rp)</th>
            <th style="border:1px solid #ccc; padding:5px;">% Tercapai</th>
          </tr>
          ${salesHarianRows}
          <tr style="background:#f0f0f0; font-weight:bold;">
            <td colspan="2" style="border:1px solid #ccc; padding:5px; text-align:right;">Total Sales Realisasi:</td>
            <td colspan="2" style="border:1px solid #ccc; padding:5px;">Rp ${totalRealSales.toLocaleString('id-ID')}</td>
          </tr>
        </table>

        <h3 style="margin-top:20px; background:#eee; padding:5px;">B & C. EVALUASI PRODUK MINUMAN & MAKANAN</h3>
        <p style="font-size:12px;"><strong>Top 3 Minuman:</strong> 1. ${data.weekly?.produk?.topMinuman?.[0] || '-'} | 2. ${data.weekly?.produk?.topMinuman?.[1] || '-'} | 3. ${data.weekly?.produk?.topMinuman?.[2] || '-'}</p>
        <p style="font-size:12px; margin-bottom:10px;"><strong>Tindakan Bottom Minuman:</strong><br><span style="white-space:pre-wrap;">${data.weekly?.produk?.bottomMinuman?.[0] || '-'}</span></p>

        <h3 style="margin-top:20px; background:#eee; padding:5px;">D. KENDALA UTAMA BERULANG</h3>
        <p style="font-size:12px; padding:10px; border:1px solid #ccc; background:#fafafa; white-space:pre-wrap;">${kendalaUtama}</p>

        <h3 style="margin-top:20px; background:#eee; padding:5px;">E. PERFORMA BARISTA (Ringkasan Harian)</h3>
        <table style="width:100%; border-collapse:collapse; font-size:12px; text-align:left;">
          <tr style="background:#f9f9f9;">
            <th style="border:1px solid #ccc; padding:5px;">Nama Barista</th>
            <th style="border:1px solid #ccc; padding:5px;">Status</th>
            <th style="border:1px solid #ccc; padding:5px;">Alasan / Bukti</th>
          </tr>
          ${staffPerformaRows}
        </table>

        <h3 style="margin-top:20px; background:#eee; padding:5px;">F. REKAP KOMPLAIN & PENGGANTIAN</h3>
        <ul style="font-size:12px; padding-left:20px;">
          <li>Total Komplain: <strong>${komplain} Kasus</strong></li>
          <li>Total Penggantian (Remake): <strong>${data.weekly?.komplain?.remake || 0} Kasus</strong></li>
          <li>Penyebab Dominan: <strong>${data.weekly?.komplain?.penyebab || '-'}</strong></li>
        </ul>

        <h3 style="margin-top:20px; background:#eee; padding:5px;">H & I. RENCANA PERBAIKAN & KEBUTUHAN GM</h3>
        <p style="font-size:12px;"><strong>Rencana Minggu Depan:</strong><br><span style="white-space:pre-wrap;">${data.weekly?.rencana || '-'}</span></p>
        <p style="font-size:12px; margin-top:10px;"><strong>Kebutuhan dari GM:</strong><br><span style="white-space:pre-wrap;">${data.weekly?.kebutuhan || '-'}</span></p>
      </div>`;
      const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF).setName(fileName);
      const file = folder.createFile(blob);
      sheet.appendRow([periode, supervisor, outlet, totalRealSales, 0, komplain, kendalaUtama, file.getUrl()]);
    }
    
    else if (reportType === 'monthly') {
      const sheet = ss.getSheetByName('Monthly');
      if (!sheet) throw new Error("Tab 'Monthly' tidak ditemukan di Spreadsheet.");
      
      const bulan = data.bulan || "-";
      const supervisor = data.supervisor || "-";
      const outlet = data.outlet || "Perintis";
      
      const totalSales = Number(data.monthly?.sales?.total) || 0;
      const targetSales = Number(data.monthly?.sales?.target) || 0;
      const persenTercapai = data.monthly?.sales?.persen || 0;
      const ratingKerja = data.monthly?.evaluasi?.ratingKerja || 0;
      
      const folder = getStructuredFolder('monthly', new Date());
      const fileName = `MonthlyReport_${bulan}_${supervisor}.pdf`;
     
      const html = `<div style="font-family: 'Helvetica Neue', sans-serif; padding: 40px; color: #222; max-width: 800px; margin: auto;">
        <div style="border-bottom: 4px solid #000; padding-bottom: 15px; margin-bottom: 30px;">
          <h1 style="margin: 0; font-size: 28px; text-transform: uppercase;">MONTHLY REPORT ZERO CAFE</h1>
          <p style="margin: 8px 0 0 0; font-size: 14px; text-transform: uppercase;">Bulan: <strong>${bulan}</strong> | SPV: <strong>${supervisor}</strong></p>
        </div>
        
        <h2 style="font-size: 14px; background-color: #eee; padding: 8px; border-left: 4px solid #000;">A. RINGKASAN EKSEKUTIF</h2>
        <p style="font-size: 12px;"><strong>3 Pencapaian:</strong><br><span style="white-space:pre-wrap;">${data.monthly?.ringkasan?.pencapaian || '-'}</span></p>
        <p style="font-size: 12px;"><strong>3 Masalah Utama:</strong><br><span style="white-space:pre-wrap;">${data.monthly?.ringkasan?.masalah || '-'}</span></p>
        
        <h2 style="font-size: 14px; background-color: #eee; padding: 8px; border-left: 4px solid #000; margin-top: 20px;">B. METRIK PENJUALAN</h2>
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 10px; font-size:12px;">
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd; width:50%;"><strong>Total Sales:</strong> Rp ${totalSales.toLocaleString('id-ID')}</td>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Target:</strong> Rp ${targetSales.toLocaleString('id-ID')}</td>
          </tr>
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Total Transaksi:</strong> ${data.monthly?.sales?.transaksi || 0}</td>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Persentase Tercapai:</strong> ${persenTercapai}%</td>
          </tr>
          <tr>
            <td colspan="2" style="padding: 8px; border: 1px solid #ddd;"><strong>Rata-rata Sales / Hari:</strong> Rp ${Number(data.monthly?.sales?.rataRata || 0).toLocaleString('id-ID')}</td>
          </tr>
        </table>
        <p style="font-size:12px;"><strong>Top 5 Produk Terlaris:</strong><br><span style="white-space:pre-wrap;">${data.monthly?.produk?.top?.[0] || '-'}</span></p>
        
        <h2 style="font-size: 14px; background-color: #eee; padding: 8px; border-left: 4px solid #000; margin-top: 20px;">C & D. OPERASIONAL & KUALITAS (QC)</h2>
        <table style="width: 100%; border-collapse: collapse; margin-bottom: 10px; font-size:12px;">
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd; width:50%;"><strong>Kepatuhan SOP:</strong> ${data.monthly?.operasional?.kepatuhanSop || 0}%</td>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Total Kejadian Telat:</strong> ${data.monthly?.operasional?.telat || 0} Kali</td>
          </tr>
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Coaching 1-on-1:</strong> ${data.monthly?.operasional?.coaching || 0} Kali</td>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Penyebab Dominan:</strong> ${data.monthly?.operasional?.penyebab || '-'}</td>
          </tr>
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Total Komplain:</strong> ${data.monthly?.qc?.komplain || 0}</td>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>Total Remake:</strong> ${data.monthly?.qc?.remake || 0}</td>
          </tr>
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>QC Espresso (Sesuai):</strong> ${data.monthly?.qc?.espresso || 0}%</td>
            <td style="padding: 8px; border: 1px solid #ddd;"><strong>QC Susu Steam (Sesuai):</strong> ${data.monthly?.qc?.susu || 0}%</td>
          </tr>
        </table>

        <h2 style="font-size: 14px; background-color: #eee; padding: 8px; border-left: 4px solid #000; margin-top:20px;">F. RENCANA BULAN DEPAN</h2>
        <p style="font-size:12px;"><strong>Fokus Strategi:</strong><br><span style="white-space:pre-wrap;">${data.monthly?.rencana?.strategi || '-'}</span></p>
        <p style="font-size:12px;"><strong>Kebutuhan GM:</strong><br><span style="white-space:pre-wrap;">${data.monthly?.rencana?.gm || '-'}</span></p>

        <h2 style="font-size: 14px; background-color: #eee; padding: 8px; border-left: 4px solid #000; margin-top:20px;">G. EVALUASI DIRI SUPERVISOR</h2>
        <p style="font-size:12px;"><strong>Paling Berhasil:</strong> ${data.monthly?.evaluasi?.berhasil || '-'}</p>
        <p style="font-size:12px;"><strong>Paling Sulit:</strong> ${data.monthly?.evaluasi?.sulit || '-'}</p>
        <h3 style="text-align:right; font-size:16px;">Rating Kepuasan Kerja: ${ratingKerja}/10</h3>
      </div>`;
      
      const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF).setName(fileName);
      const file = folder.createFile(blob);
      sheet.appendRow([bulan, supervisor, outlet, totalSales, targetSales, persenTercapai, ratingKerja, file.getUrl()]);
    }
    
    return { success: true, message: "Laporan tersimpan rapi!" };
  } catch (error) {
    return { success: false, error: "System Error: " + error.toString() };
  }
}

function api_gm_fetchReports(monthName, year) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const dailySheet = ss.getSheetByName('Daily');
    let omsetTotal = 0;
    let komplainTotal = 0;
    let listLaporan = [];
    const monthMap = { "Januari": "01", "Februari": "02", "Maret": "03", "April": "04", "Mei": "05", "Juni": "06", "Juli": "07", "Agustus": "08", "September": "09", "Oktober": "10", "November": "11", "Desember": "12" };
    const targetPrefix = `${year}-${monthMap[monthName]}`;
    if (dailySheet) {
      const data = dailySheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const tanggalStr = row[0] ? row[0].toString() : "";
        if (tanggalStr.startsWith(targetPrefix)) {
          omsetTotal += Number(row[4]) || 0; 
          komplainTotal += Number(row[5]) || 0; 
          const urlPdf = row[7] ? row[7].toString() : ""; 
          if (urlPdf && urlPdf.includes('http')) {
            listLaporan.push({ name: `Harian_${tanggalStr}_${row[1]}`, url: urlPdf, dateCreated: tanggalStr });
          }
        }
      }
    }
    return { status: 'success', data: { omsetTotal: omsetTotal, transaksiTotal: listLaporan.length, komplainTotal: komplainTotal, listLaporan: listLaporan.reverse() }};
  } catch (error) {
    return { status: 'error', error: error.toString() };
  }
}
