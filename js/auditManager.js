/**
 * auditManager.js
 * Modul untuk Audit & Sinkronisasi Pasca Ujian (CAT BKN)
 * Mencocokkan data NIP dari file Excel hasil sistem pasca ujian dengan data database eksisting:
 * 1. Mendeteksi dan memperbarui nomor sesi yang bertambah secara akumulasi.
 * 2. Mengubah otomatis status presensi menjadi 'HADIR' apabila kolom Login terisi tanggal/waktu valid.
 * 3. Mendeteksi perbedaan pada kolom jabatan, kelompok jabatan, unit kerja, jenis tes, dan waktu selesai.
 * 4. Menyediakan antarmuka komparasi ala Windows dengan checkbox individual (Keep vs Replace) dan Select All.
 */

import { calculateCumulativeSessionNumber, convertCumulativeSessionToDaily } from './sessionRules.js';


// Mapping kolom fleksibel untuk file hasil sistem pasca ujian
const AUDIT_HEADER_ALIASES = {
    nip: ['nip', 'nip baru', 'nomor induk pegawai', 'nrp', 'nip_peserta'],
    nama: ['nama', 'nama lengkap', 'nama peserta', 'pegawai', 'nama pegawai'],
    sesi: ['sesi', 'sesi ujian', 'sesi ke', 'sesi pelaksanaan', 'tahap'],
    jabatan: ['jabatan', 'nama jabatan', 'posisi'],
    kelJabatan: ['kel jabatan', 'kel. jabatan', 'kelompok jabatan', 'kel_jabatan', 'kelompok'],
    namaInstansi: ['nama instansi', 'nama_instansi', 'instansi', 'pemerintah daerah', 'pemda', 'nama pemerintah daerah'],
    unitKerja: ['unit kerja', 'unit_kerja', 'opd', 'skpd', 'satker', 'bagian', 'bidang', 'dinas', 'badan'],
    jenisTes: ['jenis tes', 'jenis_tes', 'tes', 'jenis ujian'],
    login: ['login', 'waktu login', 'jam login', 'tgl login', 'login time', 'waktu_login'],
    selesai: ['selesai', 'waktu selesai', 'jam selesai', 'tgl selesai', 'selesai time', 'waktu_selesai']
};

/**
 * Normalisasi header kolom Excel
 */
function matchAuditHeader(rawHeader) {
    if (!rawHeader) return null;
    const clean = String(rawHeader).trim().toLowerCase().replace(/[:._\-/\\]/g, ' ').replace(/\s+/g, ' ').trim();

    for (const [key, aliases] of Object.entries(AUDIT_HEADER_ALIASES)) {
        if (aliases.some(a => a === clean)) return key;
    }

    // Heuristik parsial
    if (clean.includes('kel') && clean.includes('jabatan')) return 'kelJabatan';
    if (clean.includes('jenis') && clean.includes('tes')) return 'jenisTes';
    if (clean === 'nip' || clean.startsWith('nip ')) return 'nip';
    if (clean.includes('nama') && !clean.includes('instansi')) return 'nama';
    if (clean.includes('sesi')) return 'sesi';
    if (clean.includes('login')) return 'login';
    if (clean.includes('selesai')) return 'selesai';
    if (clean.includes('instansi')) return 'namaInstansi';
    if (clean.includes('unit') || clean.includes('kerja') || clean.includes('opd') || clean.includes('skpd') || clean.includes('satker')) return 'unitKerja';
    if (clean.includes('jabatan')) return 'jabatan';

    return null;
}

/**
 * Format nilai tanggal/waktu dari cell Excel
 */
function formatCellDateTime(val) {
    if (val === undefined || val === null) return '';
    if (val instanceof Date) {
        // Format: DD MMMM YYYY HH:mm:ss atau YYYY-MM-DD HH:mm:ss
        const pad = (n) => String(n).padStart(2, '0');
        const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
        const d = val.getDate();
        const m = months[val.getMonth()];
        const y = val.getFullYear();
        const hh = pad(val.getHours());
        const mm = pad(val.getMinutes());
        const ss = pad(val.getSeconds());
        return `${pad(d)} ${m} ${y} ${hh}:${mm}:${ss}`;
    }
    return String(val).trim();
}

/**
 * Parsing file Excel Pasca Ujian
 * @param {File} file Objek File Excel
 * @returns {Promise<Array<Object>>} Array baris audit terstruktur
 */
export async function parseAuditExcel(file) {
    if (typeof XLSX === 'undefined') {
        throw new Error("Library SheetJS (XLSX) belum dimuat pada sistem.");
    }

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true, cellText: false });

    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) throw new Error("File Excel tidak memiliki sheet yang valid.");

    const sheet = workbook.Sheets[firstSheetName];
    const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    if (!rawData || rawData.length < 2) {
        throw new Error("File Excel kosong atau tidak memiliki baris data.");
    }

    // Temukan baris header
    let headerRowIdx = -1;
    let colMap = {};

    for (let r = 0; r < Math.min(10, rawData.length); r++) {
        const row = rawData[r];
        if (!Array.isArray(row)) continue;

        const currentMap = {};
        row.forEach((cell, colIdx) => {
            const matchedKey = matchAuditHeader(cell);
            if (matchedKey && !currentMap[matchedKey]) {
                currentMap[matchedKey] = colIdx;
            }
        });

        // Wajib memiliki NIP dan Nama atau Sesi/Login
        if (currentMap.nip !== undefined && (currentMap.nama !== undefined || currentMap.sesi !== undefined || currentMap.login !== undefined)) {
            headerRowIdx = r;
            colMap = currentMap;
            break;
        }
    }

    if (headerRowIdx === -1 || colMap.nip === undefined) {
        throw new Error("Kolom NIP tidak ditemukan pada file Excel. Pastikan format tabel memiliki kolom NIP, Nama, Sesi, Login, dll.");
    }

    const auditRows = [];
    for (let r = headerRowIdx + 1; r < rawData.length; r++) {
        const row = rawData[r];
        if (!row || row.length === 0) continue;

        const rawNip = String(row[colMap.nip] || '').trim();
        if (!rawNip || rawNip === '-' || rawNip.toLowerCase() === 'undefined' || rawNip.toLowerCase() === 'null') {
            continue; // Lewati baris kosong
        }

        const rawNama = colMap.nama !== undefined ? String(row[colMap.nama] || '').trim() : '';
        const rawSesi = colMap.sesi !== undefined ? row[colMap.sesi] : '';
        const rawJabatan = colMap.jabatan !== undefined ? String(row[colMap.jabatan] || '').trim() : '';
        const rawKelJabatan = colMap.kelJabatan !== undefined ? String(row[colMap.kelJabatan] || '').trim() : '';
        const rawInstansi = colMap.namaInstansi !== undefined ? String(row[colMap.namaInstansi] || '').trim() : '';
        const rawUnitKerja = colMap.unitKerja !== undefined ? String(row[colMap.unitKerja] || '').trim() : '';
        const rawJenisTes = colMap.jenisTes !== undefined ? String(row[colMap.jenisTes] || '').trim() : '';
        const rawLogin = colMap.login !== undefined ? formatCellDateTime(row[colMap.login]) : '';
        const rawSelesai = colMap.selesai !== undefined ? formatCellDateTime(row[colMap.selesai]) : '';

        // Ekstrak nomor sesi yang valid
        let sesiNum = null;
        if (rawSesi !== '' && rawSesi !== null && rawSesi !== undefined) {
            const parsedSesi = parseInt(String(rawSesi).replace(/\D/g, ''), 10);
            if (!isNaN(parsedSesi)) {
                sesiNum = parsedSesi;
            }
        }

        auditRows.push({
            rowNumber: r + 1,
            nip: rawNip,
            nama: rawNama,
            sesi: sesiNum,
            jabatan: rawJabatan,
            kelJabatan: rawKelJabatan,
            instansi: rawInstansi,
            unitKerja: rawUnitKerja,
            jenisTes: rawJenisTes,
            login: rawLogin,
            selesai: rawSelesai
        });
    }

    return auditRows;
}

/**
 * Helper untuk menghitung nomor sesi akumulasi/kumulatif kandidat database
 * Menghubungkan hari pelaksanaan dengan nomor sesi harian (Hari Jumat otomatis 2 sesi)
 */
function calculateCandidateCumulativeSession(cand, sortedDates) {
    return calculateCumulativeSessionNumber(cand, sortedDates);
}

/**
 * Membandingkan data Excel pasca ujian dengan data database peserta yang aktif
 * @param {Array<Object>} auditRows Data baris dari Excel
 * @param {Array<Object>} currentCandidates Data peserta database ujian aktif
 * @param {Object} options Opsi tambahan termasuk sortedDates untuk perhitungan sesi kumulatif
 * @returns {Object} Hasil komparasi lengkap
 */
export function compareAuditDataWithDatabase(auditRows, currentCandidates, options = {}) {
    if (!Array.isArray(auditRows) || !Array.isArray(currentCandidates)) {
        return {
            totalExcel: 0,
            matchedCount: 0,
            unmatchedCount: 0,
            diffCount: 0,
            diffCandidates: [],
            unmatchedExcelRows: []
        };
    }

    const sortedDates = options.sortedDates || [];

    // Buat lookup map untuk kandidat database berdasarkan NIP
    const candidatesByNip = new Map();
    const candidatesByCleanNip = new Map();

    currentCandidates.forEach(cand => {
        const rawNip = String(cand.nip || cand.id || '').trim();
        const cleanNip = rawNip.replace(/\D/g, '');
        if (rawNip) candidatesByNip.set(rawNip, cand);
        if (cleanNip) candidatesByCleanNip.set(cleanNip, cand);
    });

    const diffCandidates = [];
    const unmatchedExcelRows = [];
    let matchedCount = 0;
    let autoHadirCount = 0;
    let sesiChangedCount = 0;

    auditRows.forEach(row => {
        const rawNip = String(row.nip).trim();
        const cleanNip = rawNip.replace(/\D/g, '');

        const existingCand = candidatesByNip.get(rawNip) || (cleanNip ? candidatesByCleanNip.get(cleanNip) : null);

        if (!existingCand) {
            unmatchedExcelRows.push(row);
            return;
        }

        matchedCount++;
        const changes = [];

        // 1. Cek Sesi (Membandingkan SESI AKUMULASI / KUMULATIF bukan sesi harian, otomatis deteksi Jumat = 2 sesi)
        if (row.sesi !== null && row.sesi !== undefined) {
            const newCumSesi = Number(row.sesi);
            const curCumSesi = calculateCumulativeSessionNumber(existingCand, sortedDates);
            
            if (curCumSesi !== newCumSesi) {
                // Tentukan target sesi harian (1, 2, atau 3) dan tanggal pelaksanaan jika ada sortedDates
                const { targetPelaksanaan, targetDailySesi, scheduleDetail } = convertCumulativeSessionToDaily(newCumSesi, sortedDates);

                changes.push({
                    field: 'sesi',
                    label: 'Sesi Akumulasi',
                    category: 'sesi',
                    oldValue: curCumSesi !== null ? `Sesi Akumulasi ${curCumSesi}` : 'Belum Terjadwal (Sesi 00)',
                    newValue: `Sesi Akumulasi ${newCumSesi}`,
                    rawNewValue: targetDailySesi, // Nilai sesi harian (1, 2, 3) yang aman untuk database
                    targetPelaksanaan: targetPelaksanaan || existingCand.pelaksanaan, // Tanggal pelaksanaan hasil penyesuaian sesi akumulasi
                    newCumulativeSesi: newCumSesi,
                    reason: `Penyesuaian sesi akumulasi pasca ujian: ${scheduleDetail}`
                });
                sesiChangedCount++;
            }
        }

        // 2. Cek Kehadiran berdasarkan Kolom Login
        // Jika kolom Login terisi tanggal/waktu yang valid dan status saat ini BELUM atau TIDAK_HADIR
        const hasValidLogin = Boolean(row.login && row.login !== '-' && row.login !== 'NULL' && row.login !== 'null');
        const curKehadiran = String(existingCand.kehadiran || 'BELUM').toUpperCase();

        if (hasValidLogin && curKehadiran !== 'HADIR') {
            changes.push({
                field: 'kehadiran',
                label: 'Status Presensi',
                category: 'kehadiran',
                oldValue: curKehadiran === 'TIDAK_HADIR' ? 'TIDAK HADIR' : 'BELUM PRESENSI',
                newValue: 'HADIR',
                rawNewValue: 'HADIR',
                reason: `Terdeteksi waktu login CAT: ${row.login}`
            });
            autoHadirCount++;
        }

        // Simpan waktu login & selesai jika ada
        if (row.login && existingCand.loginTime !== row.login) {
            changes.push({
                field: 'loginTime',
                label: 'Waktu Login',
                category: 'waktu',
                oldValue: existingCand.loginTime || '-',
                newValue: row.login,
                rawNewValue: row.login,
                reason: 'Data timestamp login sistem'
            });
        }

        if (row.selesai && existingCand.selesaiTime !== row.selesai) {
            changes.push({
                field: 'selesaiTime',
                label: 'Waktu Selesai',
                category: 'waktu',
                oldValue: existingCand.selesaiTime || '-',
                newValue: row.selesai,
                rawNewValue: row.selesai,
                reason: 'Data timestamp selesai ujian'
            });
        }

        // 3. Cek Jabatan
        if (row.jabatan && existingCand.jabatan && row.jabatan.toLowerCase() !== existingCand.jabatan.trim().toLowerCase()) {
            changes.push({
                field: 'jabatan',
                label: 'Jabatan',
                category: 'biodata',
                oldValue: existingCand.jabatan,
                newValue: row.jabatan,
                rawNewValue: row.jabatan,
                reason: 'Pembaruan nama jabatan dari master pasca ujian'
            });
        }

        // 4. Cek Kelompok Jabatan
        if (row.kelJabatan && existingCand.kelJabatan && row.kelJabatan.toLowerCase() !== existingCand.kelJabatan.trim().toLowerCase()) {
            changes.push({
                field: 'kelJabatan',
                label: 'Kelompok Jabatan',
                category: 'biodata',
                oldValue: existingCand.kelJabatan,
                newValue: row.kelJabatan,
                rawNewValue: row.kelJabatan,
                reason: 'Pembaruan kategori kelompok jabatan'
            });
        }

        // 5. Cek Unit Kerja (Khusus jika file Excel memiliki kolom unit kerja spesifik)
        // PENTING: Jangan bandingkan kolom Nama Instansi pemerintah daerah dengan Unit Kerja peserta!
        // Unit Kerja pada jadwal adalah organisasi/bagian/bidang/dinas di dalam pemda (misal: "DINAS KESEHATAN", "BAGIAN HUKUM").
        // Sedangkan Nama Instansi adalah entitas pemerintah daerah (misal: "Pemerintah Kab. Teluk Wondama").
        if (row.unitKerja) {
            const cleanNewUnit = String(row.unitKerja).trim();
            const cleanCurUnit = String(existingCand.unitKerja || '').trim();
            const isPemdaName = cleanNewUnit.toLowerCase().includes('pemerintah') || 
                                cleanNewUnit.toLowerCase().includes('pemkab') || 
                                cleanNewUnit.toLowerCase().includes('pemkot') ||
                                cleanNewUnit.toLowerCase().includes('provinsi');

            if (cleanNewUnit && !isPemdaName && cleanCurUnit && cleanNewUnit.toLowerCase() !== cleanCurUnit.toLowerCase()) {
                changes.push({
                    field: 'unitKerja',
                    label: 'Unit Kerja',
                    category: 'biodata',
                    oldValue: existingCand.unitKerja,
                    newValue: cleanNewUnit,
                    rawNewValue: cleanNewUnit,
                    reason: 'Pembaruan nama unit kerja/bidang'
                });
            }
        }

        // 6. Cek Jenis Tes
        if (row.jenisTes && existingCand.jenisTes && row.jenisTes.toLowerCase() !== existingCand.jenisTes.trim().toLowerCase()) {
            changes.push({
                field: 'jenisTes',
                label: 'Jenis Tes',
                category: 'biodata',
                oldValue: existingCand.jenisTes || '-',
                newValue: row.jenisTes,
                rawNewValue: row.jenisTes,
                reason: 'Pembaruan jenis tes/skema ujian'
            });
        }

        if (changes.length > 0) {
            diffCandidates.push({
                nip: existingCand.nip || row.nip,
                nama: existingCand.nama || row.nama,
                existingCand,
                excelRow: row,
                changes,
                hasKehadiranChange: changes.some(c => c.field === 'kehadiran'),
                hasSesiChange: changes.some(c => c.field === 'sesi')
            });
        }
    });

    // Urutkan pegawai dengan jumlah perbedaan data terbanyak di atas, sampai yang paling sedikit
    diffCandidates.sort((a, b) => {
        const countA = a.changes ? a.changes.length : 0;
        const countB = b.changes ? b.changes.length : 0;
        if (countB !== countA) {
            return countB - countA;
        }
        return String(a.nama || '').localeCompare(String(b.nama || ''), 'id', { sensitivity: 'base' });
    });

    return {
        totalExcel: auditRows.length,
        matchedCount,
        unmatchedCount: unmatchedExcelRows.length,
        diffCount: diffCandidates.length,
        autoHadirCount,
        sesiChangedCount,
        diffCandidates,
        unmatchedExcelRows
    };
}

