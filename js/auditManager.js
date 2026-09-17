/**
 * auditManager.js
 * Modul untuk Audit & Sinkronisasi Pasca Ujian (CAT BKN)
 * Mencocokkan data NIP dari file Excel hasil sistem pasca ujian dengan data database eksisting:
 * 1. Mendeteksi dan memperbarui nomor sesi yang bertambah secara akumulasi.
 * 2. Mengubah otomatis status presensi menjadi 'HADIR' apabila kolom Login terisi tanggal/waktu valid.
 * 3. Mendeteksi perbedaan pada kolom jabatan, kelompok jabatan, unit kerja, jenis tes, dan waktu selesai.
 * 4. Menyediakan antarmuka komparasi ala Windows dengan checkbox individual (Keep vs Replace) dan Select All.
 */

import { 
    calculateCumulativeSessionNumber, 
    convertCumulativeSessionToDaily,
    parseFlexibleDate,
    formatDateDisplay,
    getSessionTime,
    formatCumulativeSessionNumber,
    getMaxSessionsForDate
} from './sessionRules.js';


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
    login: ['login', 'waktu login', 'jam login', 'tgl login', 'login time', 'waktu_login', 'login_time'],
    selesai: ['selesai', 'waktu selesai', 'jam selesai', 'tgl selesai', 'selesai time', 'waktu_selesai', 'selesai_time']
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
 * Normalisasi dan validasi nilai tanggal & waktu ke format standar: "DD MMMM YYYY HH:mm:ss"
 * Mengembalikan null jika nilainya kosong, "-", "null", atau "undefined".
 */
export function normalizeAuditDateTime(val) {
    if (val === undefined || val === null) return null;
    if (val instanceof Date) {
        if (isNaN(val.getTime())) return null;
        const pad = (n) => String(n).padStart(2, '0');
        const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
        return `${pad(val.getDate())} ${months[val.getMonth()]} ${val.getFullYear()} ${pad(val.getHours())}:${pad(val.getMinutes())}:${pad(val.getSeconds())}`;
    }

    const str = String(val).trim();
    if (!str || str === '-' || str.toLowerCase() === 'null' || str.toLowerCase() === 'undefined') {
        return null;
    }

    const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

    // 1. Format teks Indonesia: "DD MMMM YYYY HH:mm:ss" atau "DD MMMM YYYY HH:mm"
    const indPattern = /^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/;
    const indMatch = str.match(indPattern);
    if (indMatch) {
        const d = parseInt(indMatch[1], 10);
        const mStr = indMatch[2].toLowerCase();
        const y = parseInt(indMatch[3], 10);
        const hh = parseInt(indMatch[4] || '0', 10);
        const mm = parseInt(indMatch[5] || '0', 10);
        const ss = parseInt(indMatch[6] || '0', 10);

        const mIdx = months.findIndex(m => m.toLowerCase().startsWith(mStr.slice(0, 3)));
        if (mIdx !== -1) {
            const pad = (n) => String(n).padStart(2, '0');
            return `${pad(d)} ${months[mIdx]} ${y} ${pad(hh)}:${pad(mm)}:${pad(ss)}`;
        }
    }

    // 2. Format ISO / "YYYY-MM-DD HH:mm:ss" atau "YYYY/MM/DD HH:mm:ss"
    const ymdPattern = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[T\s](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/;
    const ymdMatch = str.match(ymdPattern);
    if (ymdMatch) {
        const y = parseInt(ymdMatch[1], 10);
        const m = parseInt(ymdMatch[2], 10) - 1;
        const d = parseInt(ymdMatch[3], 10);
        const hh = parseInt(ymdMatch[4] || '0', 10);
        const mm = parseInt(ymdMatch[5] || '0', 10);
        const ss = parseInt(ymdMatch[6] || '0', 10);
        if (m >= 0 && m < 12) {
            const pad = (n) => String(n).padStart(2, '0');
            return `${pad(d)} ${months[m]} ${y} ${pad(hh)}:${pad(mm)}:${pad(ss)}`;
        }
    }

    // 3. Format "DD/MM/YYYY HH:mm:ss" atau "DD-MM-YYYY HH:mm:ss"
    const dmyPattern = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})(?:[T\s](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/;
    const dmyMatch = str.match(dmyPattern);
    if (dmyMatch) {
        const d = parseInt(dmyMatch[1], 10);
        const m = parseInt(dmyMatch[2], 10) - 1;
        const y = parseInt(dmyMatch[3], 10);
        const hh = parseInt(dmyMatch[4] || '0', 10);
        const mm = parseInt(dmyMatch[5] || '0', 10);
        const ss = parseInt(dmyMatch[6] || '0', 10);
        if (m >= 0 && m < 12) {
            const pad = (n) => String(n).padStart(2, '0');
            return `${pad(d)} ${months[m]} ${y} ${pad(hh)}:${pad(mm)}:${pad(ss)}`;
        }
    }

    // 4. Coba parsing menggunakan native Date
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
        const pad = (n) => String(n).padStart(2, '0');
        return `${pad(parsed.getDate())} ${months[parsed.getMonth()]} ${parsed.getFullYear()} ${pad(parsed.getHours())}:${pad(parsed.getMinutes())}:${pad(parsed.getSeconds())}`;
    }

    return str;
}

/**
 * Format nilai tanggal/waktu dari cell Excel
 */
function formatCellDateTime(val) {
    if (val === undefined || val === null) return '';
    if (val instanceof Date) {
        if (isNaN(val.getTime())) return '';
        const pad = (n) => String(n).padStart(2, '0');
        const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
        return `${pad(val.getDate())} ${months[val.getMonth()]} ${val.getFullYear()} ${pad(val.getHours())}:${pad(val.getMinutes())}:${pad(val.getSeconds())}`;
    }
    const str = String(val).trim();
    if (!str || str === '-' || str.toLowerCase() === 'null' || str.toLowerCase() === 'undefined') return '';
    return normalizeAuditDateTime(str) || str;
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

        // 1. Cek Tanggal Pelaksanaan dari Waktu Login (Login Time)
        const normLogin = normalizeAuditDateTime(row.login);
        const normSelesai = normalizeAuditDateTime(row.selesai);
        const normDbLogin = normalizeAuditDateTime(existingCand.loginTime);
        const normDbSelesai = normalizeAuditDateTime(existingCand.selesaiTime);

        let loginDateFormatted = null;
        if (normLogin) {
            const loginDateObj = parseFlexibleDate(normLogin);
            if (loginDateObj) {
                loginDateFormatted = formatDateDisplay(loginDateObj, 'short');
            }
        }

        // Target tanggal pelaksanaan diutamakan dari tanggal pada login time
        let targetPelaksanaan = loginDateFormatted || (existingCand.pelaksanaan && existingCand.pelaksanaan !== 'NULL' && existingCand.pelaksanaan !== '-' ? existingCand.pelaksanaan : null);

        // Jika ada login time yang valid dan tanggal pelaksanaan di database berbeda / masih NULL, catat perubahan
        if (loginDateFormatted && existingCand.pelaksanaan !== loginDateFormatted) {
            changes.push({
                field: 'pelaksanaan',
                label: 'Tanggal Pelaksanaan',
                category: 'sesi',
                oldValue: existingCand.pelaksanaan || 'NULL',
                newValue: loginDateFormatted,
                rawNewValue: loginDateFormatted,
                reason: `Tanggal pelaksanaan diambil dari waktu login: ${normLogin}`
            });
        }

        // 2. Cek Sesi dan Waktu Ujian dari Kolom Sesi
        if (row.sesi !== null && row.sesi !== undefined && String(row.sesi).trim() !== '') {
            const rawSesiNum = Number(row.sesi);
            if (!isNaN(rawSesiNum) && rawSesiNum > 0) {
                // Jika targetPelaksanaan belum ada, coba dapatkan dari konversi sesi akumulasi
                if (!targetPelaksanaan && Array.isArray(sortedDates) && sortedDates.length > 0) {
                    const conv = convertCumulativeSessionToDaily(rawSesiNum, sortedDates);
                    targetPelaksanaan = conv.targetPelaksanaan;
                }

                let dayIdx = -1;
                if (targetPelaksanaan && Array.isArray(sortedDates)) {
                    dayIdx = sortedDates.indexOf(targetPelaksanaan);
                }

                let preceding = 0;
                if (dayIdx !== -1) {
                    for (let i = 0; i < dayIdx; i++) {
                        preceding += getMaxSessionsForDate(sortedDates[i]);
                    }
                }

                let dailySesi = rawSesiNum;
                let newCumSesi = rawSesiNum;

                if (dayIdx !== -1) {
                    const maxToday = getMaxSessionsForDate(targetPelaksanaan);
                    if (rawSesiNum > maxToday) {
                        // row.sesi adalah nomor sesi akumulasi (misal 4, 5, 6...)
                        newCumSesi = rawSesiNum;
                        dailySesi = Math.max(1, Math.min(rawSesiNum - preceding, maxToday));
                    } else {
                        // row.sesi adalah nomor sesi harian (1, 2, atau 3)
                        dailySesi = rawSesiNum;
                        newCumSesi = preceding + dailySesi;
                    }
                } else if (Array.isArray(sortedDates) && sortedDates.length > 0) {
                    const conv = convertCumulativeSessionToDaily(rawSesiNum, sortedDates);
                    dailySesi = conv.targetDailySesi;
                    newCumSesi = rawSesiNum;
                    targetPelaksanaan = targetPelaksanaan || conv.targetPelaksanaan;
                }

                const curCumSesi = calculateCumulativeSessionNumber(existingCand, sortedDates);
                const curDailySesi = Number(existingCand.sesi) || null;

                // Cek perubahan sesi (baik sesi harian, sesi akumulasi, ataupun jika pelaksanaan sebelumnya masih NULL)
                if (curCumSesi !== newCumSesi || curDailySesi !== dailySesi || !existingCand.pelaksanaan || existingCand.pelaksanaan === 'NULL') {
                    changes.push({
                        field: 'sesi',
                        label: 'Sesi Akumulasi',
                        category: 'sesi',
                        oldValue: curCumSesi !== null ? `Sesi ${curDailySesi || curCumSesi} (Akumulasi ${formatCumulativeSessionNumber(curCumSesi)})` : 'Belum Terjadwal (Sesi 00)',
                        newValue: `Sesi ${dailySesi} (Akumulasi ${formatCumulativeSessionNumber(newCumSesi)})`,
                        rawNewValue: dailySesi,
                        targetPelaksanaan: targetPelaksanaan || existingCand.pelaksanaan,
                        newCumulativeSesi: newCumSesi,
                        reason: `Penyesuaian sesi dari audit pasca ujian`
                    });
                    sesiChangedCount++;
                }

                // Baris waktu diisikan terbaru dengan mengikuti kolom sesi
                const standardTime = getSessionTime(dailySesi, targetPelaksanaan || existingCand.pelaksanaan);
                if (standardTime && standardTime !== '-' && existingCand.waktu !== standardTime) {
                    changes.push({
                        field: 'waktu',
                        label: 'Waktu Ujian',
                        category: 'waktu',
                        oldValue: existingCand.waktu || 'NULL',
                        newValue: standardTime,
                        rawNewValue: standardTime,
                        reason: `Waktu pelaksanaan disesuaikan untuk Sesi ${dailySesi}`
                    });
                }
            }
        }

        // 3. Cek Kehadiran berdasarkan Kolom Login
        // Jika kolom Login terisi tanggal/waktu yang valid dan status saat ini BELUM atau TIDAK_HADIR
        const hasValidLogin = Boolean(normLogin);
        const curKehadiran = String(existingCand.kehadiran || 'BELUM').toUpperCase();

        if (hasValidLogin && curKehadiran !== 'HADIR') {
            changes.push({
                field: 'kehadiran',
                label: 'Status Presensi',
                category: 'kehadiran',
                oldValue: curKehadiran === 'TIDAK_HADIR' ? 'TIDAK HADIR' : 'BELUM PRESENSI',
                newValue: 'HADIR',
                rawNewValue: 'HADIR',
                reason: `Terdeteksi waktu login CAT: ${normLogin}`
            });
            autoHadirCount++;
        }

        // Simpan waktu login jika ada timestamp baru yang valid dan berbeda dengan database
        if (normLogin && (!normDbLogin || normLogin !== normDbLogin)) {
            changes.push({
                field: 'loginTime',
                label: 'Waktu Login',
                category: 'waktu',
                oldValue: normDbLogin || existingCand.loginTime || '-',
                newValue: normLogin,
                rawNewValue: normLogin,
                reason: 'Data timestamp login sistem'
            });
        }

        // Simpan waktu selesai jika ada timestamp baru yang valid dan berbeda dengan database
        if (normSelesai && (!normDbSelesai || normSelesai !== normDbSelesai)) {
            changes.push({
                field: 'selesaiTime',
                label: 'Waktu Selesai',
                category: 'waktu',
                oldValue: normDbSelesai || existingCand.selesaiTime || '-',
                newValue: normSelesai,
                rawNewValue: normSelesai,
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
                hasSesiChange: changes.some(c => c.field === 'sesi' || c.field === 'pelaksanaan' || c.field === 'waktu')
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

