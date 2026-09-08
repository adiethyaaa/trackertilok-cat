/**
 * excelHandler.js
 * Modul untuk membaca file Excel, parsing struktur kolom referensi,
 * kalkulasi otomatis sesi/waktu, pembuatan template Excel, export data,
 * dan deteksi duplikasi NIP (internal file Excel maupun terhadap database eksisting).
 */

import { parseFlexibleDate, isFriday, getSessionTime, formatDateDisplay } from './sessionRules.js';

// Header acuan standar sesuai instruksi user
export const REQUIRED_COLUMNS = [
    { key: 'no', label: 'No', aliases: ['no', 'nomor', 'no.', 'urut', 'num'] },
    { key: 'nip', label: 'NIP', aliases: ['nip', 'nip baru', 'nomor induk pegawai', 'nrp', 'nip_peserta', 'nip peserta'] },
    { key: 'nama', label: 'NAMA', aliases: ['nama', 'nama lengkap', 'nama peserta', 'pegawai', 'nama_pegawai', 'nama pegawai'] },
    { key: 'kelJabatan', label: 'KEL JABATAN', aliases: ['kel jabatan', 'kel. jabatan', 'kelompok jabatan', 'kel_jabatan', 'kelompok_jabatan', 'kelompok'] },
    { key: 'unitKerja', label: 'UNIT KERJA', aliases: ['unit kerja', 'unit_kerja', 'instansi / unit kerja', 'skpd', 'opd', 'bagian', 'satuan kerja', 'satker', 'unit', 'nama instansi', 'nama_instansi', 'instansi'] },
    { key: 'jabatan', label: 'JABATAN', aliases: ['jabatan', 'nama jabatan', 'posisi', 'jabatan sekarang'] },
    { key: 'waktu', label: 'WAKTU', aliases: ['waktu', 'jam', 'waktu ujian', 'pukul', 'jadwal', 'jam ujian', 'waktu pelaksanaan', 'jam pelaksanaan'] },
    { key: 'pelaksanaan', label: 'PELAKSANAAN', aliases: [
        'pelaksanaan', 'pelaksaan', 'pelaksana', 'tanggal', 'tgl', 
        'tgl pelaksanaan', 'tgl pelaksaan', 'tgl. pelaksanaan', 'tgl. pelaksaan', 
        'tanggal pelaksanaan', 'tanggal pelaksaan', 'jadwal pelaksanaan', 'jadwal pelaksaan',
        'hari/tanggal', 'hari / tanggal', 'hari, tanggal', 'hari tanggal', 'tanggal ujian', 'tgl ujian',
        'tgl_pelaksanaan', 'tgl_pelaksaan'
    ] },
    { key: 'sesi', label: 'SESI', aliases: ['sesi', 'sesi ujian', 'sesi ke', 'tahap', 'sesi_ujian', 'sesi pelaksanaan'] },
    { key: 'jenisTes', label: 'JENIS TES', aliases: ['jenis tes', 'jenis_tes', 'jenis ujian', 'tes'] }
];

/**
 * Mencocokkan nama kolom di file Excel dengan kolom sistem secara cerdas & fleksibel
 */
export function findMatchingKey(headerName) {
    if (!headerName) return null;
    const clean = String(headerName)
        .trim()
        .toLowerCase()
        .replace(/[:._\-/\\]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

    // 1. Cek pencocokan persis alias
    for (const col of REQUIRED_COLUMNS) {
        if (col.aliases.some(a => a.replace(/[:._\-/\\]/g, ' ') === clean)) {
            return col.key;
        }
    }

    // 2. Cek pencocokan parsial / keyword (Urutan diperhatikan!)
    if (clean.includes('kel') && clean.includes('jabatan')) {
        return 'kelJabatan';
    }
    if (clean.includes('jenis') && clean.includes('tes')) {
        return 'jenisTes';
    }
    if (clean.includes('pelaksaan') || clean.includes('pelaksanaan') || clean.includes('tanggal') || clean.includes('tgl')) {
        return 'pelaksanaan';
    }
    if (clean === 'nip' || clean.startsWith('nip ') || clean.includes('nomor induk')) {
        return 'nip';
    }
    if (clean.includes('nama') && !clean.includes('instansi')) {
        return 'nama';
    }
    if (clean.includes('sesi')) {
        return 'sesi';
    }
    if (clean.includes('waktu') || clean.includes('pukul') || (clean.includes('jam') && !clean.includes('jambatan'))) {
        return 'waktu';
    }
    if (clean.includes('unit') || clean.includes('kerja') || clean.includes('opd') || clean.includes('skpd') || clean.includes('satker') || clean.includes('instansi')) {
        return 'unitKerja';
    }
    if (clean.includes('jabatan')) {
        return 'jabatan';
    }
    if (clean === 'no' || clean === 'no.' || clean === 'nomor' || clean === 'urut') {
        return 'no';
    }

    return null;
}

/**
 * Membaca file Excel dan mengonversi ke array data peserta terstruktur.
 * Mendukung dua mode:
 * 1. Mode 'SYSTEM': Data peserta awal dari sistem (NIP, Nama, Jabatan, Kel Jabatan, Nama Instansi, Jenis Tes) -> Jadwal & Sesi diset NULL.
 * 2. Mode 'SCHEDULE': Data jadwal peserta (NIP, Pelaksanaan, Sesi, Waktu) -> Untuk melengkapi jadwal tanpa menimpa nama/jabatan sistem.
 * 
 * @param {File} file Objek File dari input
 * @param {Object} options Opsi { autoStandardizeTime: boolean, defaultPelaksanaan?: string, uploadMode?: 'AUTO'|'SYSTEM'|'SCHEDULE', defaultInstansi?: string }
 * @returns {Promise<Object>} { success, candidates, mode, summary }
 */
export async function parseExcelFile(file, options = { autoStandardizeTime: true, defaultPelaksanaan: '', uploadMode: 'AUTO', defaultInstansi: '' }) {
    if (typeof XLSX === 'undefined') {
        throw new Error("Library SheetJS (XLSX) belum dimuat.");
    }

    const dataBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(dataBuffer, { type: 'array', cellDates: true, cellText: true });

    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    // Baca ke format array 2D
    const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    if (!rawRows || rawRows.length < 2) {
        throw new Error("File Excel kosong atau tidak memiliki data.");
    }

    // Cari baris header (scan hingga 40 baris pertama)
    let headerRowIndex = -1;
    let headerMap = {};

    for (let r = 0; r < Math.min(rawRows.length, 40); r++) {
        const row = rawRows[r];
        if (!Array.isArray(row)) continue;

        let matchCount = 0;
        const currentMap = {};

        row.forEach((cell, colIdx) => {
            const key = findMatchingKey(cell);
            if (key) {
                currentMap[colIdx] = key;
                matchCount++;
            }
        });

        const hasNip = Object.values(currentMap).includes('nip');
        const hasNama = Object.values(currentMap).includes('nama');

        if ((hasNip && hasNama) || matchCount >= 3) {
            headerRowIndex = r;
            headerMap = currentMap;
            break;
        }
    }

    if (headerRowIndex === -1) {
        throw new Error("Format kolom Excel tidak dikenali. Pastikan terdapat kolom NIP dan NAMA.");
    }

    // Deteksi mode file: SYSTEM atau SCHEDULE
    const hasKelJabatan = Object.values(headerMap).includes('kelJabatan');
    const hasPelaksanaan = Object.values(headerMap).includes('pelaksanaan');

    let effectiveMode = options.uploadMode || 'AUTO';
    if (effectiveMode === 'AUTO') {
        if (hasKelJabatan || !hasPelaksanaan) {
            effectiveMode = 'SYSTEM';
        } else {
            effectiveMode = 'SCHEDULE';
        }
    }

    const parsedCandidates = [];
    let skippedCount = 0;

    // Scan setiap baris data
    for (let i = headerRowIndex + 1; i < rawRows.length; i++) {
        const row = rawRows[i];
        if (!row || !Array.isArray(row)) {
            skippedCount++;
            continue;
        }

        const rowObj = {
            no: '',
            nip: '',
            nama: '',
            kelJabatan: '-',
            unitKerja: options.defaultInstansi || '-',
            jabatan: '-',
            waktu: effectiveMode === 'SYSTEM' ? 'NULL' : '',
            pelaksanaan: effectiveMode === 'SYSTEM' ? 'NULL' : '',
            sesi: effectiveMode === 'SYSTEM' ? 'NULL' : 1,
            isFriday: false,
            status: effectiveMode === 'SYSTEM' ? 'Belum Terjadwal' : 'Terjadwal',
            jenisTes: '',
            kehadiran: null,
            excelRowNumber: i + 1
        };

        Object.keys(headerMap).forEach(colIdx => {
            const fieldKey = headerMap[colIdx];
            let val = row[colIdx];

            if (fieldKey === 'pelaksanaan' && effectiveMode !== 'SYSTEM') {
                const cellAddress = XLSX.utils.encode_cell({ r: i, c: Number(colIdx) });
                const directCell = worksheet[cellAddress];
                const rawCellText = (directCell && directCell.w) ? directCell.w.trim() : null;

                if (rawCellText) {
                    const parsedFromText = parseFlexibleDate(rawCellText);
                    rowObj.pelaksanaan = parsedFromText ? formatDateDisplay(parsedFromText, 'short') : rawCellText;
                } else if (val instanceof Date || typeof val === 'number') {
                    const parsed = parseFlexibleDate(val);
                    rowObj.pelaksanaan = parsed ? formatDateDisplay(parsed, 'short') : String(val);
                } else {
                    const strVal = String(val || '').trim();
                    const parsed = parseFlexibleDate(strVal);
                    rowObj.pelaksanaan = parsed ? formatDateDisplay(parsed, 'short') : strVal;
                }
            } else if (fieldKey !== 'pelaksanaan' || effectiveMode !== 'SYSTEM') {
                rowObj[fieldKey] = String(val !== undefined && val !== null ? val : '').trim();
            }
        });

        // Bersihkan NIP dari karakter petik, spasi, atau simbol
        const cleanNip = String(rowObj.nip || '').replace(/['"`\s]/g, '').trim();
        const cleanNama = String(rowObj.nama || '').replace(/['"`]/g, '').trim();

        // Validasi pemenuhan NIP dan NAMA
        if (!cleanNip || !cleanNama || cleanNip.length < 5 || cleanNama.length < 2) {
            skippedCount++;
            continue;
        }

        rowObj.nip = cleanNip;
        rowObj.nama = cleanNama;
        rowObj.no = parsedCandidates.length + 1;
        rowObj.kelJabatan = rowObj.kelJabatan && rowObj.kelJabatan !== '' ? rowObj.kelJabatan : '-';
        rowObj.jabatan = rowObj.jabatan && rowObj.jabatan !== '' ? rowObj.jabatan : '-';
        if (effectiveMode === 'SYSTEM') {
            // Pada file sistem: unit kerja dibiarkan 'NULL' (tidak mengambil nama instansi/template)
            rowObj.unitKerja = 'NULL';
            rowObj.pelaksanaan = 'NULL';
            rowObj.sesi = '00'; // Diberikan nilai 00 agar bisa dipilih dan dipanggil di filter
            rowObj.waktu = 'NULL';
            rowObj.status = 'Belum Terjadwal';
            rowObj.isFriday = false;
        } else {
            rowObj.unitKerja = rowObj.unitKerja && rowObj.unitKerja !== '' ? rowObj.unitKerja : (options.defaultInstansi || '-');
            // Mode SCHEDULE (Jadwal)
            let sesiNum = parseInt(rowObj.sesi, 10);
            if (isNaN(sesiNum) || sesiNum < 1 || sesiNum > 3) {
                sesiNum = 1;
            }
            rowObj.sesi = sesiNum;

            if (!rowObj.pelaksanaan && options.defaultPelaksanaan) {
                rowObj.pelaksanaan = options.defaultPelaksanaan;
            }

            if (rowObj.pelaksanaan && rowObj.pelaksanaan !== 'NULL') {
                const parsedDate = parseFlexibleDate(rowObj.pelaksanaan);
                if (parsedDate) {
                    rowObj.pelaksanaan = formatDateDisplay(parsedDate, 'short');
                    rowObj.isFriday = isFriday(parsedDate);
                } else {
                    rowObj.isFriday = false;
                }
            }

            const standardTime = getSessionTime(rowObj.sesi, rowObj.pelaksanaan);
            if (options.autoStandardizeTime || !rowObj.waktu || rowObj.waktu === '-' || rowObj.waktu.trim() === '') {
                rowObj.waktu = standardTime;
            }
            rowObj.status = 'Terjadwal';
        }

        parsedCandidates.push(rowObj);
    }

    return {
        success: parsedCandidates.length > 0,
        candidates: parsedCandidates,
        mode: effectiveMode,
        summary: {
            totalRows: parsedCandidates.length,
            skippedRows: skippedCount,
            mode: effectiveMode,
            sesi1: parsedCandidates.filter(c => Number(c.sesi) === 1).length,
            sesi2: parsedCandidates.filter(c => Number(c.sesi) === 2).length,
            sesi3: parsedCandidates.filter(c => Number(c.sesi) === 3).length,
            nullScheduleRows: parsedCandidates.filter(c => !c.sesi || c.sesi === 'NULL' || c.sesi === '00' || c.sesi === 0 || c.sesi === '0').length,
            fridayRows: parsedCandidates.filter(c => c.isFriday).length
        }
    };
}

/**
 * Deteksi dan Analisis Duplikasi NIP
 * Mendeteksi 2 kategori:
 * 1. Duplikasi internal di dalam file Excel yang sama (NIP yang muncul lebih dari 1 kali di Excel)
 * 2. Duplikasi dengan data database yang sudah ada (NIP dari Excel yang sudah terdaftar di database)
 * 
 * @param {Array} parsedCandidates Data peserta dari file Excel
 * @param {Array} existingCandidates Data peserta yang sudah tersimpan di database ujian ini
 * @returns {Object} { hasDuplicates, internalDuplicates, existingDuplicates, cleanCandidates }
 */
export function analyzeDuplicates(parsedCandidates, existingCandidates = []) {
    // Buat map data database eksisting berdasarkan NIP
    const existingMap = new Map();
    existingCandidates.forEach(c => {
        const clean = String(c.nip || '').trim();
        if (clean) existingMap.set(clean, c);
    });

    // Petakan frekuensi kemunculan NIP di dalam file Excel
    const excelNipMap = new Map();
    parsedCandidates.forEach(c => {
        const nip = c.nip;
        if (!excelNipMap.has(nip)) {
            excelNipMap.set(nip, []);
        }
        excelNipMap.get(nip).push(c);
    });

    const internalDuplicates = []; // Muncul > 1 kali di Excel
    const existingDuplicates = []; // Muncul di Excel dan sudah ada di Database
    const duplicateNips = new Set();

    // 1. Cek Duplikasi Internal di Excel
    excelNipMap.forEach((items, nip) => {
        if (items.length > 1) {
            duplicateNips.add(nip);
            internalDuplicates.push({
                nip,
                nama: items[0].nama,
                count: items.length,
                items: items.map((item, idx) => ({
                    ...item,
                    duplicateSource: `Excel (Baris ${item.excelRowNumber || idx + 1})`,
                    isFromDB: false,
                    uniqueKey: `excel_${nip}_${idx}`
                }))
            });
        }
    });

    // 2. Cek Duplikasi dengan Database Eksisting
    excelNipMap.forEach((items, nip) => {
        if (existingMap.has(nip)) {
            duplicateNips.add(nip);
            const dbItem = existingMap.get(nip);
            existingDuplicates.push({
                nip,
                nama: items[0].nama,
                existingItem: {
                    ...dbItem,
                    duplicateSource: 'Database Eksisting',
                    isFromDB: true,
                    uniqueKey: `db_${dbItem.id}`
                },
                incomingItems: items.map((item, idx) => ({
                    ...item,
                    duplicateSource: `Excel Baru (Baris ${item.excelRowNumber || idx + 1})`,
                    isFromDB: false,
                    uniqueKey: `excel_incoming_${nip}_${idx}`
                }))
            });
        }
    });

    // 3. Data Bersih (Tidak ada duplikasi sama sekali)
    const cleanCandidates = parsedCandidates.filter(c => !duplicateNips.has(c.nip));

    return {
        hasDuplicates: internalDuplicates.length > 0 || existingDuplicates.length > 0,
        internalDuplicates,
        existingDuplicates,
        cleanCandidates,
        totalDuplicateNips: duplicateNips.size
    };
}

/**
 * Merge Data Jadwal Ujian dengan Data Peserta Eksisting (Dari File Sistem)
 * ATURAN INTEGRITAS DATA:
 * - File sistem memegang otoritas penuh atas nama, jabatan, dan kelJabatan.
 * - File jadwal hanya melengkapi pelaksanaan, sesi, dan waktu.
 * - Data nama dan jabatan asli sistem TIDAK BOLEH ditimpa oleh data file jadwal.
 * 
 * @param {Array} incomingCandidates Data peserta dari file Excel jadwal
 * @param {Array} existingCandidates Data peserta yang sudah tersimpan di database
 * @returns {Object} { updatedCandidates, newCandidates, untouchedExisting, allMerged, matchedCount, newCount }
 */
export function mergeScheduleWithExisting(incomingCandidates, existingCandidates = []) {
    const existingMap = new Map();
    existingCandidates.forEach(c => {
        const clean = String(c.nip || '').trim();
        if (clean) existingMap.set(clean, { ...c });
    });

    const updatedCandidates = [];
    const newCandidates = [];
    const matchedNips = new Set();

    incomingCandidates.forEach((cand) => {
        const nip = String(cand.nip || '').trim();
        if (existingMap.has(nip)) {
            const ex = existingMap.get(nip);
            matchedNips.add(nip);
            // UPDATE JADWAL TAPI PROTEKSI NAMA, JABATAN, & KEL JABATAN ASLI SISTEM
            // UNIT KERJA DIAMBIL DARI FILE JADWAL PESERTA
            const resolvedUnitKerja = (cand.unitKerja && cand.unitKerja !== '-' && cand.unitKerja !== 'NULL' && cand.unitKerja.trim() !== '')
                ? cand.unitKerja
                : (ex.unitKerja && ex.unitKerja !== 'NULL' ? ex.unitKerja : 'NULL');

            const merged = {
                ...ex,
                nama: ex.nama, // NAMA SISTEM TETAP UTUH
                jabatan: ex.jabatan && ex.jabatan !== '-' ? ex.jabatan : (cand.jabatan || '-'), // JABATAN SISTEM UTUH
                kelJabatan: ex.kelJabatan && ex.kelJabatan !== '-' ? ex.kelJabatan : (cand.kelJabatan || '-'),
                unitKerja: resolvedUnitKerja, // DIISI DARI EXCEL JADWAL PESERTA
                pelaksanaan: cand.pelaksanaan,
                sesi: cand.sesi,
                waktu: cand.waktu,
                isFriday: cand.isFriday,
                status: 'Terjadwal'
            };
            updatedCandidates.push(merged);
        } else {
            newCandidates.push({
                ...cand,
                status: 'Terjadwal'
            });
        }
    });

    const untouchedExisting = existingCandidates.filter(c => !matchedNips.has(String(c.nip || '').trim()));
    const allMerged = [...updatedCandidates, ...newCandidates, ...untouchedExisting];
    allMerged.forEach((c, idx) => { c.no = idx + 1; });

    return {
        updatedCandidates,
        newCandidates,
        untouchedExisting,
        allMerged,
        matchedCount: updatedCandidates.length,
        newCount: newCandidates.length
    };
}

/**
 * Unduh template Excel Opsi 1: Data Master Peserta dari Sistem
 * Struktur kolom: NIP, Nama, Sesi, Jabatan, Kel Jabatan, Nama Instansi, Jenis Tes
 */
export function downloadSystemTemplate() {
    if (typeof XLSX === 'undefined') {
        alert("Library Excel belum selesai dimuat.");
        return;
    }

    const headers = ["NIP", "Nama", "Sesi", "Jabatan", "Kel Jabatan", "Nama Instansi", "Jenis Tes"];
    
    const sampleData = [
        ["197504202009041002", "DANIAL", "", "PENGADMINISTRASI PERKANTORAN", "Pelaksana", "Pemerintah Kab. Teluk Wondama", "Profiling Talenta ASN 2026"],
        ["197608022005021003", "ROBBY LAHUMETEN", "", "Kepala BIDANG MUTASI", "Administrator", "Pemerintah Kab. Teluk Wondama", "Profiling Talenta ASN 2026"],
        ["197702272009092001", "IVONE", "", "Kepala BIDANG INFORMASI DAN FORMASI", "Administrator", "Pemerintah Kab. Teluk Wondama", "Profiling Talenta ASN 2026"]
    ];

    const wsData = [headers, ...sampleData];
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    ws['!cols'] = [
        { wch: 24 }, // NIP
        { wch: 35 }, // Nama
        { wch: 8 },  // Sesi
        { wch: 35 }, // Jabatan
        { wch: 20 }, // Kel Jabatan
        { wch: 35 }, // Nama Instansi
        { wch: 30 }  // Jenis Tes
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Data Peserta Sistem");
    XLSX.writeFile(wb, "Template_Data_Sistem_Profiling_ASN.xlsx");
}

/**
 * Unduh template Excel Opsi 2: Data Jadwal Ujian Peserta
 * Struktur kolom: No, NIP, NAMA, UNIT KERJA, JABATAN, WAKTU, PELAKSANAAN, SESI
 */
export function downloadScheduleTemplate() {
    if (typeof XLSX === 'undefined') {
        alert("Library Excel belum selesai dimuat.");
        return;
    }

    const headers = ["No", "NIP", "NAMA", "UNIT KERJA", "JABATAN", "WAKTU", "PELAKSANAAN", "SESI"];
    
    const sampleData = [
        [1, "197809092014091001", "GARDEN SEMUEL KARUBUY", "", "ANALIS PENAGIHAN DAN PENGEMBALIAN", "08.00 - 11.00 WIT", "11-Sep-26", 1],
        [2, "197812052014091002", "DARIUS AKWAN", "SUB BAGIAN UMUM DAN KEPEGAWAIAN - DINAS LINGKUNGAN HIDUP", "PENGADMINISTRASI UMUM", "13.00 - 16.00 WIT", "11-Sep-26", 2],
        [3, "197902132015121001", "PIET ALFONS SPENNER WAROPEN", "BIDANG PERDAGANGAN - DINAS PERINDUSTRIAN, PERDAGANGAN DAN KOPERASI", "PENGADMINISTRASI UMUM", "14.00 - 17.00 WIT", "11-Sep-26", 3],
        [4, "197903082015011001", "YAN WILLEM DEREK AITO INURI", "KELURAHAN WASIOR - ASISTEN BIDANG PEMERINTAHAN DAN KESEJAHTERAAN RAKYAT", "TEKNIS/ADMINISTRASI LAINNYA", "08.00 - 11.00 WIT", "11-Sep-26", 1],
        [5, "197904092014091001", "JACOB SYEMI EDUARD LESSY", "SUB BAGIAN KEUANGAN DAN ASET - SEKRETARIAT", "PENGADMINISTRASI UMUM", "08.00 - 11.00 WIT", "11-Sep-26", 1]
    ];

    const wsData = [headers, ...sampleData];
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    ws['!cols'] = [
        { wch: 6 },  // No
        { wch: 24 }, // NIP
        { wch: 35 }, // NAMA
        { wch: 45 }, // UNIT KERJA
        { wch: 35 }, // JABATAN
        { wch: 22 }, // WAKTU
        { wch: 16 }, // PELAKSANAAN
        { wch: 8 }   // SESI
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Jadwal Ujian");
    XLSX.writeFile(wb, "Template_Jadwal_Profiling_ASN.xlsx");
}

/**
 * Alias untuk backwards-compatibility
 */
export function downloadExcelTemplate() {
    downloadScheduleTemplate();
}

/**
 * Export data peserta ujian ke file Excel dengan menyertakan Kel Jabatan
 */
export function exportCandidatesToExcel(examInstansi, candidateList) {
    if (typeof XLSX === 'undefined') {
        alert("Library Excel belum selesai dimuat.");
        return;
    }

    const headers = ["No", "NIP", "NAMA", "KEL JABATAN", "UNIT KERJA", "JABATAN", "WAKTU", "PELAKSANAAN", "SESI", "KETERANGAN"];
    
    const rows = candidateList.map((c, index) => [
        index + 1,
        `'${c.nip}`,
        c.nama,
        c.kelJabatan || '-',
        c.unitKerja || '-',
        c.jabatan || '-',
        c.waktu,
        c.pelaksanaan,
        c.sesi,
        c.isFriday && c.sesi === 2 ? 'Khusus Jumat (13.00-16.00)' : (c.sesi === 'NULL' ? 'Belum Terjadwal' : 'Reguler')
    ]);

    const wsData = [headers, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    ws['!cols'] = [
        { wch: 6 },
        { wch: 24 },
        { wch: 35 },
        { wch: 20 },
        { wch: 45 },
        { wch: 35 },
        { wch: 22 },
        { wch: 16 },
        { wch: 8 },
        { wch: 28 }
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Daftar Peserta");

    const safeTitle = (examInstansi || 'Jadwal_Peserta').replace(/[^a-zA-Z0-9_-]/g, '_');
    XLSX.writeFile(wb, `${safeTitle}_Profiling_ASN.xlsx`);
}
