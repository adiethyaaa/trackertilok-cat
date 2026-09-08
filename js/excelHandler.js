/**
 * excelHandler.js
 * Modul untuk membaca file Excel, parsing struktur kolom referensi,
 * kalkulasi otomatis sesi/waktu, pembuatan template Excel, export data,
 * dan deteksi duplikasi NIP (internal file Excel maupun terhadap database eksisting).
 */

import { parseFlexibleDate, isFriday, getSessionTime, formatDateDisplay } from './sessionRules.js';

// Header acuan standar sesuai instruksi user (termasuk variasi ejaan "PELAKSAAN" dan "PELAKSANAAN")
export const REQUIRED_COLUMNS = [
    { key: 'no', label: 'No', aliases: ['no', 'nomor', 'no.', 'urut', 'num'] },
    { key: 'nip', label: 'NIP', aliases: ['nip', 'nip baru', 'nomor induk pegawai', 'nrp', 'nip_peserta', 'nip peserta'] },
    { key: 'nama', label: 'NAMA', aliases: ['nama', 'nama lengkap', 'nama peserta', 'pegawai', 'nama_pegawai', 'nama pegawai'] },
    { key: 'unitKerja', label: 'UNIT KERJA', aliases: ['unit kerja', 'unit_kerja', 'instansi / unit kerja', 'skpd', 'opd', 'bagian', 'satuan kerja', 'satker', 'unit'] },
    { key: 'jabatan', label: 'JABATAN', aliases: ['jabatan', 'nama jabatan', 'posisi', 'jabatan sekarang'] },
    { key: 'waktu', label: 'WAKTU', aliases: ['waktu', 'jam', 'waktu ujian', 'pukul', 'jadwal', 'jam ujian', 'waktu pelaksanaan', 'jam pelaksanaan'] },
    { key: 'pelaksanaan', label: 'PELAKSANAAN', aliases: [
        'pelaksanaan', 'pelaksaan', 'pelaksana', 'tanggal', 'tgl', 
        'tgl pelaksanaan', 'tgl pelaksaan', 'tgl. pelaksanaan', 'tgl. pelaksaan', 
        'tanggal pelaksanaan', 'tanggal pelaksaan', 'jadwal pelaksanaan', 'jadwal pelaksaan',
        'hari/tanggal', 'hari / tanggal', 'hari, tanggal', 'hari tanggal', 'tanggal ujian', 'tgl ujian',
        'tgl_pelaksanaan', 'tgl_pelaksaan'
    ] },
    { key: 'sesi', label: 'SESI', aliases: ['sesi', 'sesi ujian', 'sesi ke', 'tahap', 'sesi_ujian', 'sesi pelaksanaan'] }
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

    // 2. Cek pencocokan parsial / keyword
    if (clean.includes('pelaksaan') || clean.includes('pelaksanaan') || clean.includes('tanggal') || clean.includes('tgl')) {
        return 'pelaksanaan';
    }
    if (clean === 'nip' || clean.startsWith('nip ') || clean.includes('nomor induk')) {
        return 'nip';
    }
    if (clean.includes('nama')) {
        return 'nama';
    }
    if (clean.includes('sesi')) {
        return 'sesi';
    }
    if (clean.includes('waktu') || clean.includes('pukul') || (clean.includes('jam') && !clean.includes('jambatan'))) {
        return 'waktu';
    }
    if (clean.includes('unit') || clean.includes('kerja') || clean.includes('opd') || clean.includes('skpd') || clean.includes('satker')) {
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
 * Mendukung ribuan baris data dengan scanning baris demi baris, melewati baris kosong,
 * dan hanya memasukkan baris yang memenuhi kondisi NAMA dan NIP terisi.
 * 
 * @param {File} file Objek File dari input
 * @param {Object} options Opsi { autoStandardizeTime: boolean, defaultPelaksanaan?: string }
 * @returns {Promise<Object>} { success, candidates, summary }
 */
export async function parseExcelFile(file, options = { autoStandardizeTime: true, defaultPelaksanaan: '' }) {
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
        throw new Error("Format kolom Excel tidak dikenali. Pastikan terdapat kolom NIP dan NAMA (serta PELAKSANAAN / PELAKSAAN, UNIT KERJA, JABATAN, WAKTU, SESI).");
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
            unitKerja: '',
            jabatan: '',
            waktu: '',
            pelaksanaan: '',
            sesi: 1,
            isFriday: false,
            excelRowNumber: i + 1 // catat nomor baris asli di file Excel
        };

        Object.keys(headerMap).forEach(colIdx => {
            const fieldKey = headerMap[colIdx];
            let val = row[colIdx];

            if (fieldKey === 'pelaksanaan') {
                // Cek apakah cell di worksheet memiliki formatted text (.w) asli dari Excel (misal: "09-Sep-26" atau "9/9/2026")
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
            } else {
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

        // Normalisasi Sesi (default 1 jika kosong atau tidak valid)
        let sesiNum = parseInt(rowObj.sesi, 10);
        if (isNaN(sesiNum) || sesiNum < 1 || sesiNum > 3) {
            sesiNum = 1;
        }
        rowObj.sesi = sesiNum;

        // Jika kolom pelaksanaan masih kosong di Excel, gunakan default dari ujian jika tersedia
        if (!rowObj.pelaksanaan && options.defaultPelaksanaan) {
            rowObj.pelaksanaan = options.defaultPelaksanaan;
        }

        // Cek apakah tanggal pelaksanaan jatuh pada hari Jumat
        if (rowObj.pelaksanaan) {
            const parsedDate = parseFlexibleDate(rowObj.pelaksanaan);
            if (parsedDate) {
                rowObj.pelaksanaan = formatDateDisplay(parsedDate, 'short');
                rowObj.isFriday = isFriday(parsedDate);
            } else {
                rowObj.isFriday = false;
            }
        }

        // Penentuan Jam Ujian Otomatis berdasarkan Sesi & Aturan Khusus Hari Jumat
        const standardTime = getSessionTime(rowObj.sesi, rowObj.pelaksanaan);
        if (options.autoStandardizeTime || !rowObj.waktu || rowObj.waktu === '-' || rowObj.waktu.trim() === '') {
            rowObj.waktu = standardTime;
        }

        parsedCandidates.push(rowObj);
    }

    return {
        success: parsedCandidates.length > 0,
        candidates: parsedCandidates,
        summary: {
            totalRows: parsedCandidates.length,
            skippedRows: skippedCount,
            sesi1: parsedCandidates.filter(c => c.sesi === 1).length,
            sesi2: parsedCandidates.filter(c => c.sesi === 2).length,
            sesi3: parsedCandidates.filter(c => c.sesi === 3).length,
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
 * Generate dan unduh template file Excel kosong/contoh dengan format yang presisi
 */
export function downloadExcelTemplate() {
    if (typeof XLSX === 'undefined') {
        alert("Library Excel belum selesai dimuat.");
        return;
    }

    const headers = ["No", "NIP", "NAMA", "UNIT KERJA", "JABATAN", "WAKTU", "PELAKSANAAN", "SESI"];
    
    // Contoh data riil sesuai permintaan user (11-Sep-26 jatuh pada hari Jumat)
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
 * Export data peserta ujian ke file Excel
 */
export function exportCandidatesToExcel(examInstansi, candidateList) {
    if (typeof XLSX === 'undefined') {
        alert("Library Excel belum selesai dimuat.");
        return;
    }

    const headers = ["No", "NIP", "NAMA", "UNIT KERJA", "JABATAN", "WAKTU", "PELAKSANAAN", "SESI", "KETERANGAN"];
    
    const rows = candidateList.map((c, index) => [
        index + 1,
        `'${c.nip}`,
        c.nama,
        c.unitKerja || '-',
        c.jabatan || '-',
        c.waktu,
        c.pelaksanaan,
        c.sesi,
        c.isFriday && c.sesi === 2 ? 'Khusus Jumat (13.00-16.00)' : 'Reguler'
    ]);

    const wsData = [headers, ...rows];
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    ws['!cols'] = [
        { wch: 6 },
        { wch: 24 },
        { wch: 35 },
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
