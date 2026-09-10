/**
 * sessionRules.js
 * Aturan waktu dan sesi untuk aplikasi "Profiling ASN"
 * 
 * Aturan standar:
 * Sesi 1: 08.00 s.d. 11.00 WIT
 * Sesi 2: 11.00 s.d. 14.00 WIT (Hari Biasa: Senin-Kamis, Sabtu-Minggu)
 *         13.00 s.d. 16.00 WIT (Khusus Hari Jumat)
 * Sesi 3: 14.00 s.d. 17.00 WIT
 */

export const SESSION_CONFIG = {
    1: {
        name: "Sesi 1",
        timeRegular: "08.00 - 11.00 WIT",
        timeFriday: "08.00 - 11.00 WIT"
    },
    2: {
        name: "Sesi 2",
        timeRegular: "11.00 - 14.00 WIT",
        timeFriday: "13.00 - 16.00 WIT" // Khusus Jumat
    },
    3: {
        name: "Sesi 3",
        timeRegular: "14.00 - 17.00 WIT",
        timeFriday: "14.00 - 17.00 WIT"
    }
};

/**
 * Parsing teks tanggal fleksibel ke Date object
 * Mendukung format: 
 * - "11-Sep-26" atau "11-Sep-2026"
 * - "Jumat, 11-Sep-26" atau "Jumat, 11 September 2026"
 * - "2026-09-11"
 * - "11/09/2026" atau "11-09-2026"
 * - Nomor seri tanggal Excel (misal: 46276)
 */
export function parseFlexibleDate(dateInput) {
    if (!dateInput) return null;

    if (dateInput instanceof Date && !isNaN(dateInput)) {
        // Ambil komponen tanggal dan hindari pergeseran hari karena UTC/local timezone
        let y = dateInput.getFullYear();
        let m = dateInput.getMonth();
        let d = dateInput.getDate();

        // Jika jam UTC adalah 00:00 (seperti yang dihasilkan oleh SheetJS untuk sel tanggal),
        // di timezone manapun kita utamakan tanggal UTC-nya
        if (dateInput.getUTCHours() === 0 && dateInput.getUTCMinutes() === 0) {
            y = dateInput.getUTCFullYear();
            m = dateInput.getUTCMonth();
            d = dateInput.getUTCDate();
        } else if (dateInput.getUTCHours() >= 20) {
            // Misal 23:59:59 dari floating rounding error SheetJS, geser ke hari berikutnya
            const next = new Date(dateInput.getTime() + (4 * 3600 * 1000));
            y = next.getUTCFullYear();
            m = next.getUTCMonth();
            d = next.getUTCDate();
        }
        return new Date(y, m, d, 12, 0, 0);
    }

    // Jika berupa angka serial Excel (misal 45000+)
    if (typeof dateInput === 'number' || (!isNaN(dateInput) && !String(dateInput).includes('-') && !String(dateInput).includes('/'))) {
        const serial = Number(dateInput);
        if (serial > 20000 && serial < 60000) {
            // Excel epoch dimulai dari 1899-12-30
            // Tambahkan 0.0001 untuk mengatasi floating point 23:59:59.999
            const totalDays = Math.floor(serial + 0.0001);
            const utcMs = (totalDays - 25569) * 86400000;
            const dateInfo = new Date(utcMs);
            return new Date(dateInfo.getUTCFullYear(), dateInfo.getUTCMonth(), dateInfo.getUTCDate(), 12, 0, 0);
        }
    }

    let str = String(dateInput).trim();

    // Hapus awalan nama hari jika ada (misal: "Jumat, 11-Sep-26" atau "Jumat 11/09/2026")
    str = str.replace(/^(senin|selasa|rabu|kamis|jumat|sabtu|minggu|jum'at)[,\s]+/i, '').trim();

    const monthMap = {
        'jan': 0, 'januari': 0, 'january': 0,
        'feb': 1, 'februari': 1, 'february': 1,
        'mar': 2, 'maret': 2, 'march': 2,
        'apr': 3, 'april': 3,
        'mei': 4, 'may': 4,
        'jun': 5, 'juni': 5, 'june': 5,
        'jul': 6, 'juli': 6, 'july': 6,
        'agu': 7, 'agus': 7, 'agust': 7, 'agustus': 7, 'aug': 7, 'august': 7,
        'sep': 8, 'sept': 8, 'september': 8,
        'okt': 9, 'oktober': 9, 'oct': 9, 'october': 9,
        'nov': 10, 'november': 10,
        'des': 11, 'desember': 11, 'dec': 11, 'december': 11
    };

    // Format "11-Sep-26" atau "11-Sep-2026" atau "11 September 2026"
    const textMatch = str.match(/^(\d{1,2})[-/\s]+([A-Za-z]+)[-/\s]+(\d{2,4})$/);
    if (textMatch) {
        const day = parseInt(textMatch[1], 10);
        const monStr = textMatch[2].toLowerCase();
        let year = parseInt(textMatch[3], 10);
        if (year < 100) {
            year = 2000 + year; // 26 -> 2026
        }
        if (monthMap[monStr] !== undefined) {
            return new Date(year, monthMap[monStr], day, 12, 0, 0);
        }
    }

    // Format DD-MM-YYYY atau DD/MM/YYYY
    const ddmmyyyy = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/);
    if (ddmmyyyy) {
        const day = parseInt(ddmmyyyy[1], 10);
        const month = parseInt(ddmmyyyy[2], 10) - 1;
        let year = parseInt(ddmmyyyy[3], 10);
        if (year < 100) year = 2000 + year;
        return new Date(year, month, day, 12, 0, 0);
    }

    // Format ISO YYYY-MM-DD
    const isoyyyymmdd = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (isoyyyymmdd) {
        const year = parseInt(isoyyyymmdd[1], 10);
        const month = parseInt(isoyyyymmdd[2], 10) - 1;
        const day = parseInt(isoyyyymmdd[3], 10);
        return new Date(year, month, day, 12, 0, 0);
    }

    return null;
}

/**
 * Mengecek apakah tanggal pelaksanaan jatuh pada hari Jumat
 * @param {string|Date|number} dateInput 
 * @returns {boolean}
 */
export function isFriday(dateInput) {
    const d = parseFlexibleDate(dateInput);
    if (!d) return false;
    return d.getDay() === 5; // 5 = Friday
}

/**
 * Mengambil nama hari dalam Bahasa Indonesia
 * @param {string|Date|number} dateInput 
 * @returns {string}
 */
export function getDayNameID(dateInput) {
    const d = parseFlexibleDate(dateInput);
    if (!d) return '';
    const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
    return days[d.getDay()];
}

/**
 * Format tanggal ke tampilan standar Indonesia: "11-Sep-26" atau "11 September 2026"
 */
export function formatDateDisplay(dateInput, format = 'short') {
    const d = parseFlexibleDate(dateInput);
    if (!d) return String(dateInput || '');

    const monthsShort = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agust', 'Sept', 'Okt', 'Nov', 'Des'];
    const monthsLong = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

    const day = String(d.getDate()).padStart(2, '0');
    const yearFull = d.getFullYear();

    if (format === 'short') {
        return `${day} ${monthsShort[d.getMonth()]} ${yearFull}`;
    }
    if (format === 'long') {
        return `${day} ${monthsLong[d.getMonth()]} ${yearFull}`;
    }
    if (format === 'input') {
        // YYYY-MM-DD untuk input date HTML
        const m = String(d.getMonth() + 1).padStart(2, '0');
        return `${yearFull}-${m}-${day}`;
    }
    return `${day} ${monthsShort[d.getMonth()]} ${yearFull}`;
}

/**
 * Mendapatkan jam pelaksanaan ujian berdasarkan nomor sesi dan tanggal pelaksanaan
 * @param {number|string} sessionNum Nomor sesi (1, 2, atau 3)
 * @param {string|Date|number} dateInput Tanggal pelaksanaan
 * @returns {string} Contoh: "08.00 - 11.00 WIT" atau "13.00 - 16.00 WIT"
 */
export function getSessionTime(sessionNum, dateInput) {
    const s = parseInt(sessionNum, 10);
    const friday = isFriday(dateInput);

    if (SESSION_CONFIG[s]) {
        return friday ? SESSION_CONFIG[s].timeFriday : SESSION_CONFIG[s].timeRegular;
    }
    return "-";
}

/**
 * Format nomor sesi kumulatif dengan leading zero untuk angka 1-9 (01 - 09, 10, ...)
 * @param {number|string} num 
 * @returns {string} Contoh: 1 -> "01", 7 -> "07", 10 -> "10"
 */
export function formatCumulativeSessionNumber(num) {
    const n = Number(num);
    if (isNaN(n) || n <= 0) return '00';
    return String(n).padStart(2, '0');
}

/**
 * Mengambil batas maksimum sesi untuk tanggal tertentu.
 * Hari Jumat: 2 sesi (Sesi 1 & Sesi 2)
 * Hari Lainnya: 3 sesi (Sesi 1, Sesi 2, & Sesi 3)
 * @param {string|Date|number} dateInput 
 * @returns {number}
 */
export function getMaxSessionsForDate(dateInput) {
    return isFriday(dateInput) ? 2 : 3;
}

/**
 * Menghitung nomor sesi kumulatif/akumulasi kandidat berdasarkan urutan tanggal pelaksanaan
 * Hari Jumat otomatis dihitung maksimal 2 sesi, sedangkan hari biasa dihitung 3 sesi.
 * @param {Object} candidate Objek kandidat dengan properti pelaksanaan dan sesi
 * @param {Array<string>} sortedDates Array tanggal pelaksanaan terurut kronologis
 * @returns {number|null} Nomor sesi kumulatif (misal 1..n) atau null jika belum terjadwal
 */
export function calculateCumulativeSessionNumber(candidate, sortedDates) {
    if (!candidate || !candidate.pelaksanaan || candidate.pelaksanaan === 'NULL' || candidate.pelaksanaan === '-') return null;
    const rawSesi = candidate.sesi;
    if (rawSesi === undefined || rawSesi === null || rawSesi === 'NULL' || rawSesi === '00' || rawSesi === 0 || rawSesi === '0') {
        return null;
    }
    const s = Number(rawSesi);
    if (isNaN(s) || s < 1) return null;

    if (!Array.isArray(sortedDates) || sortedDates.length === 0) {
        return s;
    }

    const dayIndex = sortedDates.indexOf(candidate.pelaksanaan);
    if (dayIndex === -1) return s;

    let totalPreceding = 0;
    for (let i = 0; i < dayIndex; i++) {
        totalPreceding += getMaxSessionsForDate(sortedDates[i]);
    }
    return totalPreceding + s;
}

/**
 * Mengonversi nomor sesi kumulatif/akumulasi menjadi tanggal pelaksanaan dan sesi harian (1, 2, atau 3)
 * Otomatis mendeteksi Hari Jumat yang hanya memiliki 2 sesi.
 * @param {number|string} cumulativeSessionNum Nomor sesi kumulatif (misal 1, 2, 3, 4, ...)
 * @param {Array<string>} sortedDates Array tanggal pelaksanaan terurut kronologis
 * @returns {{ targetPelaksanaan: string|null, targetDailySesi: number, scheduleDetail: string }}
 */
export function convertCumulativeSessionToDaily(cumulativeSessionNum, sortedDates) {
    const cum = Number(cumulativeSessionNum);
    if (isNaN(cum) || cum < 1) {
        return { targetPelaksanaan: null, targetDailySesi: 1, scheduleDetail: 'Sesi Tidak Valid' };
    }

    if (!Array.isArray(sortedDates) || sortedDates.length === 0) {
        return {
            targetPelaksanaan: null,
            targetDailySesi: cum,
            scheduleDetail: `Sesi Akumulasi ${cum}`
        };
    }

    let remaining = cum;
    for (let i = 0; i < sortedDates.length; i++) {
        const dateStr = sortedDates[i];
        const dayMax = getMaxSessionsForDate(dateStr);

        if (remaining <= dayMax) {
            const isFri = isFriday(dateStr);
            return {
                targetPelaksanaan: dateStr,
                targetDailySesi: remaining,
                scheduleDetail: `Sesi Akumulasi ${cum} (${dateStr}${isFri ? ' [Jumat]' : ''}, Sesi ${remaining})`
            };
        }
        remaining -= dayMax;
    }

    // Jika nomor sesi melebihi total seluruh slot yang tersedia pada tanggal yang ada:
    const lastDate = sortedDates[sortedDates.length - 1];
    const isFri = isFriday(lastDate);
    const lastDayMax = getMaxSessionsForDate(lastDate);
    const dailySesi = Math.min(remaining, lastDayMax);

    return {
        targetPelaksanaan: lastDate,
        targetDailySesi: dailySesi,
        scheduleDetail: `Sesi Akumulasi ${cum} (${lastDate}${isFri ? ' [Jumat]' : ''}, Sesi ${dailySesi})`
    };
}


