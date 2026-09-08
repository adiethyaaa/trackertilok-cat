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
        // Hindari pergeseran hari karena UTC timezone offset pada Excel
        if (dateInput.getUTCHours() === 0 && dateInput.getUTCMinutes() === 0) {
            return new Date(dateInput.getUTCFullYear(), dateInput.getUTCMonth(), dateInput.getUTCDate());
        }
        return dateInput;
    }

    // Jika berupa angka serial Excel (misal 45000+)
    if (typeof dateInput === 'number' || (!isNaN(dateInput) && !String(dateInput).includes('-') && !String(dateInput).includes('/'))) {
        const serial = Number(dateInput);
        if (serial > 20000 && serial < 60000) {
            // Excel epoch dimulai dari 1899-12-30
            const utcDays = Math.floor(serial - 25569);
            const utcValue = utcDays * 86400;
            const dateInfo = new Date(utcValue * 1000);
            return new Date(dateInfo.getFullYear(), dateInfo.getMonth(), dateInfo.getDate());
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
            return new Date(year, monthMap[monStr], day);
        }
    }

    // Format DD-MM-YYYY atau DD/MM/YYYY
    const ddmmyyyy = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/);
    if (ddmmyyyy) {
        const day = parseInt(ddmmyyyy[1], 10);
        const month = parseInt(ddmmyyyy[2], 10) - 1;
        let year = parseInt(ddmmyyyy[3], 10);
        if (year < 100) year = 2000 + year;
        return new Date(year, month, day);
    }

    // Standard ISO format (YYYY-MM-DD)
    const parsed = new Date(str);
    if (!isNaN(parsed.getTime())) {
        return parsed;
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

