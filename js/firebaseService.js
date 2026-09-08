/**
 * firebaseService.js
 * Modul Integrasi Firebase Realtime Database untuk "Profiling ASN"
 * Menyediakan sinkronisasi online & realtime presensi peserta & jadwal ujian antar panitia.
 */

import { initializeApp, getApps } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import { 
    getDatabase, 
    ref, 
    set, 
    get, 
    onValue, 
    off, 
    update, 
    remove 
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-database.js";

const STORAGE_KEY = 'profiling_asn_firebase_config';

let firebaseApp = null;
let firebaseDb = null;
let isConnected = false;
let activeCandidatesRef = null;
let activeExamsRef = null;
let connectionStatusCallback = null;

/**
 * Mendapatkan konfigurasi Firebase yang aktif saat ini.
 * Memeriksa urutan prioritas:
 * 1. LocalStorage browser (jika pernah diset via modal pengaturan)
 * 2. window.__FIREBASE_CONFIG__ (diinjeksi otomatis oleh GitHub Actions / local dev config)
 */
export function getFirebaseConfig() {
    // 1. Cek dari LocalStorage
    try {
        const stored = localStorage.getItem(STORAGE_KEY);
        if (stored) {
            const parsed = JSON.parse(stored);
            if (parsed && parsed.apiKey && parsed.databaseURL) {
                return parsed;
            }
        }
    } catch (e) {
        console.warn("Gagal membaca config dari localStorage:", e);
    }

    // 2. Cek dari Window Object (injeksi GitHub Actions atau local script)
    if (typeof window !== 'undefined' && window.__FIREBASE_CONFIG__ && window.__FIREBASE_CONFIG__.apiKey) {
        return window.__FIREBASE_CONFIG__;
    }

    return null;
}

/**
 * Menyimpan konfigurasi Firebase ke LocalStorage
 */
export function saveFirebaseConfig(configObj) {
    if (!configObj || !configObj.apiKey || !configObj.databaseURL) {
        throw new Error("Konfigurasi Firebase tidak valid (apiKey dan databaseURL diperlukan)!");
    }
    localStorage.setItem(STORAGE_KEY, JSON.stringify(configObj));
    return initFirebaseService(configObj);
}

/**
 * Menghapus konfigurasi Firebase dari LocalStorage
 */
export function removeFirebaseConfig() {
    localStorage.removeItem(STORAGE_KEY);
    isConnected = false;
    firebaseApp = null;
    firebaseDb = null;
    if (connectionStatusCallback) connectionStatusCallback(false);
}

/**
 * Mendaftarkan callback untuk perubahan status koneksi cloud
 */
export function onConnectionStatusChange(cb) {
    connectionStatusCallback = cb;
    if (connectionStatusCallback) connectionStatusCallback(isConnected);
}

/**
 * Memeriksa apakah cloud aktif dan terkoneksi
 */
export function isCloudActive() {
    return isConnected && firebaseDb !== null;
}

/**
 * Inisialisasi Firebase App & Realtime Database
 */
export function initFirebaseService(customConfig = null) {
    const config = customConfig || getFirebaseConfig();

    if (!config || !config.apiKey || !config.databaseURL) {
        console.info("[Profiling ASN] Konfigurasi Firebase belum terdeteksi.");
        isConnected = false;
        if (connectionStatusCallback) connectionStatusCallback(false);
        return false;
    }

    try {
        const existingApps = getApps();
        if (existingApps.length > 0) {
            firebaseApp = existingApps[0];
        } else {
            firebaseApp = initializeApp(config);
        }

        firebaseDb = getDatabase(firebaseApp);
        isConnected = true;
        console.log("[Profiling ASN] Terhubung ke Firebase Realtime Database:", config.databaseURL);

        // Monitor status koneksi internal Firebase (.info/connected)
        const connectedRef = ref(firebaseDb, ".info/connected");
        onValue(connectedRef, (snap) => {
            const online = snap.val() === true;
            isConnected = online;
            if (connectionStatusCallback) connectionStatusCallback(online);
        });

        if (connectionStatusCallback) connectionStatusCallback(true);
        return true;
    } catch (err) {
        console.error("[Profiling ASN] Gagal inisialisasi Firebase:", err);
        isConnected = false;
        if (connectionStatusCallback) connectionStatusCallback(false);
        return false;
    }
}

/**
 * Helper internal untuk memastikan Firebase Database siap
 */
function ensureDb() {
    if (!firebaseDb) {
        initFirebaseService();
    }
    if (!firebaseDb) {
        throw new Error("Koneksi Firebase Database belum siap. Pastikan konfigurasi Firebase telah diatur.");
    }
    return firebaseDb;
}

// ---------------------- CLOUD EXAM OPERATIONS ----------------------

/**
 * Realtime Listener untuk Daftar Seluruh Ujian
 */
export function listenExamsCloud(callback) {
    try {
        const db = ensureDb();
        if (activeExamsRef) {
            off(activeExamsRef);
        }

        activeExamsRef = ref(db, 'exams');
        onValue(activeExamsRef, (snapshot) => {
            const val = snapshot.val();
            if (!val) {
                callback([]);
                return;
            }

            const examsList = Object.keys(val).map(key => ({
                id: key,
                ...val[key]
            })).sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));

            callback(examsList);
        }, (err) => {
            console.error("Error realtime listening exams:", err);
        });

        return () => {
            if (activeExamsRef) off(activeExamsRef);
        };
    } catch (e) {
        console.warn("listenExamsCloud ditunda:", e.message);
        return () => {};
    }
}

/**
 * Mengambil semua data Ujian dari Firebase Realtime Database
 */
export async function getAllExamsCloud() {
    const db = ensureDb();
    const examsRef = ref(db, 'exams');
    const snapshot = await get(examsRef);
    if (!snapshot.exists()) return [];

    const val = snapshot.val();
    const examsList = Object.keys(val).map(key => ({
        id: key,
        ...val[key]
    })).sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));

    return examsList;
}

/**
 * Mengambil satu Ujian berdasarkan ID
 */
export async function getExamCloud(examId) {
    if (!examId) return null;
    const db = ensureDb();
    const examRef = ref(db, 'exams/' + examId);
    const snapshot = await get(examRef);
    if (!snapshot.exists()) return null;
    return {
        id: examId,
        ...snapshot.val()
    };
}

/**
 * Simpan atau perbarui data Ujian di Firebase Realtime Database
 */
export async function saveExamToCloud(examData) {
    const db = ensureDb();
    const examId = examData.id || ('EXAM-' + Date.now());
    const examRef = ref(db, 'exams/' + examId);

    const payload = {
        title: examData.title || examData.instansi || 'Ujian Tanpa Judul',
        instansi: examData.instansi || '',
        wilker: examData.wilker || 'Papua Barat',
        startDate: examData.startDate || '',
        endDate: examData.endDate || '',
        location: examData.location || '',
        quotaPerSession: Number(examData.quotaPerSession) || 50,
        notes: examData.notes || '',
        updatedAt: new Date().toISOString()
    };

    if (examData.createdAt) {
        payload.createdAt = examData.createdAt;
    } else {
        payload.createdAt = new Date().toISOString();
    }

    await set(examRef, payload);
    return { id: examId, ...payload };
}

/**
 * Hapus data ujian dan seluruh pesertanya dari Firebase Realtime Database
 */
export async function deleteExamFromCloud(examId) {
    if (!examId) return;
    const db = ensureDb();
    const examRef = ref(db, 'exams/' + examId);
    const candidatesRef = ref(db, 'candidates/' + examId);

    await remove(examRef);
    await remove(candidatesRef);
    return true;
}

// ---------------------- CLOUD CANDIDATE OPERATIONS ----------------------

/**
 * Realtime Listener untuk Kandidat dalam Ujian Tertentu
 */
export function listenCandidatesCloud(examId, callback) {
    if (!examId) return () => {};
    try {
        const db = ensureDb();
        if (activeCandidatesRef) {
            off(activeCandidatesRef);
        }

        activeCandidatesRef = ref(db, 'candidates/' + examId);
        onValue(activeCandidatesRef, (snapshot) => {
            const val = snapshot.val();
            if (!val) {
                callback([]);
                return;
            }

            const candidatesList = Object.keys(val).map(key => ({
                id: isNaN(Number(key)) ? key : Number(key),
                ...val[key]
            })).sort((a, b) => {
                const sesiA = Number(a.sesi) || 0;
                const sesiB = Number(b.sesi) || 0;
                if (sesiA !== sesiB) return sesiA - sesiB;
                return (Number(a.no) || 0) - (Number(b.no) || 0);
            });

            callback(candidatesList);
        }, (err) => {
            console.error("Error realtime listening candidates:", err);
        });

        return () => {
            if (activeCandidatesRef) off(activeCandidatesRef);
        };
    } catch (e) {
        console.warn("listenCandidatesCloud ditunda:", e.message);
        return () => {};
    }
}

/**
 * Mengambil daftar peserta Ujian dari Firebase Realtime Database
 */
export async function getCandidatesByExamCloud(examId) {
    if (!examId) return [];
    const db = ensureDb();
    const candidatesRef = ref(db, 'candidates/' + examId);
    const snapshot = await get(candidatesRef);
    if (!snapshot.exists()) return [];

    const val = snapshot.val();
    const candidatesList = Object.keys(val).map(key => ({
        id: isNaN(Number(key)) ? key : key,
        ...val[key]
    })).sort((a, b) => {
        const sesiA = Number(a.sesi) || 0;
        const sesiB = Number(b.sesi) || 0;
        if (sesiA !== sesiB) return sesiA - sesiB;
        return (Number(a.no) || 0) - (Number(b.no) || 0);
    });

    return candidatesList;
}

/**
 * Simpan peserta secara bulk ke Firebase Realtime Database
 */
export async function bulkAddCandidatesToCloud(examId, candidateList) {
    if (!examId || !candidateList || candidateList.length === 0) return 0;
    const db = ensureDb();

    const updates = {};
    candidateList.forEach((c, idx) => {
        const key = String(c.nip || c.id || ('cand_' + idx)).trim().replace(/[.#$[\]/]/g, '_');
        updates['candidates/' + examId + '/' + key] = {
            examId: examId,
            no: c.no || (idx + 1),
            nip: String(c.nip || '').trim(),
            nama: String(c.nama || '').trim(),
            unitKerja: String(c.unitKerja || '').trim(),
            jabatan: String(c.jabatan || '').trim(),
            waktu: String(c.waktu || '').trim(),
            pelaksanaan: String(c.pelaksanaan || '').trim(),
            sesi: Number(c.sesi) || 1,
            isFriday: Boolean(c.isFriday),
            status: c.status || 'Terjadwal',
            kehadiran: c.kehadiran || 'BELUM',
            updatedAt: new Date().toISOString()
        };
    });

    await update(ref(db), updates);
    return candidateList.length;
}

/**
 * Tambah atau update single kandidat di Firebase Realtime Database
 */
export async function addOrUpdateCandidateCloud(examId, candidate) {
    const db = ensureDb();
    const cExamId = examId || candidate.examId;
    if (!cExamId) throw new Error("examId diperlukan untuk menyimpan kandidat");

    const key = String(candidate.nip || candidate.id || ('cand_' + Date.now())).trim().replace(/[.#$[\]/]/g, '_');
    const targetRef = ref(db, 'candidates/' + cExamId + '/' + key);

    const data = {
        examId: cExamId,
        no: candidate.no || 1,
        nip: String(candidate.nip || '').trim(),
        nama: String(candidate.nama || '').trim(),
        unitKerja: String(candidate.unitKerja || '').trim(),
        jabatan: String(candidate.jabatan || '').trim(),
        waktu: String(candidate.waktu || '').trim(),
        pelaksanaan: String(candidate.pelaksanaan || '').trim(),
        sesi: Number(candidate.sesi) || 1,
        isFriday: Boolean(candidate.isFriday),
        status: candidate.status || 'Terjadwal',
        kehadiran: candidate.kehadiran || 'BELUM',
        updatedAt: new Date().toISOString()
    };

    await set(targetRef, data);
    return { id: key, ...data };
}

/**
 * Hapus single kandidat di Firebase Realtime Database
 */
export async function deleteCandidateCloud(examId, candidateIdOrNip) {
    if (!examId || !candidateIdOrNip) return false;
    const db = ensureDb();
    const key = String(candidateIdOrNip).trim().replace(/[.#$[\]/]/g, '_');
    const targetRef = ref(db, 'candidates/' + examId + '/' + key);
    await remove(targetRef);
    return true;
}

/**
 * Hapus semua kandidat dalam satu ujian
 */
export async function deleteCandidatesByExamCloud(examId) {
    if (!examId) return false;
    const db = ensureDb();
    await remove(ref(db, 'candidates/' + examId));
    return true;
}

/**
 * Hapus kandidat berdasarkan daftar NIP
 */
export async function deleteCandidatesByNipsCloud(examId, nipList) {
    if (!examId || !nipList || nipList.length === 0) return 0;
    const db = ensureDb();
    const updates = {};
    nipList.forEach(nip => {
        const key = String(nip).trim().replace(/[.#$[\]/]/g, '_');
        updates['candidates/' + examId + '/' + key] = null;
    });
    await update(ref(db), updates);
    return nipList.length;
}

/**
 * Update Status Kehadiran secara Realtime ke Firebase
 */
export async function updateAttendanceInCloud(examId, candidateIdOrNip, status) {
    if (!examId || !candidateIdOrNip) return false;
    const db = ensureDb();

    const safeKey = String(candidateIdOrNip).trim().replace(/[.#$[\]/]/g, '_');
    const targetRef = ref(db, 'candidates/' + examId + '/' + safeKey);

    try {
        await update(targetRef, {
            kehadiran: status,
            attendanceTimestamp: new Date().toISOString()
        });
        return true;
    } catch (e) {
        console.warn("Gagal update via direct key, mencoba pencarian NIP:", e);
        const parentRef = ref(db, 'candidates/' + examId);
        const snap = await get(parentRef);
        if (snap.exists()) {
            const data = snap.val();
            for (const key of Object.keys(data)) {
                if (data[key].nip === String(candidateIdOrNip) || key === String(candidateIdOrNip)) {
                    await update(ref(db, 'candidates/' + examId + '/' + key), {
                        kehadiran: status,
                        attendanceTimestamp: new Date().toISOString()
                    });
                    return true;
                }
            }
        }
        return false;
    }
}

/**
 * Menghitung statistik ujian langsung dari data Cloud
 */
export async function getExamStatsCloud(examId) {
    const list = await getCandidatesByExamCloud(examId);
    const stats = {
        total: list.length,
        sesi1: 0,
        sesi2: 0,
        sesi3: 0,
        byDate: {},
        unitKerjaCount: new Set()
    };

    list.forEach(item => {
        if (item.sesi === 1) stats.sesi1++;
        else if (item.sesi === 2) stats.sesi2++;
        else if (item.sesi === 3) stats.sesi3++;

        const d = item.pelaksanaan || 'Tidak Ditentukan';
        stats.byDate[d] = (stats.byDate[d] || 0) + 1;

        if (item.unitKerja) {
            stats.unitKerjaCount.add(item.unitKerja);
        }
    });

    stats.totalUnitKerja = stats.unitKerjaCount.size;
    return stats;
}
