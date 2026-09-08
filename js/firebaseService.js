/**
 * firebaseService.js
 * Modul Integrasi Firebase Realtime Database untuk "Profiling ASN"
 * Mendukung sinkronisasi realtime presensi peserta & jadwal ujian antar panitia.
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
    remove, 
    child 
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-database.js";

const STORAGE_KEY = 'profiling_asn_firebase_config';

let firebaseApp = null;
let firebaseDb = null;
let isConnected = false;
let activeCandidatesRef = null;
let activeExamsRef = null;
let connectionStatusCallback = null;

/**
 * Mendapatkan konfigurasi Firebase yang aktif saat ini
 */
export function getFirebaseConfig() {
    // 1. Cek dari Window Object (injeksi GitHub Actions / file local script)
    if (window.__FIREBASE_CONFIG__ && window.__FIREBASE_CONFIG__.apiKey) {
        return window.__FIREBASE_CONFIG__;
    }

    // 2. Cek dari LocalStorage (disimpan via UI Pengaturan)
    try {
        const stored = localStorage.getItem(STORAGE_KEY);
        if (stored) {
            const parsed = JSON.parse(stored);
            if (parsed && parsed.apiKey) {
                return parsed;
            }
        }
    } catch (e) {
        console.warn("Gagal membaca config dari localStorage:", e);
    }

    return null;
}

/**
 * Menyimpan konfigurasi Firebase ke LocalStorage
 */
export function saveFirebaseConfig(configObj) {
    if (!configObj || !configObj.apiKey) {
        throw new Error("Konfigurasi Firebase tidak valid (apiKey diperlukan)!");
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
        console.info("[Profiling ASN] Firebase belum dikonfigurasi. Berjalan dalam mode Offline (IndexedDB lokal).");
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

        return true;
    } catch (err) {
        console.error("[Profiling ASN] Gagal inisialisasi Firebase:", err);
        isConnected = false;
        if (connectionStatusCallback) connectionStatusCallback(false);
        return false;
    }
}

// ---------------------- CLOUD EXAM OPERATIONS ----------------------

/**
 * Realtime Listener untuk Daftar Seluruh Ujian
 */
export function listenExamsCloud(callback) {
    if (!isCloudActive()) return () => {};

    if (activeExamsRef) {
        off(activeExamsRef);
    }

    activeExamsRef = ref(firebaseDb, 'exams');
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
}

/**
 * Simpan atau perbarui data Ujian di Firebase Realtime Database
 */
export async function saveExamToCloud(examData) {
    if (!isCloudActive()) return null;
    const examId = examData.id;
    const examRef = ref(firebaseDb, exams/);

    const payload = {
        title: examData.title || examData.instansi || '',
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
    if (!isCloudActive()) return;
    const examRef = ref(firebaseDb, exams/);
    const candidatesRef = ref(firebaseDb, candidates/);

    await remove(examRef);
    await remove(candidatesRef);
}

// ---------------------- CLOUD CANDIDATE OPERATIONS ----------------------

/**
 * Realtime Listener untuk Kandidat dalam Ujian Tertentu
 */
export function listenCandidatesCloud(examId, callback) {
    if (!isCloudActive() || !examId) return () => {};

    if (activeCandidatesRef) {
        off(activeCandidatesRef);
    }

    activeCandidatesRef = ref(firebaseDb, candidates/);
    onValue(activeCandidatesRef, (snapshot) => {
        const val = snapshot.val();
        if (!val) {
            callback([]);
            return;
        }

        const candidatesList = Object.keys(val).map(key => ({
            id: isNaN(Number(key)) ? key : Number(key),
            ...val[key]
        }));

        callback(candidatesList);
    }, (err) => {
        console.error("Error realtime listening candidates:", err);
    });

    return () => {
        if (activeCandidatesRef) off(activeCandidatesRef);
    };
}

/**
 * Simpan peserta secara bulk ke Firebase Realtime Database
 */
export async function bulkAddCandidatesToCloud(examId, candidateList) {
    if (!isCloudActive() || !examId) return 0;

    const updates = {};
    candidateList.forEach((c, idx) => {
        // Gunakan NIP atau ID unik yang aman
        const key = String(c.id || c.nip || ('cand_' + idx)).replace(/[.#$[\]/]/g, '_');
        updates[candidates//] = {
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

    await update(ref(firebaseDb), updates);
    return candidateList.length;
}

/**
 * Update Status Kehadiran secara Realtime ke Firebase
 */
export async function updateAttendanceInCloud(examId, candidateIdOrNip, status) {
    if (!isCloudActive() || !examId) return false;

    const safeKey = String(candidateIdOrNip).replace(/[.#$[\]/]/g, '_');
    const targetRef = ref(firebaseDb, candidates//);

    try {
        await update(targetRef, {
            kehadiran: status,
            attendanceTimestamp: new Date().toISOString()
        });
        return true;
    } catch (e) {
        console.warn("Gagal update via direct key, mencoba pencarian NIP:", e);
        // Fallback jika candidateId menggunakan auto-increment ID IndexedDB
        const parentRef = ref(firebaseDb, candidates/);
        const snap = await get(parentRef);
        if (snap.exists()) {
            const data = snap.val();
            for (const key of Object.keys(data)) {
                if (data[key].nip === String(candidateIdOrNip) || key === String(candidateIdOrNip)) {
                    await update(ref(firebaseDb, candidates//), {
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
 * Migrasi seluruh data dari IndexedDB lokal ke Firebase Realtime Database
 */
export async function migrateIndexedDBToFirebase(getAllExamsFn, getCandidatesByExamFn) {
    if (!isCloudActive()) {
        throw new Error("Koneksi Firebase Cloud belum aktif!");
    }

    const exams = await getAllExamsFn();
    if (exams.length === 0) {
        return { examsCount: 0, candidatesCount: 0 };
    }

    let totalCandidates = 0;

    for (const exam of exams) {
        await saveExamToCloud(exam);
        const candidates = await getCandidatesByExamFn(exam.id);
        if (candidates.length > 0) {
            await bulkAddCandidatesToCloud(exam.id, candidates);
            totalCandidates += candidates.length;
        }
    }

    return {
        examsCount: exams.length,
        candidatesCount: totalCandidates
    };
}
