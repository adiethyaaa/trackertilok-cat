/**
 * db.js
 * Modul Database IndexedDB untuk "Profiling ASN"
 * Menyimpan data Ujian (exams) dan Peserta (candidates) secara lokal, aman, dan berkapasitas besar.
 */

const DB_NAME = 'ProfilingASN_DB';
const DB_VERSION = 1;

let dbInstance = null;

export function openDatabase() {
    return new Promise((resolve, reject) => {
        if (dbInstance) {
            resolve(dbInstance);
            return;
        }

        const request = indexedDB.open(DB_NAME, DB_VERSION);

        request.onupgradeneeded = (event) => {
            const db = event.target.result;

            // Store untuk Ujian
            if (!db.objectStoreNames.contains('exams')) {
                const examStore = db.createObjectStore('exams', { keyPath: 'id' });
                examStore.createIndex('instansi', 'instansi', { unique: false });
                examStore.createIndex('wilker', 'wilker', { unique: false });
                examStore.createIndex('createdAt', 'createdAt', { unique: false });
            }

            // Store untuk Peserta Ujian
            if (!db.objectStoreNames.contains('candidates')) {
                const candidateStore = db.createObjectStore('candidates', { keyPath: 'id', autoIncrement: true });
                candidateStore.createIndex('examId', 'examId', { unique: false });
                candidateStore.createIndex('nip', 'nip', { unique: false });
                candidateStore.createIndex('sesi', 'sesi', { unique: false });
                candidateStore.createIndex('pelaksanaan', 'pelaksanaan', { unique: false });
                candidateStore.createIndex('exam_sesi', ['examId', 'sesi'], { unique: false });
            }
        };

        request.onsuccess = (event) => {
            dbInstance = event.target.result;
            resolve(dbInstance);
        };

        request.onerror = (event) => {
            console.error('IndexedDB Error:', event.target.error);
            reject(event.target.error);
        };
    });
}

// ---------------------- EXAM OPERATIONS ----------------------

export async function addExam(examData) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('exams', 'readwrite');
        const store = tx.objectStore('exams');
        const item = {
            id: examData.id || 'EXAM-' + Date.now(),
            title: examData.title || 'Ujian Tanpa Judul',
            wilker: examData.wilker || 'Papua Barat',
            instansi: examData.instansi || '',
            startDate: examData.startDate || '',
            endDate: examData.endDate || '',
            location: examData.location || 'Lokasi Belum Ditentukan',
            quotaPerSession: Number(examData.quotaPerSession) || 50,
            notes: examData.notes || '',
            createdAt: new Date().toISOString()
        };
        const req = store.add(item);
        req.onsuccess = () => resolve(item);
        req.onerror = (e) => reject(e.target.error);
    });
}

export async function getAllExams() {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('exams', 'readonly');
        const store = tx.objectStore('exams');
        const req = store.getAll();
        req.onsuccess = () => {
            // Urutkan dari yang terbaru dibuat
            const list = (req.result || []).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
            resolve(list);
        };
        req.onerror = (e) => reject(e.target.error);
    });
}

export async function getExamById(id) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('exams', 'readonly');
        const store = tx.objectStore('exams');
        const req = store.get(id);
        req.onsuccess = () => resolve(req.result || null);
        req.onerror = (e) => reject(e.target.error);
    });
}

export async function updateExam(examData) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('exams', 'readwrite');
        const store = tx.objectStore('exams');
        const req = store.put(examData);
        req.onsuccess = () => resolve(examData);
        req.onerror = (e) => reject(e.target.error);
    });
}

export async function deleteExam(examId) {
    const db = await openDatabase();
    // Hapus ujian sekaligus seluruh pesertanya
    await deleteCandidatesByExam(examId);
    return new Promise((resolve, reject) => {
        const tx = db.transaction('exams', 'readwrite');
        const store = tx.objectStore('exams');
        const req = store.delete(examId);
        req.onsuccess = () => resolve(true);
        req.onerror = (e) => reject(e.target.error);
    });
}

// ---------------------- CANDIDATE OPERATIONS ----------------------

export async function bulkAddCandidates(examId, candidateList) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('candidates', 'readwrite');
        const store = tx.objectStore('candidates');

        let insertedCount = 0;
        candidateList.forEach((c, idx) => {
            const row = {
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
                createdAt: new Date().toISOString()
            };
            store.add(row);
            insertedCount++;
        });

        tx.oncomplete = () => resolve(insertedCount);
        tx.onerror = (e) => reject(e.target.error);
    });
}

export async function getCandidatesByExam(examId) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('candidates', 'readonly');
        const store = tx.objectStore('candidates');
        const index = store.index('examId');
        const req = index.getAll(examId);
        req.onsuccess = () => {
            const list = req.result || [];
            // Sort by Sesi ascending, lalu No ascending
            list.sort((a, b) => {
                if (a.sesi !== b.sesi) return a.sesi - b.sesi;
                return (Number(a.no) || 0) - (Number(b.no) || 0);
            });
            resolve(list);
        };
        req.onerror = (e) => reject(e.target.error);
    });
}

export async function addCandidate(candidate) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('candidates', 'readwrite');
        const store = tx.objectStore('candidates');
        const req = store.add(candidate);
        req.onsuccess = () => resolve(req.result);
        req.onerror = (e) => reject(e.target.error);
    });
}

export async function updateCandidate(candidate) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('candidates', 'readwrite');
        const store = tx.objectStore('candidates');
        const req = store.put(candidate);
        req.onsuccess = () => resolve(candidate);
        req.onerror = (e) => reject(e.target.error);
    });
}

export async function deleteCandidate(candidateId) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('candidates', 'readwrite');
        const store = tx.objectStore('candidates');
        const req = store.delete(candidateId);
        req.onsuccess = () => resolve(true);
        req.onerror = (e) => reject(e.target.error);
    });
}

export async function deleteCandidatesByExam(examId) {
    const db = await openDatabase();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('candidates', 'readwrite');
        const store = tx.objectStore('candidates');
        const index = store.index('examId');
        const req = index.openKeyCursor(IDBKeyRange.only(examId));

        req.onsuccess = (event) => {
            const cursor = event.target.result;
            if (cursor) {
                store.delete(cursor.primaryKey);
                cursor.continue();
            }
        };

        tx.oncomplete = () => resolve(true);
        tx.onerror = (e) => reject(e.target.error);
    });
}

export async function deleteCandidatesByNips(examId, nipList) {
    if (!nipList || nipList.length === 0) return 0;
    const db = await openDatabase();
    const nipSet = new Set(nipList.map(n => String(n).trim()));

    return new Promise((resolve, reject) => {
        const tx = db.transaction('candidates', 'readwrite');
        const store = tx.objectStore('candidates');
        const index = store.index('examId');
        const req = index.openCursor(IDBKeyRange.only(examId));
        let deletedCount = 0;

        req.onsuccess = (event) => {
            const cursor = event.target.result;
            if (cursor) {
                const candidate = cursor.value;
                if (nipSet.has(String(candidate.nip).trim())) {
                    store.delete(cursor.primaryKey);
                    deletedCount++;
                }
                cursor.continue();
            }
        };

        tx.oncomplete = () => resolve(deletedCount);
        tx.onerror = (e) => reject(e.target.error);
    });
}

export async function getExamStats(examId) {
    const list = await getCandidatesByExam(examId);
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

