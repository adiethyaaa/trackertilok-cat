/**
 * db.js
 * Modul Database Realtime Online untuk "Profiling ASN"
 * Sepenuhnya ditenagai oleh Firebase Realtime Database (menggantikan IndexedDB).
 * Seluruh data ujian & presensi peserta terintegrasi langsung secara online dan realtime.
 */

import {
    initFirebaseService,
    getAllExamsCloud,
    getExamCloud,
    saveExamToCloud,
    deleteExamFromCloud,
    getCandidatesByExamCloud,
    bulkAddCandidatesToCloud,
    addOrUpdateCandidateCloud,
    deleteCandidateCloud,
    deleteCandidatesByExamCloud,
    deleteCandidatesByNipsCloud,
    getExamStatsCloud,
    isCloudActive,
    getMasterInstansiCloud,
    saveMasterInstansiCloud,
    updateInstansiPinCloud,
    getSecurityConfigCloud,
    saveSecurityConfigCloud
} from './firebaseService.js';

/**
 * Memastikan koneksi database online siap digunakan
 */
export async function openDatabase() {
    initFirebaseService();
    return true;
}

// ---------------------- EXAM OPERATIONS (ONLINE REALTIME) ----------------------

export async function addExam(examData) {
    return await saveExamToCloud(examData);
}

export async function getAllExams() {
    return await getAllExamsCloud();
}

export async function getExamById(id) {
    return await getExamCloud(id);
}

export async function updateExam(examData) {
    return await saveExamToCloud(examData);
}

export async function deleteExam(examId) {
    return await deleteExamFromCloud(examId);
}

// ---------------------- CANDIDATE OPERATIONS (ONLINE REALTIME) ----------------------

export async function bulkAddCandidates(examId, candidateList) {
    return await bulkAddCandidatesToCloud(examId, candidateList);
}

export async function getCandidatesByExam(examId) {
    return await getCandidatesByExamCloud(examId);
}

export async function addCandidate(candidate) {
    return await addOrUpdateCandidateCloud(candidate.examId, candidate);
}

export async function updateCandidate(candidate) {
    return await addOrUpdateCandidateCloud(candidate.examId, candidate);
}

// Alias saveCandidate untuk kompatibilitas
export const saveCandidate = updateCandidate;

export async function deleteCandidate(candidateId, examId = null) {
    if (examId) {
        return await deleteCandidateCloud(examId, candidateId);
    }
    // Jika examId tidak diteruskan, coba dapatkan dari candidateId atau window state
    const currentExamId = (typeof window !== 'undefined' && window.currentExam) ? window.currentExam.id : null;
    return await deleteCandidateCloud(currentExamId, candidateId);
}

export async function deleteCandidatesByExam(examId) {
    return await deleteCandidatesByExamCloud(examId);
}

export async function deleteCandidatesByNips(examId, nipList) {
    return await deleteCandidatesByNipsCloud(examId, nipList);
}

export async function getExamStats(examId) {
    return await getExamStatsCloud(examId);
}

// ---------------------- MASTER INSTANSI & PIN SECURITY (DATABASE-DRIVEN) ----------------------

let cachedMasterInstansi = null;
let cachedSecurity = null;

/**
 * Inisialisasi awal Master Instansi & Keamanan di Database
 * Jika di database belum ada, inisialisasi dari default data sekali saja.
 */
export async function initMasterSecurityInDb(defaultInstansiList = []) {
    try {
        let dbList = await getMasterInstansiCloud();
        if (!dbList || dbList.length === 0) {
            if (defaultInstansiList && defaultInstansiList.length > 0) {
                await saveMasterInstansiCloud(defaultInstansiList);
                dbList = defaultInstansiList;
            }
        }
        cachedMasterInstansi = dbList || [];

        let sec = await getSecurityConfigCloud();
        if (!sec || !sec.adminPin || !sec.superAdminPin) {
            sec = {
                adminPin: '1414',
                superAdminPin: '141414',
                updatedAt: new Date().toISOString()
            };
            await saveSecurityConfigCloud(sec);
        }
        cachedSecurity = sec;
        return { masterInstansi: cachedMasterInstansi, security: cachedSecurity };
    } catch (e) {
        console.warn("Gagal inisialisasi master security di DB:", e);
        return { masterInstansi: cachedMasterInstansi || [], security: cachedSecurity || {} };
    }
}

export async function getMasterInstansiFromDb() {
    try {
        const list = await getMasterInstansiCloud();
        if (list && list.length > 0) {
            cachedMasterInstansi = list;
            return list;
        }
    } catch (e) {
        console.warn("Gagal getMasterInstansiCloud, menggunakan cache:", e);
    }
    return cachedMasterInstansi || [];
}

export async function saveMasterInstansiToDb(list) {
    cachedMasterInstansi = list;
    try {
        await saveMasterInstansiCloud(list);
    } catch (e) {
        console.warn("Gagal saveMasterInstansiCloud:", e);
    }
    return list;
}

export async function updateInstansiPinInDb(instansiName, newPin) {
    try {
        const updated = await updateInstansiPinCloud(instansiName, newPin);
        cachedMasterInstansi = updated;
        return updated;
    } catch (e) {
        console.warn("Gagal updateInstansiPinCloud:", e);
        if (cachedMasterInstansi) {
            const clean = String(instansiName).trim().toLowerCase();
            const itm = cachedMasterInstansi.find(i => String(i.name || '').trim().toLowerCase() === clean);
            if (itm) itm.pin = String(newPin).trim();
        }
    }
    return cachedMasterInstansi;
}

export async function getInstansiPinFromDb(instansiName) {
    if (!instansiName) return null;
    const clean = String(instansiName).trim().toLowerCase();
    const list = await getMasterInstansiFromDb();
    const found = list.find(i => String(i.name || '').trim().toLowerCase() === clean);
    return found ? (found.pin || '') : '';
}

/**
 * Validasi PIN Admin Menu Terkunci langsung ke Database (Bebas Hardcode)
 */
export async function verifyAdminPin(enteredPin) {
    if (!enteredPin) return false;
    const cleanInput = String(enteredPin).trim();
    try {
        const sec = (await getSecurityConfigCloud()) || cachedSecurity;
        if (sec && sec.adminPin) {
            if (cleanInput === String(sec.adminPin).trim() || cleanInput === String(sec.superAdminPin).trim()) {
                return true;
            }
        }
    } catch (e) {
        console.warn("Error verifikasi PIN admin:", e);
    }
    if (cachedSecurity && cachedSecurity.adminPin) {
        return cleanInput === String(cachedSecurity.adminPin).trim() || cleanInput === String(cachedSecurity.superAdminPin).trim();
    }
    return false;
}

/**
 * Validasi PIN Super Admin langsung ke Database (Bebas Hardcode)
 */
export async function verifySuperAdminPin(enteredPin) {
    if (!enteredPin) return false;
    const cleanInput = String(enteredPin).trim();
    try {
        const sec = (await getSecurityConfigCloud()) || cachedSecurity;
        if (sec && sec.superAdminPin) {
            return cleanInput === String(sec.superAdminPin).trim();
        }
    } catch (e) {
        console.warn("Error verifikasi PIN super admin:", e);
    }
    if (cachedSecurity && cachedSecurity.superAdminPin) {
        return cleanInput === String(cachedSecurity.superAdminPin).trim();
    }
    return false;
}

