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
    isCloudActive
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
