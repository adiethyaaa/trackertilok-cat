/**
 * examManager.js
 * Modul untuk manajemen pembuatan Ujian, pengelolaan pilihan wilayah/instansi
 * Papua Barat & Papua Barat Daya, serta koordinasi data ujian aktif.
 */

import { masterInstansiData } from '../masterInstansi.js';
import * as db from './db.js';

let currentSelectedExamId = null;

/**
 * Mengambil daftar instansi yang dikelompokkan berdasarkan Wilayah
 */
export function getGroupedInstansi() {
    const papuaBarat = masterInstansiData.filter(i => i.wilker === 'Papua Barat');
    const papuaBaratDaya = masterInstansiData.filter(i => i.wilker === 'Papua Barat Daya');
    const vertikal = masterInstansiData.filter(i => i.wilker === 'Instansi Vertikal');

    return {
        'Papua Barat': papuaBarat,
        'Papua Barat Daya': papuaBaratDaya,
        'Instansi Vertikal': vertikal
    };
}

/**
 * Mengisi pilihan dropdown instansi pada form Create Ujian
 */
export function populateInstansiDropdown(selectElementId) {
    const select = document.getElementById(selectElementId);
    if (!select) return;

    const grouped = getGroupedInstansi();
    let html = '<option value="">-- Pilih Instansi --</option>';

    // Group Papua Barat
    if (grouped['Papua Barat'].length > 0) {
        html += `<optgroup label="📍 Papua Barat">`;
        grouped['Papua Barat'].forEach(item => {
            html += `<option value="${item.name}" data-region="Papua Barat">${item.name}</option>`;
        });
        html += `</optgroup>`;
    }

    // Group Papua Barat Daya
    if (grouped['Papua Barat Daya'].length > 0) {
        html += `<optgroup label="📍 Papua Barat Daya">`;
        grouped['Papua Barat Daya'].forEach(item => {
            html += `<option value="${item.name}" data-region="Papua Barat Daya">${item.name}</option>`;
        });
        html += `</optgroup>`;
    }

    // Group Instansi Vertikal
    if (grouped['Instansi Vertikal'].length > 0) {
        html += `<optgroup label="🏢 Instansi Vertikal">`;
        grouped['Instansi Vertikal'].forEach(item => {
            html += `<option value="${item.name}" data-region="Instansi Vertikal">${item.name}</option>`;
        });
        html += `</optgroup>`;
    }

    select.innerHTML = html;
}

/**
 * Mendapatkan ID Ujian yang sedang aktif/dipilih
 */
export function getSelectedExamId() {
    if (!currentSelectedExamId) {
        currentSelectedExamId = localStorage.getItem('selected_exam_id_profiling');
    }
    return currentSelectedExamId;
}

/**
 * Mengubah Ujian aktif
 */
export function setSelectedExamId(id) {
    currentSelectedExamId = id;
    if (id) {
        localStorage.setItem('selected_exam_id_profiling', id);
    } else {
        localStorage.removeItem('selected_exam_id_profiling');
    }
}

/**
 * Simpan Ujian Baru (Hanya memerlukan Instansi, Tanggal, Tilok, Kuota, Catatan)
 */
export async function createNewExam(formData) {
    if (!formData.instansi) {
        throw new Error("Instansi wajib dipilih!");
    }

    formData.title = formData.title || formData.instansi;

    const newExam = await db.addExam(formData);
    setSelectedExamId(newExam.id);
    return newExam;
}
