/**
 * app.js
 * Controller Utama Aplikasi "Profiling ASN"
 * Mengintegrasikan IndexedDB, Excel Handler, Session Rules, dan Manajemen Wilayah Papua Barat & PB Daya
 */

import { masterInstansiData, toTitleCase, getInstansiPin } from '../masterInstansi.js';
import * as db from './db.js';
import { parseFlexibleDate, isFriday, getSessionTime, formatDateDisplay, getDayNameID, formatCumulativeSessionNumber } from './sessionRules.js';
import { 
    parseExcelFile, 
    downloadExcelTemplate, 
    downloadSystemTemplate, 
    downloadScheduleTemplate, 
    exportCandidatesToExcel, 
    analyzeDuplicates, 
    mergeScheduleWithExisting 
} from './excelHandler.js';
import { populateInstansiDropdown, getSelectedExamId, setSelectedExamId, createNewExam } from './examManager.js';
import {
    initFirebaseService,
    getFirebaseConfig,
    saveFirebaseConfig,
    removeFirebaseConfig,
    onConnectionStatusChange,
    isCloudActive,
    listenExamsCloud,
    listenCandidatesCloud,
    saveExamToCloud,
    deleteExamFromCloud,
    bulkAddCandidatesToCloud,
    updateAttendanceInCloud
} from './firebaseService.js';

// State Aplikasi
let currentExam = null;
let allExams = [];
let currentCandidates = [];
let filteredCandidates = [];
let previewParsedData = null;
let currentSessionFilter = 'ALL';
let currentCumulativeSessionFilter = 'ALL';
let currentSearchTerm = '';
let currentDateFilter = 'ALL';
let currentKelJabatanFilter = 'ALL';
let selectedDashboardDates = new Set();
let currentRekapData = { tab1Rows: [], tab2Rows: [] };
let currentSortColumn = 'sesi';
let currentSortDirection = 'asc';
let isPinAuthorized = false;
let isSuperAdmin = false;
let pendingTargetTab = null;
let pendingActionAfterPin = null;
let activeCandidatesUnsubscribe = null;

// Inisialisasi Aplikasi Saat Halaman Dimuat
document.addEventListener('DOMContentLoaded', async () => {
    initLiveClockWIT();
    setupTabNavigation();
    setupCreateExamForm();
    setupEditExamForm();
    setupExcelUpload();
    setupManualCandidateForm();
    setupMasterInstansiUI();
    setupFirebaseIntegration();

    // Isi dropdown instansi
    populateInstansiDropdown('selectExamInstansi');

    // Jika belum ada sesi PIN yang terotorisasi, langsung munculkan pop-up PIN tanpa menunggu network
    if (!checkSuperAdminSession() && !getActiveExamSession()) {
        window.openSelectExamWithPinModal(null, false);
    }

    // Muat data ujian (hanya metadata ujian, TANPA data peserta)
    await loadInitialData();

    // Set default tab ke Jadwal & Peserta (Tab paling awal/kiri)
    switchTab('daftar-peserta');

    // Refresh icon Lucide
    if (window.lucide) {
        window.lucide.createIcons();
    }
});

/**
 * Live Clock Waktu Indonesia Timur (WIT / UTC+9)
 */
function initLiveClockWIT() {
    const clockEl = document.getElementById('liveClockWIT');
    const dateEl = document.getElementById('liveDateWIT');

    function updateTime() {
        const now = new Date();
        const utc = now.getTime() + (now.getTimezoneOffset() * 60000);
        const witDate = new Date(utc + (3600000 * 9));

        const hours = String(witDate.getHours()).padStart(2, '0');
        const minutes = String(witDate.getMinutes()).padStart(2, '0');
        const seconds = String(witDate.getSeconds()).padStart(2, '0');

        if (clockEl) {
            clockEl.textContent = `${hours}:${minutes}:${seconds} WIT`;
        }

        if (dateEl) {
            const dayName = getDayNameID(witDate);
            const formatted = formatDateDisplay(witDate, 'long');
            dateEl.textContent = `${dayName}, ${formatted}`;
        }
    }

    updateTime();
    setInterval(updateTime, 1000);
}

const EXAM_SESSION_KEY = 'profiling_asn_active_session';

export function checkSuperAdminSession() {
    return sessionStorage.getItem('is_super_admin') === 'true';
}

export function setSuperAdminSession(active) {
    if (active) {
        sessionStorage.setItem('is_super_admin', 'true');
        sessionStorage.setItem('is_admin_pin_authorized', 'true');
        isSuperAdmin = true;
        isPinAuthorized = true;
    } else {
        sessionStorage.removeItem('is_super_admin');
        sessionStorage.removeItem('is_admin_pin_authorized');
        isSuperAdmin = false;
    }
}

export function getActiveExamSession() {
    try {
        const raw = sessionStorage.getItem(EXAM_SESSION_KEY);
        if (raw) {
            return JSON.parse(raw);
        }
    } catch (e) {
        console.warn("Gagal membaca active session:", e);
    }
    return null;
}

export function saveActiveExamSession(examId, instansi) {
    try {
        sessionStorage.setItem(EXAM_SESSION_KEY, JSON.stringify({
            examId: examId,
            instansi: instansi,
            timestamp: Date.now()
        }));
    } catch (e) {
        console.warn("Gagal menyimpan active session:", e);
    }
}

export function clearActiveExamSession() {
    try {
        sessionStorage.removeItem(EXAM_SESSION_KEY);
    } catch (e) {
        console.warn("Gagal menghapus active session:", e);
    }
}

/**
 * Memuat data awal daftar ujian dari Firebase Realtime Database
 * HANYA memuat metadata instansi ujian, TIDAK memuat data peserta sebelum PIN diinput
 */
async function loadInitialData() {
    try {
        allExams = await db.getAllExams();

        renderExamSelectDropdowns();
        renderExamListInCreateTab();
        refreshModalSelectExamPicker();

        // 1. Cek apakah Super Admin sudah login di sesi ini
        if (checkSuperAdminSession()) {
            isSuperAdmin = true;
            isPinAuthorized = true;
            const activeId = getSelectedExamId() || (allExams.length > 0 ? allExams[0].id : null);
            if (activeId) {
                await setActiveExam(activeId);
            }
            return;
        }

        const session = getActiveExamSession();

        // 2. Validasi apakah sesi yang tersimpan masih valid di database ujian
        if (session && session.examId && allExams.some(e => e.id === session.examId)) {
            await setActiveExam(session.examId);
        } else {
            // Belum ada sesi PIN di browser ini (pertama kali buka / browser baru dibuka)
            clearActiveExamSession();
            currentExam = null;
            currentCandidates = [];
            renderExamSelectDropdowns();
            renderDashboardExamInfo();
            renderDashboardStats();
            applyCandidateFilters();

            // Munculkan pop up pilih ujian aktif & isi PIN instansi
            if (allExams.length > 0) {
                window.openSelectExamWithPinModal(null, false);
            }
        }

    } catch (err) {
        console.error("Gagal memuat data awal ujian:", err);
        showToast("Terjadi kendala memuat daftar ujian: " + err.message, "error");
    }
}

/**
 * Mengubah Ujian Aktif dan memperbarui seluruh tampilan
 * HANYA mengambil data kandidat jika instansi telah dipilih & diotorisasi PIN
 */
async function setActiveExam(examId) {
    if (!examId) {
        setSelectedExamId(null);
        currentExam = null;
        currentCandidates = [];
        if (activeCandidatesUnsubscribe) {
            activeCandidatesUnsubscribe();
            activeCandidatesUnsubscribe = null;
        }

        const selectNav = document.getElementById('selectActiveExamNavbar');
        if (selectNav) selectNav.value = '';
        const selectUpload = document.getElementById('selectUploadTargetExam');
        if (selectUpload) selectUpload.value = '';

        renderDashboardExamInfo();
        renderExamListInCreateTab();
        renderDashboardStats();
        populatePelaksanaanFilterDropdown();
        populateSesiFilterDropdown('ALL');
        applyCandidateFilters();
        return;
    }

    // Pastikan user berhak mengakses ujian ini (Super Admin atau Sesi PIN Instansi Cocok)
    const session = getActiveExamSession();
    const isAuthorized = isSuperAdmin || (session && session.examId === examId);

    if (!isAuthorized) {
        console.warn("Akses ke ujian belum diverifikasi PIN:", examId);
        currentExam = null;
        currentCandidates = [];
        if (activeCandidatesUnsubscribe) {
            activeCandidatesUnsubscribe();
            activeCandidatesUnsubscribe = null;
        }
        applyCandidateFilters();
        window.openSelectExamWithPinModal(examId, Boolean(currentExam));
        return;
    }

    setSelectedExamId(examId);
    currentExam = allExams.find(e => e.id === examId) || null;
    window.currentExam = currentExam;
    selectedDashboardDates.clear();

    // Perbarui dropdown navbar & upload
    const selectNav = document.getElementById('selectActiveExamNavbar');
    if (selectNav) selectNav.value = examId || '';

    const selectUpload = document.getElementById('selectUploadTargetExam');
    if (selectUpload) selectUpload.value = examId || '';

    // Render ulang info banner & list di tab create
    renderDashboardExamInfo();
    renderExamListInCreateTab();

    // HANYA ambil data kandidat jika instansi sudah terpilih dan diotorisasi PIN
    if (currentExam) {
        showTableLoading(currentExam.instansi);

        try {
            currentCandidates = await db.getCandidatesByExam(currentExam.id);
        } catch (err) {
            console.error("Gagal memuat kandidat:", err);
            currentCandidates = [];
            showToast("Gagal memuat data peserta: " + err.message, "error");
        }

        // Pasang Realtime Listener jika Firebase Cloud aktif (HANYA untuk instansi ini)
        if (isCloudActive()) {
            if (activeCandidatesUnsubscribe) {
                activeCandidatesUnsubscribe();
                activeCandidatesUnsubscribe = null;
            }
            activeCandidatesUnsubscribe = listenCandidatesCloud(currentExam.id, (cloudCandidates) => {
                if (cloudCandidates) {
                    currentCandidates = cloudCandidates;
                    renderDashboardStats();
                    populatePelaksanaanFilterDropdown();
                    populateSesiFilterDropdown(currentDateFilter || 'ALL');
                    populateKelJabatanFilterDropdown();
                    applyCandidateFilters();
                }
            });
        }
    } else {
        currentCandidates = [];
        if (activeCandidatesUnsubscribe) {
            activeCandidatesUnsubscribe();
            activeCandidatesUnsubscribe = null;
        }
    }

    renderDashboardStats();
    populatePelaksanaanFilterDropdown();
    populateSesiFilterDropdown('ALL');
    populateKelJabatanFilterDropdown();
    applyCandidateFilters();
}

/**
 * Tampilkan indikator loading tabel saat sedang mengambil data peserta instansi terpilih
 */
function showTableLoading(instansiName) {
    const tbody = document.getElementById('tbodyCandidateList');
    if (tbody) {
        tbody.innerHTML = `
            <tr>
                <td colspan="10" class="p-8 text-center text-slate-500">
                    <div class="inline-block animate-spin rounded-full h-6 w-6 border-b-2 border-bkn-700 mb-2"></div>
                    <p class="font-bold text-slate-700 text-sm">Memuat data peserta ${instansiName}...</p>
                    <p class="text-xs text-slate-400 mt-1">Mengambil data dari database online...</p>
                </td>
            </tr>
        `;
        if (window.lucide) window.lucide.createIcons();
    }
}

/**
 * Render opsi pada dropdown ujian (Hanya menampilkan nama Instansi yang sudah di-create)
 */
function renderExamSelectDropdowns() {
    const selectNav = document.getElementById('selectActiveExamNavbar');
    const selectUpload = document.getElementById('selectUploadTargetExam');

    if (allExams.length === 0) {
        const emptyHtml = `<option value="">Belum ada instansi ujian</option>`;
        if (selectNav) selectNav.innerHTML = emptyHtml;
        if (selectUpload) selectUpload.innerHTML = `<option value="">-- Belum ada instansi yang dibuat --</option>`;
        return;
    }

    // Hanya tampilkan nama Instansi pada pilihan
    const optionsHtml = allExams.map(e => `
        <option value="${e.id}" class="text-slate-800">${e.instansi}</option>
    `).join('');

    if (selectNav) {
        selectNav.innerHTML = `<option value="" class="text-slate-800">-- Pilih Instansi Ujian --</option>` + optionsHtml;
        selectNav.value = currentExam ? currentExam.id : '';
        selectNav.onchange = (e) => {
            const val = e.target.value;
            if (val) {
                if (currentExam && currentExam.id === val) return;
                // Jika Super Admin, bebas ganti ujian tanpa minta PIN
                if (isSuperAdmin) {
                    setActiveExam(val);
                    return;
                }
                window.openSelectExamWithPinModal(val, Boolean(currentExam));
            } else {
                selectNav.value = currentExam ? currentExam.id : '';
            }
        };
    }

    if (selectUpload) {
        selectUpload.innerHTML = `<option value="">-- Pilih Instansi Penerima Data --</option>` + optionsHtml;
        if (currentExam) selectUpload.value = currentExam.id;
    }
}

/**
 * Render info ujian di banner Dashboard
 */
function renderDashboardExamInfo() {
    const instansiTitleEl = document.getElementById('dashExamInstansiTitle');
    const regionBadge = document.getElementById('dashRegionBadge');
    const dateRangeEl = document.getElementById('dashDateRange');
    const locationEl = document.getElementById('dashExamLocation');

    if (!currentExam) {
        if (instansiTitleEl) instansiTitleEl.textContent = 'Belum Ada Ujian Aktif';
        if (regionBadge) regionBadge.textContent = 'Status';
        if (dateRangeEl) dateRangeEl.textContent = '-';
        if (locationEl) locationEl.innerHTML = '<i data-lucide="map-pin" class="w-4 h-4 text-slate-400 inline"></i> <span>Silakan buat atau pilih ujian terlebih dahulu</span>';
        return;
    }

    if (instansiTitleEl) instansiTitleEl.textContent = currentExam.instansi;
    if (regionBadge) {
        regionBadge.textContent = currentExam.wilker || 'Instansi Terdaftar';
        if (currentExam.wilker === 'Papua Barat Daya') {
            regionBadge.className = 'text-xs font-bold px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-800';
        } else if (currentExam.wilker === 'Instansi Vertikal') {
            regionBadge.className = 'text-xs font-bold px-2.5 py-0.5 rounded-full bg-amber-100 text-amber-800';
        } else {
            regionBadge.className = 'text-xs font-bold px-2.5 py-0.5 rounded-full bg-blue-100 text-blue-800';
        }
    }

    if (dateRangeEl) {
        const start = currentExam.startDate ? formatDateDisplay(currentExam.startDate, 'short') : '-';
        const end = currentExam.endDate ? formatDateDisplay(currentExam.endDate, 'short') : start;
        dateRangeEl.textContent = `${start} s.d. ${end}`;
    }

    if (locationEl) {
        locationEl.innerHTML = `
            <i data-lucide="map-pin" class="w-4 h-4 text-slate-400 inline mr-1"></i>
            <span>Tilok: <strong>${currentExam.location}</strong> (Kapasitas: ${currentExam.quotaPerSession} PC/Sesi)</span>
        `;
    }

    if (window.lucide) window.lucide.createIcons();
}

/**
 * Buka/Tutup dropdown panel filter tanggal di Dashboard
 */
window.toggleDashboardDateDropdown = () => {
    const panel = document.getElementById('dropdownPanelFilterTanggalDash');
    if (panel) {
        panel.classList.toggle('hidden');
    }
};

/**
 * Handler saat checkbox master "Ceklis Semua Tanggal" diubah
 */
window.onToggleAllDashboardDates = (isChecked) => {
    const uniqueDates = getSortedExamDates();
    if (isChecked) {
        selectedDashboardDates = new Set(uniqueDates);
    } else {
        selectedDashboardDates.clear();
    }
    updateDashboardDateFilterUI();
    renderDashboardStats();
};

/**
 * Handler saat checkbox tanggal individual diubah
 */
window.onToggleSingleDashboardDate = (dateVal, isChecked) => {
    const uniqueDates = getSortedExamDates();
    if (selectedDashboardDates.size === 0) {
        selectedDashboardDates = new Set(uniqueDates);
    }

    if (isChecked) {
        selectedDashboardDates.add(dateVal);
    } else {
        selectedDashboardDates.delete(dateVal);
    }

    updateDashboardDateFilterUI();
    renderDashboardStats();
};

/**
 * Update UI Filter Tanggal Dashboard (Label Tombol, Master Checkbox, Badge Count)
 */
function updateDashboardDateFilterUI() {
    const uniqueDates = getSortedExamDates();
    const labelEl = document.getElementById('labelDropdownFilterTanggalDash');
    const masterCb = document.getElementById('cbDashTanggalAll');
    const countBadge = document.getElementById('countSelectedDashDates');

    const totalDays = uniqueDates.length;
    const isAllSelected = selectedDashboardDates.size === 0 || selectedDashboardDates.size === totalDays;

    if (masterCb) {
        masterCb.checked = isAllSelected && totalDays > 0;
        masterCb.indeterminate = !isAllSelected && selectedDashboardDates.size > 0;
    }

    if (countBadge) {
        const count = isAllSelected ? totalDays : selectedDashboardDates.size;
        countBadge.textContent = `${count} Hari`;
    }

    if (labelEl) {
        if (totalDays === 0) {
            labelEl.textContent = '-- Belum Ada Tanggal Ujian --';
        } else if (isAllSelected) {
            labelEl.textContent = `Semua Tanggal Pelaksanaan (${totalDays} Hari)`;
        } else if (selectedDashboardDates.size === 1) {
            const onlyDate = Array.from(selectedDashboardDates)[0];
            labelEl.textContent = `1 Tanggal: ${onlyDate}`;
        } else {
            labelEl.textContent = `${selectedDashboardDates.size} Tanggal Terpilih`;
        }
    }
}

/**
 * Mengisi daftar checkbox tanggal di Dashboard
 */
function populateDashboardFilterTanggalDropdown() {
    const listContainer = document.getElementById('listDashTanggalCheckboxes');
    if (!listContainer) return;

    const uniqueDates = getSortedExamDates();

    // Hapus tanggal lama yang sudah tidak ada
    const validSelected = new Set();
    selectedDashboardDates.forEach(d => {
        if (uniqueDates.includes(d)) validSelected.add(d);
    });
    selectedDashboardDates = validSelected;

    const isAllSelected = selectedDashboardDates.size === 0 || selectedDashboardDates.size === uniqueDates.length;

    if (uniqueDates.length === 0) {
        listContainer.innerHTML = `<p class="text-slate-400 italic text-[11px] py-2 text-center">Belum ada tanggal pelaksanaan.</p>`;
        updateDashboardDateFilterUI();
        return;
    }

    let html = '';
    uniqueDates.forEach(d => {
        const isChecked = isAllSelected || selectedDashboardDates.has(d);
        const fri = isFriday(d);
        const totalPesertaDate = currentCandidates.filter(c => c.pelaksanaan === d).length;

        html += `
            <label class="flex items-center justify-between p-1.5 rounded-lg hover:bg-slate-50 cursor-pointer select-none transition border border-transparent hover:border-slate-200">
                <div class="flex items-center space-x-2 truncate">
                    <input type="checkbox" 
                           value="${d}" 
                           ${isChecked ? 'checked' : ''} 
                           onchange="onToggleSingleDashboardDate('${d}', this.checked)"
                           class="w-3.5 h-3.5 text-bkn-700 rounded border-slate-300 focus:ring-bkn-600 cursor-pointer">
                    <span class="font-medium text-slate-700 truncate ${fri ? 'text-indigo-900 font-semibold' : ''}">
                        ${d} ${fri ? '<span class="text-[10px] text-indigo-600 font-bold ml-1">(Jumat)</span>' : ''}
                    </span>
                </div>
                <span class="text-[10px] bg-slate-100 text-slate-600 px-1.5 py-0.5 rounded font-mono flex-shrink-0">
                    ${totalPesertaDate}
                </span>
            </label>
        `;
    });

    listContainer.innerHTML = html;
    updateDashboardDateFilterUI();
}

/**
 * Event listener untuk menutup dropdown tanggal saat klik di luar
 */
document.addEventListener('click', (e) => {
    const btn = document.getElementById('btnDropdownFilterTanggalDash');
    const panel = document.getElementById('dropdownPanelFilterTanggalDash');
    if (panel && !panel.classList.contains('hidden')) {
        if (btn && !btn.contains(e.target) && !panel.contains(e.target)) {
            panel.classList.add('hidden');
        }
    }
});

/**
 * Mengklasifikasikan kandidat ke dalam kategori kelompok jabatan yang baku
 */
function classifyCandidateKelompok(c) {
    const rawK = String(c.kelJabatan || '').trim();
    if (!rawK || rawK === '-' || rawK.toUpperCase() === 'NULL' || rawK.toLowerCase() === 'belum terdata' || rawK.toLowerCase().includes('belum terdaftar')) {
        return {
            category: 'KOSONG',
            label: 'Belum Terdata / Kosong',
            sub: null,
            isWarning: true
        };
    }
    const kLower = rawK.toLowerCase();
    const jLower = String(c.jabatan || '').toLowerCase();
    const combined = `${kLower} ${jLower}`;

    // 1. JPT Pratama / Tinggi
    if (kLower.includes('jpt') || kLower.includes('pratama') || (kLower.includes('tinggi') && !kLower.includes('fungsional'))) {
        return { category: 'JPT_PRATAMA', label: 'JPT Pratama', sub: null, isWarning: false };
    }
    // 2. Administrator
    if (kLower.includes('administrator')) {
        return { category: 'ADMINISTRATOR', label: 'Administrator', sub: null, isWarning: false };
    }
    // 3. Eselon V
    if (kLower.includes('eselon v') || kLower.includes('eselon 5')) {
        return { category: 'ESELON_V', label: 'Eselon V', sub: null, isWarning: false };
    }
    // 4. Pengawas
    if (kLower.includes('pengawas') || kLower.includes('eselon iv') || kLower.includes('eselon 4')) {
        return { category: 'PENGAWAS', label: 'Pengawas', sub: null, isWarning: false };
    }
    // 5. Jabatan Fungsional (beserta jenjang anak/child)
    const isJF = kLower.includes('fungsional') || kLower.includes('jf') || 
                 combined.includes('terampil') || combined.includes('mahir') || combined.includes('penyelia') ||
                 combined.includes('pertama') || combined.includes('muda') || (combined.includes('madya') && !combined.includes('jpt'));
    
    if (isJF) {
        let sub = 'Fungsional Lainnya';
        if (combined.includes('terampil')) sub = 'Terampil';
        else if (combined.includes('mahir')) sub = 'Mahir';
        else if (combined.includes('penyelia')) sub = 'Penyelia';
        else if (combined.includes('pertama')) sub = 'Ahli Pertama';
        else if (combined.includes('muda')) sub = 'Ahli Muda';
        else if (combined.includes('madya')) sub = 'Ahli Madya';

        return { category: 'FUNGSIONAL', label: 'Jab. Fungsional', sub: sub, isWarning: false };
    }
    // 6. Pelaksana
    if (kLower.includes('pelaksana') || kLower.includes('staf')) {
        return { category: 'PELAKSANA', label: 'Pelaksana', sub: null, isWarning: false };
    }

    return { category: 'LAINNYA', label: rawK, sub: null, isWarning: false };
}

/**
 * Render statistik di Dashboard (termasuk Statistik Kehadiran, Filter Tanggal, & Hierarki Kel. Jabatan)
 */
function renderDashboardStats() {
    // 1. Populate opsi filter tanggal di dashboard
    populateDashboardFilterTanggalDropdown();

    const statTotal = document.getElementById('statTotalPeserta');
    const statS1 = document.getElementById('statSesi1');
    const statS2 = document.getElementById('statSesi2');
    const statS3 = document.getElementById('statSesi3');
    const distContainer = document.getElementById('dashboardDistributionContainer');

    // Filter kandidat berdasarkan tanggal dashboard yang dipilih
    const uniqueDates = getSortedExamDates();
    const isAllDatesSelected = selectedDashboardDates.size === 0 || selectedDashboardDates.size === uniqueDates.length;
    const dashboardCandidates = isAllDatesSelected
        ? currentCandidates
        : currentCandidates.filter(c => selectedDashboardDates.has(c.pelaksanaan));

    const total = dashboardCandidates.length;
    const s1 = dashboardCandidates.filter(c => c.sesi === 1).length;
    const s2 = dashboardCandidates.filter(c => c.sesi === 2).length;
    const s3 = dashboardCandidates.filter(c => c.sesi === 3).length;

    // Hitung Kehadiran
    const hadir = dashboardCandidates.filter(c => c.kehadiran === 'HADIR').length;
    const tidakHadir = dashboardCandidates.filter(c => c.kehadiran === 'TIDAK_HADIR').length;
    const belumPresensi = Math.max(0, total - hadir - tidakHadir);

    const hadirPct = total > 0 ? Math.round((hadir / total) * 100) : 0;
    const tidakHadirPct = total > 0 ? Math.round((tidakHadir / total) * 100) : 0;
    const belumPct = total > 0 ? Math.round((belumPresensi / total) * 100) : 0;

    // Update Kartu Kehadiran di Dashboard
    const elHadir = document.getElementById('dashStatHadir');
    const elHadirPct = document.getElementById('dashStatHadirPct');
    const elTidakHadir = document.getElementById('dashStatTidakHadir');
    const elTidakHadirPct = document.getElementById('dashStatTidakHadirPct');
    const elBelum = document.getElementById('dashStatBelumPresensi');
    const elBelumPct = document.getElementById('dashStatBelumPct');

    if (elHadir) elHadir.textContent = hadir;
    if (elHadirPct) elHadirPct.textContent = `${hadirPct}% dari total ${total} peserta`;
    if (elTidakHadir) elTidakHadir.textContent = tidakHadir;
    if (elTidakHadirPct) elTidakHadirPct.textContent = `${tidakHadirPct}% dari total ${total} peserta`;
    if (elBelum) elBelum.textContent = belumPresensi;
    if (elBelumPct) elBelumPct.textContent = `${belumPct}% belum presensi`;

    // Update Kartu Statistik Sesi
    if (statTotal) statTotal.textContent = total;
    if (statS1) statS1.textContent = s1;
    if (statS2) statS2.textContent = s2;
    if (statS3) statS3.textContent = s3;

    // Filter badge counts di Tab Jadwal (selalu mencerminkan total keseluruhan ujian aktif)
    const grandTotal = currentCandidates.length;
    const grandS0 = currentCandidates.filter(c => !c.sesi || c.sesi === 'NULL' || c.sesi === '00' || c.sesi === 0).length;
    const grandS1 = currentCandidates.filter(c => Number(c.sesi) === 1).length;
    const grandS2 = currentCandidates.filter(c => Number(c.sesi) === 2).length;
    const grandS3 = currentCandidates.filter(c => Number(c.sesi) === 3).length;

    const cfAll = document.getElementById('countFilterAll');
    const cf0 = document.getElementById('countFilter0');
    const cf1 = document.getElementById('countFilter1');
    const cf2 = document.getElementById('countFilter2');
    const cf3 = document.getElementById('countFilter3');
    if (cfAll) cfAll.textContent = grandTotal;
    if (cf0) cf0.textContent = grandS0;
    if (cf1) cf1.textContent = grandS1;
    if (cf2) cf2.textContent = grandS2;
    if (cf3) cf3.textContent = grandS3;

    // Render statistik kehadiran per kelompok jabatan dengan susunan & hierarki yang diminta:
    // 1. JPT Pratama
    // 2. Administrator
    // 3. Pengawas
    // 4. Eselon V
    // 5. Jab. Fungsional (Parent) -> Child: Terampil, Mahir, Penyelia, Ahli Pertama, Ahli Muda, Ahli Madya
    // 6. Belum Terdata / Kosong (Peserta Belum Terdaftar) -> DI ATAS PELAKSANA
    // 7. Pelaksana
    // 8. Lainnya (bila ada)
    const kelJabatanContainer = document.getElementById('dashKelJabatanContainer');
    if (kelJabatanContainer) {
        if (total === 0) {
            kelJabatanContainer.innerHTML = `<p class="text-sm text-slate-500 py-6 text-center">Belum ada data peserta untuk ujian/tanggal ini.</p>`;
        } else {
            // Struktur penampung statistik
            const createStatHolder = () => ({ hadir: 0, tidakHadir: 0, belum: 0, total: 0 });
            const categories = {
                JPT_PRATAMA: { label: 'JPT Pratama', stat: createStatHolder() },
                ADMINISTRATOR: { label: 'Administrator', stat: createStatHolder() },
                PENGAWAS: { label: 'Pengawas', stat: createStatHolder() },
                ESELON_V: { label: 'Eselon V', stat: createStatHolder() },
                FUNGSIONAL: { 
                    label: 'Jab. Fungsional', 
                    stat: createStatHolder(),
                    children: {
                        'Terampil': createStatHolder(),
                        'Mahir': createStatHolder(),
                        'Penyelia': createStatHolder(),
                        'Ahli Pertama': createStatHolder(),
                        'Ahli Muda': createStatHolder(),
                        'Ahli Madya': createStatHolder(),
                        'Fungsional Lainnya': createStatHolder()
                    }
                },
                KOSONG: { label: 'Belum Terdata / Kosong', stat: createStatHolder(), isWarning: true },
                PELAKSANA: { label: 'Pelaksana', stat: createStatHolder() },
                LAINNYA: {} // Dynamic key jika ada kelompok lain di luar standar
            };

            // Hitung agregat tiap kandidat
            dashboardCandidates.forEach(c => {
                const cls = classifyCandidateKelompok(c);
                const att = c.kehadiran;

                const addAtt = (statObj) => {
                    if (att === 'HADIR') statObj.hadir++;
                    else if (att === 'TIDAK_HADIR') statObj.tidakHadir++;
                    else statObj.belum++;
                    statObj.total++;
                };

                if (cls.category === 'FUNGSIONAL') {
                    addAtt(categories.FUNGSIONAL.stat);
                    const subName = cls.sub || 'Fungsional Lainnya';
                    if (!categories.FUNGSIONAL.children[subName]) {
                        categories.FUNGSIONAL.children[subName] = createStatHolder();
                    }
                    addAtt(categories.FUNGSIONAL.children[subName]);
                } else if (categories[cls.category]) {
                    addAtt(categories[cls.category].stat);
                } else {
                    if (!categories.LAINNYA[cls.label]) {
                        categories.LAINNYA[cls.label] = createStatHolder();
                    }
                    addAtt(categories.LAINNYA[cls.label]);
                }
            });

            // Helper render satu baris data tabel
            const renderRow = (label, stat, options = {}) => {
                const { isParent = false, isChild = false, isWarning = false } = options;
                const pct = stat.total > 0 ? Math.round((stat.hadir / stat.total) * 100) : 0;
                
                let rowBg = 'hover:bg-slate-50 transition';
                if (isWarning) {
                    rowBg = 'bg-rose-50/75 border-l-4 border-l-red-900 font-medium transition';
                } else if (isParent) {
                    rowBg = 'bg-indigo-50/40 hover:bg-indigo-50/70 font-bold border-t border-b border-indigo-100/70 transition';
                } else if (isChild) {
                    rowBg = 'bg-slate-50/50 hover:bg-slate-100/60 text-slate-600 transition';
                }

                return `
                    <tr class="${rowBg}">
                        <td class="p-2.5 ${isChild ? 'pl-8 text-xs font-semibold' : 'font-bold'} ${isWarning ? 'text-rose-900 flex items-center gap-1.5' : (isParent ? 'text-indigo-950 flex items-center gap-1.5' : 'text-slate-800')}">
                            ${isWarning ? '<i data-lucide="alert-circle" class="w-3.5 h-3.5 text-rose-700 inline-block flex-shrink-0"></i>' : ''}
                            ${isChild ? '<span class="text-slate-400 font-bold mr-1">↳</span>' : ''}
                            <span>${label}</span>
                            ${isParent ? '<span class="text-[10px] bg-indigo-100 text-indigo-800 font-extrabold px-1.5 py-0.5 rounded uppercase tracking-wider">Parent</span>' : ''}
                            ${isWarning ? '<span class="text-[10px] bg-red-900 text-red-100 px-1.5 py-0.5 rounded font-bold uppercase shadow-xs">Peserta Belum Terdaftar</span>' : ''}
                        </td>
                        <td class="p-2.5 text-center text-emerald-700 font-bold text-sm">${stat.hadir}</td>
                        <td class="p-2.5 text-center text-rose-700 font-bold text-sm">${stat.tidakHadir}</td>
                        <td class="p-2.5 text-center text-amber-700 font-semibold">${stat.belum}</td>
                        <td class="p-2.5 text-center font-bold text-slate-900">${stat.total}</td>
                        <td class="p-2.5 text-center">
                            <div class="flex items-center justify-center gap-2">
                                <div class="w-16 bg-slate-200 rounded-full h-2 overflow-hidden">
                                    <div class="bg-emerald-600 h-2 rounded-full" style="width: ${pct}%"></div>
                                </div>
                                <span class="font-bold text-[11px] ${pct >= 80 ? 'text-emerald-700' : (pct >= 50 ? 'text-amber-700' : 'text-slate-600')}">${pct}%</span>
                            </div>
                        </td>
                    </tr>
                `;
            };

            let rowsHtml = '';

            // 1. JPT Pratama
            if (categories.JPT_PRATAMA.stat.total > 0) {
                rowsHtml += renderRow(categories.JPT_PRATAMA.label, categories.JPT_PRATAMA.stat);
            }

            // 2. Administrator
            if (categories.ADMINISTRATOR.stat.total > 0) {
                rowsHtml += renderRow(categories.ADMINISTRATOR.label, categories.ADMINISTRATOR.stat);
            }

            // 3. Pengawas
            if (categories.PENGAWAS.stat.total > 0) {
                rowsHtml += renderRow(categories.PENGAWAS.label, categories.PENGAWAS.stat);
            }

            // 4. Eselon V
            if (categories.ESELON_V.stat.total > 0) {
                rowsHtml += renderRow(categories.ESELON_V.label, categories.ESELON_V.stat);
            }

            // 5. Jabatan Fungsional (Parent & Child Rows)
            if (categories.FUNGSIONAL.stat.total > 0) {
                // Render Parent Row
                rowsHtml += renderRow(categories.FUNGSIONAL.label, categories.FUNGSIONAL.stat, { isParent: true });
                
                // Susunan baku Child: terampil, mahir, penyelia, ahli pertama, ahli muda, ahli madya
                const standardJfOrder = ['Terampil', 'Mahir', 'Penyelia', 'Ahli Pertama', 'Ahli Muda', 'Ahli Madya', 'Fungsional Lainnya'];
                standardJfOrder.forEach(subName => {
                    const childStat = categories.FUNGSIONAL.children[subName];
                    if (childStat && childStat.total > 0) {
                        rowsHtml += renderRow(subName, childStat, { isChild: true });
                    }
                });
            }

            // 6. Pelaksana
            if (categories.PELAKSANA.stat.total > 0) {
                rowsHtml += renderRow(categories.PELAKSANA.label, categories.PELAKSANA.stat);
            }

            // 7. Belum Terdata / Kosong (Peserta Belum Terdaftar) -> DI BAWAH BARIS PELAKSANA
            if (categories.KOSONG.stat.total > 0) {
                rowsHtml += renderRow(categories.KOSONG.label, categories.KOSONG.stat, { isWarning: true });
            }

            // 8. Kelompok Lainnya (jika ada)
            Object.keys(categories.LAINNYA).forEach(otherLabel => {
                const stat = categories.LAINNYA[otherLabel];
                if (stat.total > 0) {
                    rowsHtml += renderRow(otherLabel, stat);
                }
            });

            // Jika semua kategori standar 0 totalnya (jarang terjadi tapi fallback aman)
            if (!rowsHtml) {
                rowsHtml = `<tr><td colspan="6" class="p-4 text-center text-slate-400 italic">Belum ada rincian jabatan untuk peserta yang dipilih.</td></tr>`;
            }

            kelJabatanContainer.innerHTML = `
                <table class="w-full text-xs text-left text-slate-700">
                    <thead class="bg-slate-100 text-slate-700 font-bold uppercase">
                        <tr>
                            <th class="p-2.5">Kelompok Jabatan</th>
                            <th class="p-2.5 text-center text-emerald-700 font-bold">Hadir</th>
                            <th class="p-2.5 text-center text-rose-700 font-bold">Tidak Hadir</th>
                            <th class="p-2.5 text-center text-amber-700 font-bold">Belum Presensi</th>
                            <th class="p-2.5 text-center font-bold">Total Peserta</th>
                            <th class="p-2.5 text-center font-bold">% Kehadiran</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100">
                        ${rowsHtml}
                    </tbody>
                </table>
            `;
        }
    }

    // Render tabel distribusi berdasarkan unit kerja (mengikuti filter dashboard yang aktif)
    if (distContainer) {
        if (total === 0) {
            distContainer.innerHTML = `<p class="text-sm text-slate-500 py-6 text-center">Belum ada data peserta untuk ujian/tanggal ini. Silakan upload file Excel atau pilih tanggal lain.</p>`;
            if (window.lucide) window.lucide.createIcons();
            return;
        }

        const byUnit = {};
        dashboardCandidates.forEach(c => {
            const rawU = String(c.unitKerja || '').trim();
            const isNullUnit = !rawU || rawU === 'NULL' || rawU === '-' || rawU === '(Unit Kerja Tidak Terisi)';
            const u = isNullUnit ? 'Tidak terdata' : rawU;
            if (!byUnit[u]) byUnit[u] = { isNull: isNullUnit, s1: 0, s2: 0, s3: 0, hadir: 0, tidakHadir: 0, total: 0 };
            const s = Number(c.sesi);
            if (s === 1) byUnit[u].s1++;
            else if (s === 2) byUnit[u].s2++;
            else if (s === 3) byUnit[u].s3++;
            if (c.kehadiran === 'HADIR') byUnit[u].hadir++;
            else if (c.kehadiran === 'TIDAK_HADIR') byUnit[u].tidakHadir++;
            byUnit[u].total++;
        });

        // Urutkan unit kerja: "Tidak terdata" diprioritaskan di atas jika ada, selebihnya berdasarkan total terbanyak
        const sortedUnits = Object.keys(byUnit).sort((a, b) => {
            if (a === 'Tidak terdata') return -1;
            if (b === 'Tidak terdata') return 1;
            return byUnit[b].total - byUnit[a].total;
        });

        distContainer.innerHTML = `
            <table class="w-full text-xs text-left text-slate-700">
                <thead class="bg-slate-100 text-slate-700 font-bold uppercase">
                    <tr>
                        <th class="p-2.5">Unit Kerja / OPD</th>
                        <th class="p-2.5 text-center">Sesi 1</th>
                        <th class="p-2.5 text-center">Sesi 2</th>
                        <th class="p-2.5 text-center">Sesi 3</th>
                        <th class="p-2.5 text-center text-emerald-700 font-bold">Hadir</th>
                        <th class="p-2.5 text-center text-rose-700 font-bold">Tidak Hadir</th>
                        <th class="p-2.5 text-center font-bold">Total Peserta</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-slate-100">
                    ${sortedUnits.slice(0, 15).map(u => {
                        const item = byUnit[u];
                        const isRed = item.isNull || u === 'Tidak terdata';
                        return `
                            <tr class="${isRed ? 'bg-rose-50/90 text-rose-950 font-bold border-l-4 border-l-rose-600 hover:bg-rose-100' : 'hover:bg-slate-50'} transition">
                                <td class="p-2.5 font-medium ${isRed ? 'text-rose-700 font-extrabold flex items-center gap-1.5' : 'text-slate-800'}">
                                    ${isRed ? '<i data-lucide="alert-circle" class="w-3.5 h-3.5 text-rose-600 inline-block"></i>' : ''}
                                    <span>${u}</span>
                                    ${isRed ? '<span class="text-[10px] bg-rose-200 text-rose-800 px-1.5 py-0.2 rounded font-bold ml-1 uppercase">Perlu Update</span>' : ''}
                                </td>
                                <td class="p-2.5 text-center text-blue-700 font-semibold">${item.s1}</td>
                                <td class="p-2.5 text-center text-amber-700 font-semibold">${item.s2}</td>
                                <td class="p-2.5 text-center text-emerald-700 font-semibold">${item.s3}</td>
                                <td class="p-2.5 text-center text-emerald-600 font-bold">${item.hadir}</td>
                                <td class="p-2.5 text-center text-rose-600 font-bold">${item.tidakHadir}</td>
                                <td class="p-2.5 text-center font-bold ${isRed ? 'text-rose-700 font-black' : 'text-slate-900'}">${item.total}</td>
                            </tr>
                        `;
                    }).join('')}
                </tbody>
            </table>
        `;
    }
    if (window.lucide) window.lucide.createIcons();
}

/**
 * Setup Form Create Ujian Baru
 * (Disesuaikan: Tanpa input Nama/Judul Kegiatan, hanya Instansi, Tanggal, Tilok, Kuota, Catatan)
 */
function setupCreateExamForm() {
    const form = document.getElementById('formCreateExam');
    const selectInstansi = document.getElementById('selectExamInstansi');

    if (form) {
        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            const instansi = selectInstansi.value;
            const selectedOpt = selectInstansi.options[selectInstansi.selectedIndex];
            const region = selectedOpt.getAttribute('data-region') || 'Papua Barat';

            const startDate = document.getElementById('inputExamStartDate').value;
            const endDate = document.getElementById('inputExamEndDate').value || startDate;
            const location = document.getElementById('inputExamLocation').value.trim();
            const quota = document.getElementById('inputExamQuota').value;
            const notes = document.getElementById('inputExamNotes').value.trim();

            if (!instansi) {
                showToast("Instansi wajib dipilih!", "warning");
                return;
            }

            try {
                const newExam = await createNewExam({
                    title: instansi,
                    instansi: instansi,
                    wilker: region,
                    startDate,
                    endDate,
                    location,
                    quotaPerSession: quota,
                    notes
                });

                allExams.unshift(newExam);
                if (isCloudActive()) {
                    await saveExamToCloud(newExam);
                }
                renderExamSelectDropdowns();
                saveActiveExamSession(newExam.id, instansi);
                await setActiveExam(newExam.id);

                showToast(`Ujian untuk "${instansi}" berhasil dibuat!`, "success");
                form.reset();

                // Beralih ke tab upload excel
                switchTab('upload-excel');

            } catch (err) {
                console.error(err);
                showToast("Gagal membuat ujian: " + err.message, "error");
            }
        });
    }
}

/**
 * Render daftar ujian yang sudah dibuat pada Tab Create Ujian
 */
function renderExamListInCreateTab() {
    const container = document.getElementById('listExamContainer');
    const countBadge = document.getElementById('examCountBadge');
    if (!container) return;

    if (countBadge) countBadge.textContent = `${allExams.length} Instansi`;

    if (allExams.length === 0) {
        container.innerHTML = `<p class="text-sm text-slate-400 py-8 text-center">Belum ada ujian yang dibuat.</p>`;
        return;
    }

    const activeId = getSelectedExamId();

    container.innerHTML = allExams.map(exam => {
        const isActive = exam.id === activeId;
        return `
            <div class="p-3.5 rounded-xl border ${isActive ? 'border-bkn-600 bg-blue-50/60 ring-2 ring-bkn-600/20' : 'border-slate-200 bg-white hover:border-slate-300'} transition flex flex-col justify-between space-y-2">
                <div>
                    <div class="flex items-center justify-between">
                        <span class="text-[10px] font-bold px-2 py-0.5 rounded-full ${exam.wilker === 'Papua Barat Daya' ? 'bg-emerald-100 text-emerald-800' : (exam.wilker === 'Instansi Vertikal' ? 'bg-amber-100 text-amber-800' : 'bg-blue-100 text-blue-800')}">${exam.wilker}</span>
                        ${isActive ? '<span class="text-[10px] font-bold text-bkn-700 bg-blue-100 px-2 py-0.5 rounded">AKTIF</span>' : ''}
                    </div>
                    <h4 class="font-bold text-sm text-slate-900 mt-1 line-clamp-1">${exam.instansi}</h4>
                    <p class="text-[11px] text-slate-500 mt-0.5 flex items-center gap-1">
                        <i data-lucide="map-pin" class="w-3 h-3 text-slate-400 inline"></i>
                        <span>${exam.location}</span>
                    </p>
                </div>
                <div class="pt-2 border-t border-slate-100 flex items-center justify-between text-xs">
                    <span class="text-[11px] text-slate-400">${exam.startDate ? formatDateDisplay(exam.startDate, 'short') : '-'}</span>
                    <div class="flex items-center space-x-1.5">
                        ${!isActive ? `
                            <button onclick="selectAndActivateExam('${exam.id}')" class="px-2.5 py-1 text-[11px] bg-bkn-700 hover:bg-bkn-800 text-white font-medium rounded transition">
                                Pilih
                            </button>
                        ` : ''}
                        <button onclick="openModalEditExam('${exam.id}')" class="p-1 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded transition" title="Edit Data Ujian">
                            <i data-lucide="edit-3" class="w-3.5 h-3.5"></i>
                        </button>
                        <button onclick="confirmDeleteExam('${exam.id}', '${exam.instansi}')" class="p-1 text-rose-500 hover:text-rose-700 hover:bg-rose-50 rounded transition" title="Hapus Ujian">
                            <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                        </button>
                    </div>
                </div>
            </div>
        `;
    }).join('');

    if (window.lucide) window.lucide.createIcons();
}

window.selectAndActivateExam = async (examId) => {
    const ex = allExams.find(e => e.id === examId);
    if (ex) {
        saveActiveExamSession(ex.id, ex.instansi);
    }
    await setActiveExam(examId);
    showToast("Instansi aktif berhasil dipilih.", "info");
};

window.openModalEditExam = (examId) => {
    const exam = allExams.find(e => e.id === examId);
    if (!exam) return;

    const elId = document.getElementById('editExamId');
    const elInstansi = document.getElementById('editExamInstansi');
    const elStart = document.getElementById('editExamStartDate');
    const elEnd = document.getElementById('editExamEndDate');
    const elLocation = document.getElementById('editExamLocation');
    const elQuota = document.getElementById('editExamQuota');
    const elNotes = document.getElementById('editExamNotes');

    if (elId) elId.value = exam.id;
    if (elInstansi) elInstansi.value = exam.instansi || '';
    if (elStart) elStart.value = exam.startDate || '';
    if (elEnd) elEnd.value = exam.endDate || exam.startDate || '';
    if (elLocation) elLocation.value = exam.location || '';
    if (elQuota) elQuota.value = exam.quotaPerSession || 50;
    if (elNotes) elNotes.value = exam.notes || '';

    const modal = document.getElementById('modalEditExam');
    if (modal) {
        modal.classList.remove('hidden');
        modal.classList.add('flex');
    }
    if (window.lucide) window.lucide.createIcons();
};

window.closeModalEditExam = () => {
    const modal = document.getElementById('modalEditExam');
    if (modal) {
        modal.classList.add('hidden');
        modal.classList.remove('flex');
    }
};

function setupEditExamForm() {
    const form = document.getElementById('formEditExam');
    if (form) {
        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            const examId = document.getElementById('editExamId').value;
            const examIndex = allExams.findIndex(e => e.id === examId);
            if (examIndex === -1) return;

            const startDate = document.getElementById('editExamStartDate').value;
            const endDate = document.getElementById('editExamEndDate').value || startDate;
            const location = document.getElementById('editExamLocation').value.trim();
            const quota = Number(document.getElementById('editExamQuota').value) || 50;
            const notes = document.getElementById('editExamNotes').value.trim();

            const updatedExam = {
                ...allExams[examIndex],
                startDate,
                endDate,
                location,
                quotaPerSession: quota,
                notes
            };

            try {
                await db.updateExam(updatedExam);
                allExams[examIndex] = updatedExam;

                if (isCloudActive()) {
                    await saveExamToCloud(updatedExam);
                }

                if (currentExam && currentExam.id === examId) {
                    currentExam = updatedExam;
                    renderDashboardExamInfo();
                }

                renderExamListInCreateTab();
                window.closeModalEditExam();
                showToast(`Data ujian "${updatedExam.instansi}" berhasil diperbarui!`, "success");
            } catch (err) {
                console.error("Gagal update ujian:", err);
                showToast("Gagal memperbarui ujian: " + err.message, "error");
            }
        });
    }
}

window.confirmDeleteExam = async (examId, instansiName) => {
    if (confirm(`Yakin ingin menghapus ujian untuk "${instansiName}"?\nSeluruh data peserta di dalam ujian ini juga akan terhapus permanen.`)) {
        try {
            await db.deleteExam(examId);
            if (isCloudActive()) {
                await deleteExamFromCloud(examId);
            }
            allExams = allExams.filter(e => e.id !== examId);
            showToast(`Ujian "${instansiName}" berhasil dihapus.`, "success");

            if (getSelectedExamId() === examId) {
                const nextId = allExams.length > 0 ? allExams[0].id : null;
                await setActiveExam(nextId);
            } else {
                renderExamSelectDropdowns();
                renderExamListInCreateTab();
            }
        } catch (err) {
            console.error(err);
            showToast("Gagal menghapus ujian: " + err.message, "error");
        }
    }
};

/**
 * Setup Upload Excel & Drag Drop Area
 */
function setupExcelUpload() {
    const dropzone = document.getElementById('dropzoneExcel');
    const fileInput = document.getElementById('fileInputExcel');

    if (!dropzone || !fileInput) return;

    dropzone.addEventListener('click', () => fileInput.click());

    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('border-bkn-600', 'bg-blue-50/50');
    });

    dropzone.addEventListener('dragleave', () => {
        dropzone.classList.remove('border-bkn-600', 'bg-blue-50/50');
    });

    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('border-bkn-600', 'bg-blue-50/50');
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            handleSelectedExcelFile(e.dataTransfer.files[0]);
        }
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files.length > 0) {
            handleSelectedExcelFile(e.target.files[0]);
        }
    });
}

/**
 * Memproses file Excel yang dipilih (Mendukung ribuan baris data)
 */
// State untuk modal penyelesaian duplikasi
let activeDuplicateState = null;

/**
 * Memproses file Excel yang dipilih (Mendukung ribuan baris data & deteksi duplikasi)
 */
async function handleSelectedExcelFile(file) {
    const targetExamSelect = document.getElementById('selectUploadTargetExam');
    const targetExamId = targetExamSelect ? targetExamSelect.value : getSelectedExamId();

    if (!targetExamId) {
        showToast("Harap pilih target instansi penerima data terlebih dahulu!", "warning");
        return;
    }

    const nameBadge = document.getElementById('uploadFileNameBadge');
    const nameText = document.getElementById('uploadFileNameText');
    if (nameBadge && nameText) {
        nameText.textContent = file.name;
        nameBadge.classList.remove('hidden');
    }

    const autoTime = document.getElementById('checkAutoStandardize')?.checked ?? true;
    const selectedMode = document.querySelector('input[name="uploadModeChoice"]:checked')?.value || 'AUTO';

    try {
        showToast("Sedang memproses & membaca file Excel...", "info");
        const parseResult = await parseExcelFile(file, { 
            autoStandardizeTime: autoTime,
            defaultPelaksanaan: currentExam?.startDate || '',
            uploadMode: selectedMode,
            defaultInstansi: currentExam?.instansi || ''
        });

        if (!parseResult.candidates || parseResult.candidates.length === 0) {
            showToast("Tidak ditemukan baris peserta dengan Nama dan NIP yang valid!", "warning");
            return;
        }

        // Ambil data database yang sudah ada untuk instansi ini
        const existingCandidates = await db.getCandidatesByExam(targetExamId);

        // JIKA INI ADALAH FILE JADWAL (SCHEDULE) DAN DATABASE SUDAH MEMILIKI DATA PESERTA (MISAL DARI FILE SISTEM):
        // Lakukan penggabungan jadwal berdasarkan kecocokan NIP tanpa menimpa nama dan jabatan asli sistem
        if (parseResult.mode === 'SCHEDULE' && existingCandidates && existingCandidates.length > 0) {
            const mergeResult = mergeScheduleWithExisting(parseResult.candidates, existingCandidates);

            previewParsedData = {
                examId: targetExamId,
                candidates: mergeResult.allMerged,
                summary: {
                    totalRows: mergeResult.allMerged.length,
                    skippedRows: parseResult.summary.skippedRows,
                    sesi1: mergeResult.allMerged.filter(c => Number(c.sesi) === 1).length,
                    sesi2: mergeResult.allMerged.filter(c => Number(c.sesi) === 2).length,
                    sesi3: mergeResult.allMerged.filter(c => Number(c.sesi) === 3).length,
                    nullScheduleRows: mergeResult.allMerged.filter(c => !c.sesi || c.sesi === 'NULL').length,
                    fridayRows: mergeResult.allMerged.filter(c => c.isFriday).length
                },
                isMergedUpdate: true,
                mergeInfo: {
                    matchedCount: mergeResult.matchedCount,
                    newCount: mergeResult.newCount
                }
            };

            renderExcelPreview(previewParsedData);
            showToast(`Berhasil mencocokkan jadwal ${mergeResult.matchedCount} peserta. Nama & jabatan asli sistem diproteksi!`, "success");
            return;
        }

        // Untuk mode SYSTEM atau upload normal pertama kali: lakukan analisis duplikasi
        const dupAnalysis = analyzeDuplicates(parseResult.candidates, existingCandidates);

        if (dupAnalysis.hasDuplicates) {
            // Tampilkan pop-up modal konfirmasi duplikasi
            openModalDuplicateResolution(dupAnalysis, targetExamId, parseResult);
            showToast(`Ditemukan ${dupAnalysis.totalDuplicateNips} NIP duplikat. Silakan tentukan data pada pop-up konfirmasi.`, "warning");
            return;
        }

        // Jika tidak ada duplikasi sama sekali, langsung ke preview
        previewParsedData = {
            examId: targetExamId,
            candidates: parseResult.candidates,
            summary: parseResult.summary,
            nipsToReplace: [],
            isMergedUpdate: false
        };

        renderExcelPreview(parseResult);
        showToast(`Berhasil membaca ${parseResult.candidates.length} baris peserta (${parseResult.summary.skippedRows} baris kosong/tidak lengkap dilewati)!`, "success");

    } catch (err) {
        console.error("Gagal membaca Excel:", err);
        showToast("Error membaca Excel: " + err.message, "error");
    }
}

/**
 * Buka Modal Konfirmasi Duplikasi Data Peserta
 */
function openModalDuplicateResolution(dupAnalysis, targetExamId, parseResult) {
    activeDuplicateState = {
        dupAnalysis,
        targetExamId,
        parseResult
    };

    const modal = document.getElementById('modalDuplicateResolution');
    const countTotal = document.getElementById('countModalTotalDupNip');
    const countInternal = document.getElementById('countModalInternalDup');
    const countExisting = document.getElementById('countModalExistingDup');

    if (countTotal) countTotal.textContent = `${dupAnalysis.totalDuplicateNips} NIP`;
    if (countInternal) countInternal.textContent = `${dupAnalysis.internalDuplicates.length} NIP`;
    if (countExisting) countExisting.textContent = `${dupAnalysis.existingDuplicates.length} NIP`;

    renderDuplicateGroupsInModal(dupAnalysis);
    updateDuplicateSelectedCount();

    if (modal) {
        modal.classList.remove('hidden');
        modal.classList.add('flex');
    }

    if (window.lucide) window.lucide.createIcons();
}

/**
 * Otomatis pilih baris pertama untuk setiap NIP
 */
window.autoSelectFirstExcelDuplicates = () => {
    const seenNips = new Set();
    document.querySelectorAll('.dup-checkbox').forEach(cb => {
        const nip = cb.getAttribute('data-nip');
        const type = cb.getAttribute('data-type');

        if (type === 'internal') {
            if (!seenNips.has(nip)) {
                cb.checked = true;
                seenNips.add(nip);
            } else {
                cb.checked = false;
            }
        } else if (type === 'keep-db') {
            // Default keep data DB
            cb.checked = true;
        } else if (type === 'existing') {
            // Incoming default false
            cb.checked = false;
        }
    });
    updateDuplicateSelectedCount();
    showToast("Dipilih: Hanya baris pertama tiap NIP.", "info");
};

/**
 * Ganti Data Database dengan File Excel Baru
 */
window.autoSelectLatestExcelDuplicates = () => {
    const seenNips = new Set();
    document.querySelectorAll('.dup-checkbox').forEach(cb => {
        const nip = cb.getAttribute('data-nip');
        const type = cb.getAttribute('data-type');
        
        if (type === 'keep-db') {
            cb.checked = false;
        } else if (type === 'existing' || type === 'internal') {
            if (!seenNips.has(nip)) {
                cb.checked = true;
                seenNips.add(nip);
            } else {
                cb.checked = false;
            }
        }
    });
    updateDuplicateSelectedCount();
    showToast("Dipilih: Versi terbaru dari file Excel untuk menggantikan data lama.", "info");
};

/**
 * Tutup Modal Duplikasi
 */
window.closeModalDuplicateResolution = () => {
    const modal = document.getElementById('modalDuplicateResolution');
    if (modal) {
        modal.classList.add('hidden');
        modal.classList.remove('flex');
    }
    activeDuplicateState = null;
};

/**
 * Terapkan Hasil Pilihan Duplikasi & Lanjutkan ke Preview
 */
window.applyDuplicateResolutionAndProceed = () => {
    if (!activeDuplicateState) return;

    const { dupAnalysis, targetExamId, parseResult } = activeDuplicateState;

    // Kumpulkan key yang dicentang
    const checkedKeys = new Set();
    const checkedExistingNipsToReplace = [];

    document.querySelectorAll('.dup-checkbox:checked').forEach(cb => {
        const key = cb.getAttribute('data-key');
        const nip = cb.getAttribute('data-nip');
        const type = cb.getAttribute('data-type');

        checkedKeys.add(key);

        if (type === 'existing') {
            checkedExistingNipsToReplace.push(nip);
        }
    });

    // Kumpulkan kandidat terpilih dari internal duplicates
    const selectedDuplicates = [];
    dupAnalysis.internalDuplicates.forEach(group => {
        group.items.forEach(item => {
            if (checkedKeys.has(item.uniqueKey)) {
                selectedDuplicates.push(item);
            }
        });
    });

    // Kumpulkan kandidat terpilih dari existing duplicates (incoming items)
    dupAnalysis.existingDuplicates.forEach(group => {
        group.incomingItems.forEach(item => {
            if (checkedKeys.has(item.uniqueKey)) {
                selectedDuplicates.push(item);
            }
        });
    });

    // Gabungkan data bersih (tanpa duplikat) dengan data duplikat yang telah dipilih
    const finalCandidates = [...dupAnalysis.cleanCandidates, ...selectedDuplicates];

    if (finalCandidates.length === 0) {
        showToast("Tidak ada peserta yang dipilih untuk di-upload!", "warning");
        return;
    }

    // Urutkan kembali nomor urut
    finalCandidates.forEach((c, idx) => {
        c.no = idx + 1;
    });

    previewParsedData = {
        examId: targetExamId,
        candidates: finalCandidates,
        nipsToReplace: checkedExistingNipsToReplace,
        isMergedUpdate: false,
        summary: {
            totalRows: finalCandidates.length,
            skippedRows: parseResult.summary.skippedRows,
            sesi1: finalCandidates.filter(c => Number(c.sesi) === 1).length,
            sesi2: finalCandidates.filter(c => Number(c.sesi) === 2).length,
            sesi3: finalCandidates.filter(c => Number(c.sesi) === 3).length,
            nullScheduleRows: finalCandidates.filter(c => !c.sesi || c.sesi === 'NULL' || c.sesi === '00' || c.sesi === 0 || c.sesi === '0').length,
            fridayRows: finalCandidates.filter(c => c.isFriday).length
        }
    };

    closeModalDuplicateResolution();
    renderExcelPreview(previewParsedData);
    showToast(`Pilihan duplikasi diterapkan! Total ${finalCandidates.length} peserta siap disimpan ke database.`, "success");
};

/**
 * Render Tabel Preview dan Ringkasan Sesi sebelum Simpan
 */
function renderExcelPreview(result) {
    const previewCard = document.getElementById('previewCard');
    const summaryContainer = document.getElementById('previewSummaryContainer');
    const tbody = document.getElementById('tbodyExcelPreview');

    if (!previewCard || !tbody) return;

    previewCard.classList.remove('hidden');

    const s = result.summary;
    if (summaryContainer) {
        summaryContainer.innerHTML = `
            <div class="bg-blue-50 p-3 rounded-lg border border-blue-200">
                <div class="text-[10px] uppercase font-bold text-blue-700">Total Terbaca</div>
                <div class="text-xl font-bold text-blue-900 mt-0.5">${s.totalRows} Peserta</div>
            </div>
            ${s.nullScheduleRows && s.nullScheduleRows > 0 ? `
                <div class="bg-slate-100 p-3 rounded-lg border border-slate-200">
                    <div class="text-[10px] uppercase font-bold text-slate-600">Belum Terjadwal (NULL)</div>
                    <div class="text-xl font-bold text-slate-800 mt-0.5">${s.nullScheduleRows} Peserta</div>
                </div>
            ` : `
                <div class="bg-indigo-50 p-3 rounded-lg border border-indigo-200">
                    <div class="text-[10px] uppercase font-bold text-indigo-700">Sesi 1 (08.00-11.00)</div>
                    <div class="text-xl font-bold text-indigo-900 mt-0.5">${s.sesi1} Orang</div>
                </div>
            `}
            <div class="bg-amber-50 p-3 rounded-lg border border-amber-200 relative">
                ${s.fridayRows > 0 ? '<span class="absolute top-2 right-2 text-[9px] bg-amber-200 text-amber-900 font-bold px-1.5 py-0.2 rounded">JUMAT DETECTED</span>' : ''}
                <div class="text-[10px] uppercase font-bold text-amber-700">Sesi 2 (11.00 / Jumat 13.00)</div>
                <div class="text-xl font-bold text-amber-900 mt-0.5">${s.sesi2} Orang</div>
            </div>
            <div class="bg-emerald-50 p-3 rounded-lg border border-emerald-200">
                <div class="text-[10px] uppercase font-bold text-emerald-700">Sesi 3 (14.00-17.00)</div>
                <div class="text-xl font-bold text-emerald-900 mt-0.5">${s.sesi3} Orang</div>
            </div>
        `;
    }

    const previewList = result.candidates.slice(0, 25);
    tbody.innerHTML = previewList.map((c, idx) => {
        const isNullSchedule = !c.pelaksanaan || c.pelaksanaan === 'NULL' || !c.sesi || c.sesi === 'NULL' || c.sesi === '00' || c.sesi === 0;
        const isFri = !isNullSchedule && c.isFriday && Number(c.sesi) === 2;
        const isNullUnit = !c.unitKerja || c.unitKerja === 'NULL' || c.unitKerja === '-';
        return `
            <tr class="${isFri ? 'bg-amber-50/60 font-medium' : 'hover:bg-slate-50'}">
                <td class="p-2.5 text-center text-slate-500">${c.no || (idx + 1)}</td>
                <td class="p-2.5 font-mono text-slate-900">${c.nip}</td>
                <td class="p-2.5 font-semibold text-slate-900">${c.nama}</td>
                <td class="p-2.5 font-medium text-blue-700">${c.kelJabatan || '-'}</td>
                <td class="p-2.5 text-slate-600">
                    ${isNullUnit ? '<span class="px-2 py-0.5 rounded text-[10px] font-bold bg-slate-100 text-slate-500 border border-slate-200">NULL</span>' : c.unitKerja}
                </td>
                <td class="p-2.5 text-slate-600">${c.jabatan || '-'}</td>
                <td class="p-2.5 whitespace-nowrap">
                    ${isNullSchedule ? `
                        <span class="px-2 py-0.5 rounded text-[10px] font-bold bg-slate-100 text-slate-500 border border-slate-200">NULL</span>
                    ` : `
                        <span class="font-medium ${c.isFriday ? 'text-amber-800' : 'text-slate-800'}">${c.pelaksanaan}</span>
                        ${c.isFriday ? '<span class="text-[9px] bg-amber-100 text-amber-800 font-bold px-1 rounded ml-1">Jumat</span>' : ''}
                    `}
                </td>
                <td class="p-2.5 text-center whitespace-nowrap min-w-[95px]">
                    ${isNullSchedule ? `
                        <span class="inline-block whitespace-nowrap px-2 py-0.5 rounded text-[11px] font-bold bg-slate-100 text-slate-700 border border-slate-300">
                            Sesi 00
                        </span>
                    ` : `
                        <span class="inline-block whitespace-nowrap px-2.5 py-0.5 rounded text-[11px] font-bold ${Number(c.sesi) === 1 ? 'bg-blue-100 text-blue-800' : (Number(c.sesi) === 2 ? 'bg-amber-100 text-amber-800' : 'bg-emerald-100 text-emerald-800')}">
                            Sesi ${c.sesi}
                        </span>
                    `}
                </td>
                <td class="p-2.5 whitespace-nowrap font-medium ${isFri ? 'text-amber-800 font-bold' : 'text-slate-700'}">
                    ${isNullSchedule ? '<span class="text-slate-400 font-bold">NULL</span>' : c.waktu}
                </td>
            </tr>
        `;
    }).join('');

    if (result.candidates.length > 25) {
        tbody.innerHTML += `
            <tr>
                <td colspan="9" class="p-3 text-center text-xs text-slate-500 bg-slate-50 font-medium italic">
                    ... dan ${result.candidates.length - 25} peserta lainnya akan dimasukkan ke database saat disimpan.
                </td>
            </tr>
        `;
    }
}

/**
 * Batalkan preview upload
 */
window.cancelUploadPreview = () => {
    previewParsedData = null;
    activeDuplicateState = null;
    const previewCard = document.getElementById('previewCard');
    const fileInput = document.getElementById('fileInputExcel');
    const nameBadge = document.getElementById('uploadFileNameBadge');

    if (previewCard) previewCard.classList.add('hidden');
    if (fileInput) fileInput.value = '';
    if (nameBadge) nameBadge.classList.add('hidden');
};

/**
 * Simpan Data Hasil Parsing Excel ke Database (IndexedDB & Firebase Cloud)
 */
window.savePreviewDataToDatabase = async () => {
    if (!previewParsedData || !previewParsedData.candidates || previewParsedData.candidates.length === 0) {
        showToast("Tidak ada data untuk disimpan!", "warning");
        return;
    }

    const btn = document.getElementById('btnSaveExcelToDB');
    if (btn) {
        btn.disabled = true;
        btn.innerHTML = `<i data-lucide="loader-2" class="w-4 h-4 animate-spin inline mr-1"></i> Menyimpan ke DB...`;
    }

    try {
        if (previewParsedData.isMergedUpdate) {
            // Update gabungan jadwal: ganti seluruh dataset instansi dengan dataset yang sudah di-merge
            await db.deleteCandidatesByExam(previewParsedData.examId);
            const count = await db.bulkAddCandidates(previewParsedData.examId, previewParsedData.candidates);
            if (isCloudActive()) {
                await bulkAddCandidatesToCloud(previewParsedData.examId, previewParsedData.candidates);
            }
            showToast(`Sukses! Jadwal ${count} peserta berhasil diperbarui (Nama & Jabatan sistem tetap terlindungi).`, "success");
        } else {
            // Hapus data lama yang digantikan jika ada dari resolusi duplikasi
            if (previewParsedData.nipsToReplace && previewParsedData.nipsToReplace.length > 0) {
                await db.deleteCandidatesByNips(previewParsedData.examId, previewParsedData.nipsToReplace);
            }

            const count = await db.bulkAddCandidates(previewParsedData.examId, previewParsedData.candidates);
            if (isCloudActive()) {
                await bulkAddCandidatesToCloud(previewParsedData.examId, previewParsedData.candidates);
            }
            showToast(`Sukses! ${count} peserta berhasil disimpan ke dalam database.`, "success");
        }

        await setActiveExam(previewParsedData.examId);
        window.cancelUploadPreview();

        // Beralih otomatis ke tab data peserta
        switchTab('daftar-peserta');

    } catch (err) {
        console.error("Gagal simpan ke DB:", err);
        showToast("Gagal menyimpan ke database: " + err.message, "error");
    } finally {
        if (btn) {
            btn.disabled = false;
            btn.innerHTML = `<i data-lucide="database" class="w-4 h-4 inline mr-1"></i> Simpan ke Database`;
            if (window.lucide) window.lucide.createIcons();
        }
    }
};

/**
 * Download Template Excel Master Sistem
 */
window.triggerDownloadSystemTemplate = () => {
    downloadSystemTemplate();
    showToast("Template Data Master Sistem berhasil diunduh!", "info");
};

/**
 * Download Template Excel Jadwal Ujian
 */
window.triggerDownloadScheduleTemplate = () => {
    downloadScheduleTemplate();
    showToast("Template Jadwal Ujian berhasil diunduh!", "info");
};

/**
 * Download Template Excel (Backwards compatibility)
 */
window.triggerDownloadTemplate = () => {
    downloadScheduleTemplate();
    showToast("Template Excel berhasil diunduh!", "info");
};

/**
 * Filter dan Render Tabel Data Peserta
 */
window.setCandidateFilterSession = (session) => {
    currentSessionFilter = session;

    // Update status aktif tombol pills
    document.querySelectorAll('.filter-sesi-btn').forEach(btn => {
        btn.className = 'filter-sesi-btn px-3 py-1.5 text-xs font-semibold rounded-lg bg-slate-100 text-slate-700 hover:bg-slate-200 transition';
    });

    const activeBtn = session === 'ALL' 
        ? document.getElementById('btnFilterSesiAll') 
        : document.getElementById(`btnFilterSesi${session}`);

    if (activeBtn) {
        activeBtn.className = 'filter-sesi-btn px-3 py-1.5 text-xs font-semibold rounded-lg bg-bkn-800 text-white shadow-sm transition';
    }

    const inputTyping = document.getElementById('inputFilterSesiTyping');
    const selectDropdown = document.getElementById('selectFilterSesiDropdown');

    if (session === 'ALL') {
        currentCumulativeSessionFilter = 'ALL';
        if (inputTyping) inputTyping.value = '';
        if (selectDropdown) selectDropdown.value = 'ALL';
    } else if (session === '00' || session === 0 || session === '0') {
        currentCumulativeSessionFilter = '00';
        if (inputTyping) inputTyping.value = '00';
        if (selectDropdown) {
            const opt00 = selectDropdown.querySelector('option[value="00"]');
            if (opt00) selectDropdown.value = '00';
        }
        // Pastikan filter tanggal di-reset ke ALL agar semua peserta sesi 00 langsung terlihat
        const selectDate = document.getElementById('selectFilterPelaksanaan');
        if (selectDate && currentDateFilter !== 'ALL') {
            currentDateFilter = 'ALL';
            selectDate.value = 'ALL';
        }
    } else {
        currentCumulativeSessionFilter = 'ALL';
        if (inputTyping) inputTyping.value = '';
        if (selectDropdown) selectDropdown.value = 'ALL';
    }

    applyCandidateFilters();
};

/**
 * Event handler saat dropdown filter tanggal pelaksanaan berubah
 * Mengatur filter tanggal sekaligus menyaring opsi filter sesi agar hanya sesi pada tanggal tersebut yang aktif
 */
window.onFilterPelaksanaanChange = (dateVal) => {
    currentDateFilter = dateVal || 'ALL';
    populateSesiFilterDropdown(currentDateFilter);
    applyCandidateFilters();
};

/**
 * Event handler saat user mengetik angka sesi pada input (misal ketik 7, 07, atau 00)
 */
let sessionTypeDebounce = null;
window.onFilterSesiTypeInput = (inputVal) => {
    clearTimeout(sessionTypeDebounce);
    sessionTypeDebounce = setTimeout(() => {
        const select = document.getElementById('selectFilterSesiDropdown');
        const clean = String(inputVal || '').trim();

        if (!clean) {
            currentCumulativeSessionFilter = 'ALL';
            currentSessionFilter = 'ALL';
            if (select) select.value = 'ALL';
            document.querySelectorAll('.filter-sesi-btn').forEach(btn => {
                btn.className = 'filter-sesi-btn px-3 py-1.5 text-xs font-semibold rounded-lg bg-slate-100 text-slate-700 hover:bg-slate-200 transition';
            });
            const btnAll = document.getElementById('btnFilterSesiAll');
            if (btnAll) btnAll.className = 'filter-sesi-btn px-3 py-1.5 text-xs font-semibold rounded-lg bg-bkn-800 text-white shadow-sm transition';
            applyCandidateFilters();
            return;
        }

        if (clean === '0' || clean === '00') {
            setCandidateFilterSession('00');
            return;
        }

        const num = parseInt(clean, 10);
        if (!isNaN(num) && num > 0) {
            currentCumulativeSessionFilter = num;
            currentSessionFilter = 'ALL';
            document.querySelectorAll('.filter-sesi-btn').forEach(btn => {
                btn.className = 'filter-sesi-btn px-3 py-1.5 text-xs font-semibold rounded-lg bg-slate-100 text-slate-700 hover:bg-slate-200 transition';
            });
            if (select) {
                const opt = select.querySelector(`option[value="${num}"]`);
                if (opt) {
                    select.value = String(num);
                } else {
                    select.value = 'ALL';
                }
            }
            applyCandidateFilters();
        }
    }, 150);
};

/**
 * Event handler saat dropdown filter sesi dipilih
 */
window.onFilterSesiDropdownChange = (sessionValue) => {
    const inputTyping = document.getElementById('inputFilterSesiTyping');
    if (sessionValue === 'ALL') {
        currentCumulativeSessionFilter = 'ALL';
        currentSessionFilter = 'ALL';
        if (inputTyping) inputTyping.value = '';
        document.querySelectorAll('.filter-sesi-btn').forEach(btn => {
            btn.className = 'filter-sesi-btn px-3 py-1.5 text-xs font-semibold rounded-lg bg-slate-100 text-slate-700 hover:bg-slate-200 transition';
        });
        const btnAll = document.getElementById('btnFilterSesiAll');
        if (btnAll) btnAll.className = 'filter-sesi-btn px-3 py-1.5 text-xs font-semibold rounded-lg bg-bkn-800 text-white shadow-sm transition';
    } else if (sessionValue === '00' || sessionValue === 0 || sessionValue === '0') {
        setCandidateFilterSession('00');
        return;
    } else {
        currentCumulativeSessionFilter = Number(sessionValue);
        currentSessionFilter = 'ALL';
        if (inputTyping) inputTyping.value = formatCumulativeSessionNumber(sessionValue);
        document.querySelectorAll('.filter-sesi-btn').forEach(btn => {
            btn.className = 'filter-sesi-btn px-3 py-1.5 text-xs font-semibold rounded-lg bg-slate-100 text-slate-700 hover:bg-slate-200 transition';
        });
    }
    applyCandidateFilters();
};

let debounceTimer = null;
window.debounceSearchCandidate = () => {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => {
        const input = document.getElementById('inputSearchCandidate');
        currentSearchTerm = input ? input.value.trim().toLowerCase() : '';
        applyCandidateFilters();
    }, 200);
};

window.renderCandidateTable = () => {
    const selectDate = document.getElementById('selectFilterPelaksanaan');
    window.onFilterPelaksanaanChange(selectDate ? selectDate.value : 'ALL');
};

/**
 * Reset Seluruh Filter dan Input Pencarian ke Nilai Default
 */
window.resetAllCandidateFilters = () => {
    // 1. Reset input search
    const inputSearch = document.getElementById('inputSearchCandidate');
    if (inputSearch) inputSearch.value = '';
    currentSearchTerm = '';

    // 2. Reset Kelompok Jabatan
    const selectKel = document.getElementById('selectFilterKelJabatan');
    if (selectKel) selectKel.value = 'ALL';
    currentKelJabatanFilter = 'ALL';

    // 3. Reset Tanggal Pelaksanaan
    const selectDate = document.getElementById('selectFilterPelaksanaan');
    if (selectDate) selectDate.value = 'ALL';
    currentDateFilter = 'ALL';

    // 4. Reset Sesi Pills & Sesi Kumulatif
    currentSessionFilter = 'ALL';
    currentCumulativeSessionFilter = 'ALL';

    const inputTyping = document.getElementById('inputFilterSesiTyping');
    if (inputTyping) inputTyping.value = '';

    const selectSesi = document.getElementById('selectFilterSesiDropdown');
    if (selectSesi) selectSesi.value = 'ALL';

    document.querySelectorAll('.filter-sesi-btn').forEach(btn => {
        btn.className = 'filter-sesi-btn px-3 py-1.5 text-xs font-semibold rounded-lg bg-slate-100 text-slate-700 hover:bg-slate-200 transition';
    });
    const btnAll = document.getElementById('btnFilterSesiAll');
    if (btnAll) btnAll.className = 'filter-sesi-btn px-3 py-1.5 text-xs font-semibold rounded-lg bg-bkn-800 text-white shadow-sm transition';

    populateSesiFilterDropdown('ALL');
    populateKelJabatanFilterDropdown();
    applyCandidateFilters();

    showToast("Semua filter dan pencarian telah di-reset.", "info");
};

function populatePelaksanaanFilterDropdown() {
    const select = document.getElementById('selectFilterPelaksanaan');
    if (!select) return;

    const uniqueDates = getSortedExamDates();
    
    let html = `<option value="ALL">-- Tanggal Pelaksanaan (${uniqueDates.length} Tanggal) --</option>`;
    uniqueDates.forEach(d => {
        const fri = isFriday(d);
        html += `<option value="${d}">${d} ${fri ? '(Hari Jumat - Sesi 2: 13.00)' : ''}</option>`;
    });

    select.innerHTML = html;
}

/**
 * Mengisi dropdown filter sesi kumulatif (1..36)
 * Jika selectedDate !== 'ALL', otomatis hanya menampilkan sesi pada tanggal tersebut
 */
function populateSesiFilterDropdown(selectedDate = 'ALL') {
    const select = document.getElementById('selectFilterSesiDropdown');
    const inputTyping = document.getElementById('inputFilterSesiTyping');
    if (!select) return;

    const sortedDates = getSortedExamDates();

    // Saring kandidat sesuai tanggal jika dipilih
    const poolCandidates = (selectedDate && selectedDate !== 'ALL')
        ? currentCandidates.filter(c => c.pelaksanaan === selectedDate)
        : currentCandidates;

    // Kumpulkan peserta yang belum terjadwal (Sesi 00 / NULL)
    const countSesi00 = poolCandidates.filter(c => !c.sesi || c.sesi === 'NULL' || c.sesi === '00' || c.sesi === 0 || c.sesi === '0').length;

    // Kumpulkan seluruh sesi kumulatif unik pada pool ini
    const sessionMap = new Map();
    poolCandidates.forEach(c => {
        const cum = getCumulativeSessionNumber(c, sortedDates);
        if (cum !== null) {
            if (!sessionMap.has(cum)) {
                sessionMap.set(cum, {
                    cumNum: cum,
                    formattedCum: formatCumulativeSessionNumber(cum),
                    dateStr: c.pelaksanaan,
                    dailySession: c.sesi,
                    count: 0
                });
            }
            sessionMap.get(cum).count++;
        }
    });

    const sortedSessions = Array.from(sessionMap.values()).sort((a, b) => a.cumNum - b.cumNum);

    let defaultText = selectedDate === 'ALL'
        ? `-- Sesi (${sortedSessions.length > 0 ? `01 s.d. ${formatCumulativeSessionNumber(sortedSessions[sortedSessions.length - 1].cumNum)}` : '0 Sesi'}) --`
        : `-- Sesi di Tanggal Ini (${sortedSessions.length} Sesi) --`;

    let html = `<option value="ALL">${defaultText}</option>`;

    if (countSesi00 > 0) {
        html += `<option value="00">Sesi 00 (Belum Terjadwal / NULL) [${countSesi00} Peserta]</option>`;
    }

    sortedSessions.forEach(s => {
        const dateLabel = selectedDate === 'ALL' ? `${s.dateStr} - ` : '';
        html += `<option value="${s.cumNum}">Sesi ${s.formattedCum} (${dateLabel}Sesi ${s.dailySession}) [${s.count} Peserta]</option>`;
    });

    select.innerHTML = html;

    // Update placeholder input ketik sesi
    if (inputTyping) {
        if (sortedSessions.length > 0) {
            const minCum = sortedSessions[0].formattedCum;
            const maxCum = sortedSessions[sortedSessions.length - 1].formattedCum;
            inputTyping.placeholder = `${minCum}-${maxCum}`;
            inputTyping.title = `Ketik angka sesi (misal: 00 untuk NULL, atau ${minCum} s.d. ${maxCum})`;
        } else {
            inputTyping.placeholder = 'Sesi #';
        }
    }

    // Validasi apakah filter sesi terpilih masih ada di dalam daftar sesi yang aktif
    if (currentCumulativeSessionFilter === '00' || currentSessionFilter === '00') {
        select.value = '00';
        if (inputTyping) inputTyping.value = '00';
    } else if (currentCumulativeSessionFilter !== 'ALL') {
        const exists = sortedSessions.some(s => s.cumNum === Number(currentCumulativeSessionFilter));
        if (exists) {
            select.value = String(currentCumulativeSessionFilter);
            if (inputTyping) inputTyping.value = formatCumulativeSessionNumber(currentCumulativeSessionFilter);
        } else {
            currentCumulativeSessionFilter = 'ALL';
            select.value = 'ALL';
            if (inputTyping) inputTyping.value = '';
        }
    } else {
        select.value = 'ALL';
        if (inputTyping && !inputTyping.value) inputTyping.value = '';
    }
}

/**
 * Mengisi dropdown filter kelompok jabatan beserta counter jumlah peserta
 */
function populateKelJabatanFilterDropdown() {
    const select = document.getElementById('selectFilterKelJabatan');
    if (!select) return;

    const kelMap = new Map();
    let countEmpty = 0;

    currentCandidates.forEach(c => {
        const raw = String(c.kelJabatan || '').trim();
        if (!raw || raw === '-' || raw === 'NULL') {
            countEmpty++;
        } else {
            kelMap.set(raw, (kelMap.get(raw) || 0) + 1);
        }
    });

    const sortedKels = Array.from(kelMap.keys()).sort((a, b) => a.localeCompare(b, 'id', { sensitivity: 'base' }));

    let html = `<option value="ALL">-- Kel. Jabatan (${currentCandidates.length}) --</option>`;

    if (countEmpty > 0) {
        html += `<option value="EMPTY">⚠️ [Kosong / Peserta Belum Terdaftar] (${countEmpty} Peserta)</option>`;
    }

    sortedKels.forEach(k => {
        html += `<option value="${k}">${k} (${kelMap.get(k)} Peserta)</option>`;
    });

    select.innerHTML = html;

    if (currentKelJabatanFilter === 'EMPTY') {
        select.value = countEmpty > 0 ? 'EMPTY' : 'ALL';
        if (select.value === 'ALL') currentKelJabatanFilter = 'ALL';
    } else if (currentKelJabatanFilter !== 'ALL') {
        if (kelMap.has(currentKelJabatanFilter)) {
            select.value = currentKelJabatanFilter;
        } else {
            currentKelJabatanFilter = 'ALL';
            select.value = 'ALL';
        }
    } else {
        select.value = 'ALL';
    }
}

window.onFilterKelJabatanChange = (val) => {
    currentKelJabatanFilter = val || 'ALL';
    applyCandidateFilters();
};
 
/**
 * Konversi tanggal fleksibel ke format ISO "YYYY-MM-DD" untuk <input type="date">
 */
function dateToISOInput(dateInput) {
    if (!dateInput || dateInput === 'NULL' || dateInput === '-') return '';
    const d = parseFlexibleDate(dateInput);
    if (!d) return '';
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
}

/**
 * Mendapatkan daftar tanggal pelaksanaan unik yang terurut secara kronologis
 */
function getSortedExamDates() {
    const dateStrings = Array.from(new Set(currentCandidates.map(c => c.pelaksanaan).filter(p => p && p !== 'NULL' && p !== '-')));
    return dateStrings.sort((a, b) => {
        const da = parseFlexibleDate(a);
        const db = parseFlexibleDate(b);
        const ta = da ? da.getTime() : 0;
        const tb = db ? db.getTime() : 0;
        return ta - tb;
    });
}

/**
 * Menghitung nomor sesi kumulatif berlanjut antar hari
 * Hari 1: Sesi 1 -> 1, Sesi 2 -> 2, Sesi 3 -> 3
 * Hari 2: Sesi 1 -> 4, Sesi 2 -> 5, Sesi 3 -> 6, dst.
 */
function getCumulativeSessionNumber(c, sortedDates) {
    if (!c || !c.pelaksanaan || c.pelaksanaan === 'NULL' || !c.sesi || c.sesi === 'NULL') return null;
    const dayIndex = sortedDates.indexOf(c.pelaksanaan);
    if (dayIndex === -1) return null;
    const s = Number(c.sesi);
    if (isNaN(s) || s < 1) return null;
    return (dayIndex * 3) + s;
}

/**
 * Toggle Status Kehadiran Peserta (HADIR, TIDAK_HADIR, RESET)
 */
window.toggleAttendance = async (candidateNipOrId, action) => {
    const lookup = String(candidateNipOrId || '').trim();
    const cand = currentCandidates.find(c => String(c.nip || '').trim() === lookup || String(c.id || '').trim() === lookup);
    if (!cand) {
        console.warn("Peserta tidak ditemukan untuk presensi:", candidateNipOrId);
        return;
    }

    if (action === 'RESET') {
        cand.kehadiran = null;
    } else {
        cand.kehadiran = action; // 'HADIR' atau 'TIDAK_HADIR'
    }

    try {
        await db.updateCandidate(cand);

        // Sinkronkan ke Firebase Cloud secara Realtime dengan NIP string asli
        if (isCloudActive() && currentExam) {
            updateAttendanceInCloud(currentExam.id, String(cand.nip || cand.id).trim(), cand.kehadiran);
        }

        applyCandidateFilters();
        const statusText = cand.kehadiran === 'HADIR' ? 'Hadir' : (cand.kehadiran === 'TIDAK_HADIR' ? 'Tidak Hadir' : 'Direset');
        showToast(`Status ${cand.nama}: ${statusText}`, cand.kehadiran === 'HADIR' ? 'success' : (cand.kehadiran === 'TIDAK_HADIR' ? 'error' : 'info'));
    } catch (err) {
        console.error("Gagal update status kehadiran:", err);
        showToast("Gagal menyimpan status kehadiran: " + err.message, "error");
    }
};

/**
 * Mengurutkan tabel berdasarkan header yang diklik
 */
window.sortTable = (colKey) => {
    if (currentSortColumn === colKey) {
        currentSortDirection = currentSortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        currentSortColumn = colKey;
        currentSortDirection = 'asc';
    }
    updateSortIcons();
    applyCandidateFilters();
};

/**
 * Memperbarui ikon panah sorting pada header tabel
 */
function updateSortIcons() {
    const columns = ['no', 'kehadiran', 'nip', 'nama', 'kelJabatan', 'unitKerja', 'jabatan', 'pelaksanaan', 'sesi', 'waktu'];
    columns.forEach(col => {
        const iconEl = document.getElementById(`sort-icon-${col}`);
        if (!iconEl) return;
        if (col === currentSortColumn) {
            iconEl.textContent = currentSortDirection === 'asc' ? '▲' : '▼';
            iconEl.className = 'text-bkn-800 font-bold text-[11px]';
        } else {
            iconEl.textContent = '↕';
            iconEl.className = 'text-slate-400 text-[10px]';
        }
    });
}

function applyCandidateFilters() {
    const sortedDates = getSortedExamDates();

    filteredCandidates = currentCandidates.filter(c => {
        const isSesi00 = !c.sesi || c.sesi === 'NULL' || c.sesi === '00' || c.sesi === 0 || c.sesi === '0';

        // Filter Sesi Kumulatif & Harian (Prioritas Tinggi)
        if (currentCumulativeSessionFilter === '00' || currentSessionFilter === '00') {
            if (!isSesi00) return false;
        } else {
            // Filter Tanggal (jika bukan filter sesi 00)
            if (currentDateFilter !== 'ALL' && c.pelaksanaan !== currentDateFilter) {
                return false;
            }

            // Filter Sesi Kumulatif (1..36)
            if (currentCumulativeSessionFilter !== 'ALL') {
                if (isSesi00) return false;
                const cum = getCumulativeSessionNumber(c, sortedDates);
                if (cum !== Number(currentCumulativeSessionFilter)) {
                    return false;
                }
            } else if (currentSessionFilter !== 'ALL') {
                // Filter Sesi Harian (1, 2, 3) jika sesi kumulatif ALL
                if (isSesi00) return false;
                if (Number(c.sesi) !== Number(currentSessionFilter)) {
                    return false;
                }
            }
        }

        // Filter Kelompok Jabatan
        if (currentKelJabatanFilter === 'EMPTY') {
            const rawK = String(c.kelJabatan || '').trim();
            const isKelEmpty = !rawK || rawK === '-' || rawK === 'NULL';
            if (!isKelEmpty) return false;
        } else if (currentKelJabatanFilter !== 'ALL') {
            if (String(c.kelJabatan || '').trim() !== currentKelJabatanFilter) {
                return false;
            }
        }

        // Filter Search (NIP, Nama, Kel Jabatan, Unit Kerja, Jabatan)
        if (currentSearchTerm) {
            const matchNip = String(c.nip || '').toLowerCase().includes(currentSearchTerm);
            const matchNama = String(c.nama || '').toLowerCase().includes(currentSearchTerm);
            const matchKel = String(c.kelJabatan || '').toLowerCase().includes(currentSearchTerm);
            const matchUnit = String(c.unitKerja || '').toLowerCase().includes(currentSearchTerm);
            const matchJabatan = String(c.jabatan || '').toLowerCase().includes(currentSearchTerm);
            if (!matchNip && !matchNama && !matchKel && !matchUnit && !matchJabatan) return false;
        }

        return true;
    });

    // Pengurutan data (Sorting) - Default Sesi ASC lalu Nama ASC
    filteredCandidates.sort((a, b) => {
        let valA, valB;

        const isSesi00A = !a.sesi || a.sesi === 'NULL' || a.sesi === '00' || a.sesi === 0 || a.sesi === '0';
        const isSesi00B = !b.sesi || b.sesi === 'NULL' || b.sesi === '00' || b.sesi === 0 || b.sesi === '0';

        switch (currentSortColumn) {
            case 'no':
                valA = Number(a.no) || 0;
                valB = Number(b.no) || 0;
                break;
            case 'kehadiran': {
                const getAttWeight = (k) => k === 'HADIR' ? 1 : (k === 'TIDAK_HADIR' ? 2 : 3);
                valA = getAttWeight(a.kehadiran);
                valB = getAttWeight(b.kehadiran);
                break;
            }
            case 'nip':
                valA = String(a.nip || '');
                valB = String(b.nip || '');
                break;
            case 'nama':
                valA = String(a.nama || '');
                valB = String(b.nama || '');
                break;
            case 'kelJabatan':
                valA = String(a.kelJabatan || '');
                valB = String(b.kelJabatan || '');
                break;
            case 'unitKerja':
                valA = String(a.unitKerja || '');
                valB = String(b.unitKerja || '');
                break;
            case 'jabatan':
                valA = String(a.jabatan || '');
                valB = String(b.jabatan || '');
                break;
            case 'pelaksanaan': {
                const da = parseFlexibleDate(a.pelaksanaan);
                const db = parseFlexibleDate(b.pelaksanaan);
                valA = da ? da.getTime() : 0;
                valB = db ? db.getTime() : 0;
                break;
            }
            case 'sesi': {
                valA = isSesi00A ? 0 : (getCumulativeSessionNumber(a, sortedDates) || 9999);
                valB = isSesi00B ? 0 : (getCumulativeSessionNumber(b, sortedDates) || 9999);
                break;
            }
            case 'waktu':
                valA = String(a.waktu || '');
                valB = String(b.waktu || '');
                break;
            default:
                valA = isSesi00A ? 0 : (getCumulativeSessionNumber(a, sortedDates) || 9999);
                valB = isSesi00B ? 0 : (getCumulativeSessionNumber(b, sortedDates) || 9999);
        }

        let cmp = 0;
        if (typeof valA === 'string' && typeof valB === 'string') {
            cmp = currentSortDirection === 'asc' 
                ? valA.localeCompare(valB, 'id', { sensitivity: 'base', numeric: true }) 
                : valB.localeCompare(valA, 'id', { sensitivity: 'base', numeric: true });
        } else {
            if (valA < valB) cmp = currentSortDirection === 'asc' ? -1 : 1;
            else if (valA > valB) cmp = currentSortDirection === 'asc' ? 1 : -1;
            else cmp = 0;
        }

        // Secondary Sort: Jika nilai sama, selalu urutkan kedua berdasarkan NAMA ASCENDING (A-Z)
        if (cmp === 0) {
            const nameA = String(a.nama || '');
            const nameB = String(b.nama || '');
            return nameA.localeCompare(nameB, 'id', { sensitivity: 'base' });
        }

        return cmp;
    });

    renderCandidateListTable();
    updateActiveFilterStyles();
}

/**
 * Memperbarui efek warna soft (soft highlight) pada kolom filter yang sedang aktif
 */
function updateActiveFilterStyles() {
    // 1. Search Box
    const inputSearch = document.getElementById('inputSearchCandidate');
    if (inputSearch) {
        const isSearchActive = !!(currentSearchTerm && currentSearchTerm.trim().length > 0);
        if (isSearchActive) {
            inputSearch.classList.remove('bg-white', 'border-slate-300', 'text-slate-800');
            inputSearch.classList.add('bg-blue-50', 'border-blue-400', 'text-blue-950', 'font-semibold', 'ring-2', 'ring-blue-100');
        } else {
            inputSearch.classList.remove('bg-blue-50', 'border-blue-400', 'text-blue-950', 'font-semibold', 'ring-2', 'ring-blue-100');
            inputSearch.classList.add('bg-white', 'border-slate-300', 'text-slate-800');
        }
    }

    // 2. Kelompok Jabatan
    const selectKel = document.getElementById('selectFilterKelJabatan');
    if (selectKel) {
        const isKelActive = currentKelJabatanFilter && currentKelJabatanFilter !== 'ALL';
        selectKel.classList.remove(
            'bg-white', 'border-slate-300', 'text-slate-800',
            'bg-emerald-50', 'border-emerald-400', 'text-emerald-950', 'ring-2', 'ring-emerald-100',
            'bg-rose-50', 'border-rose-400', 'text-rose-950', 'ring-rose-100'
        );
        if (isKelActive) {
            if (currentKelJabatanFilter === 'EMPTY') {
                selectKel.classList.add('bg-rose-50', 'border-rose-400', 'text-rose-950', 'font-bold', 'ring-2', 'ring-rose-100');
            } else {
                selectKel.classList.add('bg-emerald-50', 'border-emerald-400', 'text-emerald-950', 'font-bold', 'ring-2', 'ring-emerald-100');
            }
        } else {
            selectKel.classList.add('bg-white', 'border-slate-300', 'text-slate-800');
        }
    }

    // 3. Tanggal Pelaksanaan
    const selectDate = document.getElementById('selectFilterPelaksanaan');
    if (selectDate) {
        const isDateActive = currentDateFilter && currentDateFilter !== 'ALL';
        if (isDateActive) {
            selectDate.classList.remove('bg-white', 'border-slate-300', 'text-slate-800');
            selectDate.classList.add('bg-indigo-50', 'border-indigo-400', 'text-indigo-950', 'font-bold', 'ring-2', 'ring-indigo-100');
        } else {
            selectDate.classList.remove('bg-indigo-50', 'border-indigo-400', 'text-indigo-950', 'font-bold', 'ring-2', 'ring-indigo-100');
            selectDate.classList.add('bg-white', 'border-slate-300', 'text-slate-800');
        }
    }

    // 4. Sesi (Typing + Dropdown)
    const inputSesiTyping = document.getElementById('inputFilterSesiTyping');
    const selectSesi = document.getElementById('selectFilterSesiDropdown');
    const isSesiActive = (currentCumulativeSessionFilter && currentCumulativeSessionFilter !== 'ALL') ||
                         (currentSessionFilter && currentSessionFilter !== 'ALL') ||
                         (inputSesiTyping && inputSesiTyping.value.trim().length > 0);

    if (inputSesiTyping) {
        if (isSesiActive) {
            inputSesiTyping.classList.remove('bg-white', 'border-slate-300', 'text-bkn-900');
            inputSesiTyping.classList.add('bg-amber-50', 'border-amber-400', 'text-amber-950', 'ring-2', 'ring-amber-200');
        } else {
            inputSesiTyping.classList.remove('bg-amber-50', 'border-amber-400', 'text-amber-950', 'ring-2', 'ring-amber-200');
            inputSesiTyping.classList.add('bg-white', 'border-slate-300', 'text-bkn-900');
        }
    }

    if (selectSesi) {
        if (isSesiActive) {
            selectSesi.classList.remove('bg-white', 'border-slate-300', 'text-slate-800');
            selectSesi.classList.add('bg-amber-50', 'border-amber-400', 'text-amber-950', 'font-bold', 'ring-2', 'ring-amber-200');
        } else {
            selectSesi.classList.remove('bg-amber-50', 'border-amber-400', 'text-amber-950', 'font-bold', 'ring-2', 'ring-amber-200');
            selectSesi.classList.add('bg-white', 'border-slate-300', 'text-slate-800');
        }
    }

    // 5. Tombol Reset Filter
    const btnReset = document.getElementById('btnResetCandidateFilter');
    if (btnReset) {
        const isSearchActive = !!(currentSearchTerm && currentSearchTerm.trim().length > 0);
        const isKelActive = currentKelJabatanFilter && currentKelJabatanFilter !== 'ALL';
        const isDateActive = currentDateFilter && currentDateFilter !== 'ALL';
        const anyActive = isSearchActive || isKelActive || isDateActive || isSesiActive;

        if (anyActive) {
            btnReset.classList.remove('bg-slate-100', 'border-slate-300', 'text-slate-700', 'hover:bg-slate-200', 'hover:text-slate-900');
            btnReset.classList.add('bg-rose-50', 'border-rose-300', 'text-rose-700', 'hover:bg-rose-100', 'hover:border-rose-400', 'hover:text-rose-800', 'ring-2', 'ring-rose-100');
        } else {
            btnReset.classList.remove('bg-rose-50', 'border-rose-300', 'text-rose-700', 'hover:bg-rose-100', 'hover:border-rose-400', 'hover:text-rose-800', 'ring-2', 'ring-rose-100');
            btnReset.classList.add('bg-slate-100', 'border-slate-300', 'text-slate-700', 'hover:bg-slate-200', 'hover:text-slate-900');
        }
    }
}
window.updateActiveFilterStyles = updateActiveFilterStyles;

function renderCandidateListTable() {
    const tbody = document.getElementById('tbodyCandidateList');
    const countBadge = document.getElementById('countTableVisible');
    const paginationInfo = document.getElementById('tablePaginationInfo');

    if (!tbody) return;

    if (countBadge) countBadge.textContent = `${filteredCandidates.length} Data`;
    if (paginationInfo) paginationInfo.textContent = `Menampilkan ${filteredCandidates.length} dari ${currentCandidates.length} total peserta`;

    // Hitung peserta dengan Kelompok Jabatan kosong
    const countKelEmptyBadge = document.getElementById('countTableKelEmpty');
    const emptyKelCount = filteredCandidates.filter(c => {
        const raw = String(c.kelJabatan || '').trim();
        return !raw || raw === '-' || raw === 'NULL';
    }).length;
    if (countKelEmptyBadge) {
        const formattedEmpty = emptyKelCount < 10 ? `0${emptyKelCount}` : `${emptyKelCount}`;
        countKelEmptyBadge.textContent = `${formattedEmpty} Data`;
    }

    if (!currentExam) {
        tbody.innerHTML = `
            <tr>
                <td colspan="10" class="p-8 text-center text-slate-400">
                    <i data-lucide="lock" class="w-8 h-8 mx-auto mb-2 text-slate-300"></i>
                    <p class="font-bold text-slate-700 text-sm">Pilih Instansi Ujian & Masukkan PIN</p>
                    <p class="text-xs text-slate-500 mt-1">Data peserta hanya akan dimuat setelah instansi dipilih dan PIN berhasil diverifikasi.</p>
                </td>
            </tr>
        `;
        if (window.lucide) window.lucide.createIcons();
        return;
    }

    if (currentCandidates.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="10" class="p-8 text-center text-slate-400">
                    <i data-lucide="inbox" class="w-8 h-8 mx-auto mb-2 text-slate-300"></i>
                    <p class="font-bold text-slate-700 text-sm">Belum Ada Data Peserta</p>
                    <p class="text-xs text-slate-500 mt-1">Belum ada peserta yang diunggah untuk instansi <strong>${currentExam.instansi}</strong>.</p>
                </td>
            </tr>
        `;
        if (window.lucide) window.lucide.createIcons();
        return;
    }

    if (filteredCandidates.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="10" class="p-8 text-center text-slate-400">
                    <i data-lucide="search-x" class="w-8 h-8 mx-auto mb-2 text-slate-300"></i>
                    <p class="font-bold text-slate-700 text-sm">Tidak Ada Data Peserta</p>
                    <p class="text-xs text-slate-500 mt-1">Tidak ada peserta yang cocok dengan kriteria filter pencarian saat ini.</p>
                </td>
            </tr>
        `;
        if (window.lucide) window.lucide.createIcons();
        return;
    }

    const sortedDates = getSortedExamDates();

    tbody.innerHTML = filteredCandidates.map((c, idx) => {
        const isSesi00 = !c.sesi || c.sesi === 'NULL' || c.sesi === '00' || c.sesi === 0 || c.sesi === '0';
        const isNullDate = !c.pelaksanaan || c.pelaksanaan === 'NULL' || c.pelaksanaan === '-';
        const isNullSchedule = isNullDate || isSesi00;
        const isFriSession2 = !isNullSchedule && c.isFriday && Number(c.sesi) === 2;
        const cumSesi = isNullSchedule ? null : getCumulativeSessionNumber(c, sortedDates);
        const cumSesiFormatted = cumSesi ? formatCumulativeSessionNumber(cumSesi) : '00';
        const isNullUnit = !c.unitKerja || c.unitKerja === 'NULL' || c.unitKerja === '-';

        const rawKel = String(c.kelJabatan || '').trim();
        const isKelEmpty = !rawKel || rawKel === '-' || rawKel === 'NULL';

        const sesiColorBadge = Number(c.sesi) === 1 
            ? 'bg-blue-100 text-blue-800 border border-blue-200' 
            : (Number(c.sesi) === 2 
                ? 'bg-amber-100 text-amber-800 border border-amber-200' 
                : 'bg-emerald-100 text-emerald-800 border border-emerald-200');

        const rowBgClass = isKelEmpty 
            ? 'bg-rose-50/70 border-l-4 border-l-red-900 hover:bg-rose-100/60' 
            : (isFriSession2 ? 'bg-amber-50/60' : 'hover:bg-slate-50');

        return `
            <tr class="${rowBgClass} transition">
                <td class="p-3 text-center text-slate-500 font-medium">${idx + 1}</td>
                <td class="p-2.5 text-center whitespace-nowrap">
                    ${(() => {
                        const candidateKey = String(c.nip || c.id || '').trim();
                        if (isKelEmpty) {
                            return `
                                <span class="inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-[11px] font-semibold bg-slate-100 text-slate-500 border border-slate-200 select-none" title="Presensi tidak tersedia karena peserta belum terdaftar di sistem. Silakan lengkapi kelompok jabatan terlebih dahulu.">
                                    <i data-lucide="slash" class="w-3 h-3 text-slate-400"></i>
                                    <span>Tidak tersedia</span>
                                </span>
                            `;
                        }
                        if (c.kehadiran === 'HADIR') {
                            return `
                                <button onclick="toggleAttendance('${candidateKey}', 'RESET')" class="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800 border border-emerald-300 hover:bg-emerald-200 transition shadow-2xs cursor-pointer" title="Status: Hadir. Klik untuk ubah/batal">
                                    <i data-lucide="check" class="w-3.5 h-3.5 stroke-[3]"></i>
                                    <span>Hadir</span>
                                </button>
                            `;
                        }
                        if (c.kehadiran === 'TIDAK_HADIR') {
                            return `
                                <button onclick="toggleAttendance('${candidateKey}', 'RESET')" class="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-bold bg-rose-100 text-rose-800 border border-rose-300 hover:bg-rose-200 transition shadow-2xs cursor-pointer" title="Status: Tidak Hadir. Klik untuk ubah/batal">
                                    <i data-lucide="x" class="w-3.5 h-3.5 stroke-[3]"></i>
                                    <span>Tidak Hadir</span>
                                </button>
                            `;
                        }
                        return `
                            <div class="inline-flex items-center justify-center gap-1.5">
                                <button onclick="toggleAttendance('${candidateKey}', 'HADIR')" class="p-1.5 rounded-lg bg-emerald-50 hover:bg-emerald-600 hover:text-white text-emerald-600 border border-emerald-300 transition shadow-2xs cursor-pointer" title="Tandai Hadir">
                                    <i data-lucide="check" class="w-4 h-4 stroke-[2.5]"></i>
                                </button>
                                <button onclick="toggleAttendance('${candidateKey}', 'TIDAK_HADIR')" class="p-1.5 rounded-lg bg-rose-50 hover:bg-rose-600 hover:text-white text-rose-600 border border-rose-300 transition shadow-2xs cursor-pointer" title="Tandai Tidak Hadir">
                                    <i data-lucide="x" class="w-4 h-4 stroke-[2.5]"></i>
                                </button>
                            </div>
                        `;
                    })()}
                </td>
                <td class="p-3 font-mono font-medium text-slate-900">${c.nip}</td>
                <td class="p-3 font-bold text-slate-900">${c.nama}</td>
                <td class="p-3">
                    ${isKelEmpty ? `
                        <span class="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-[11px] font-bold bg-red-900 text-red-100 border border-red-950 shadow-xs tracking-wide whitespace-nowrap" title="Kelompok Jabatan belum diisi / Peserta Belum Terdaftar di File Sistem">
                            <i data-lucide="alert-circle" class="w-3.5 h-3.5 text-red-200 stroke-[2.5]"></i>
                            <span>Peserta Belum Terdaftar</span>
                        </span>
                    ` : `
                        <span class="font-semibold text-blue-800 bg-blue-50/60 px-2 py-0.5 rounded border border-blue-200/50">${c.kelJabatan}</span>
                    `}
                </td>
                <td class="p-3 text-slate-600 max-w-[220px] truncate" title="${c.unitKerja || '-'}">
                    ${isNullUnit ? `
                        <span class="px-2 py-0.5 rounded text-[10px] font-bold bg-slate-100 text-slate-500 border border-slate-200">NULL</span>
                    ` : c.unitKerja}
                </td>
                <td class="p-3 text-slate-600 max-w-[200px] truncate" title="${c.jabatan || '-'}">${c.jabatan || '-'}</td>
                <td class="p-3 whitespace-nowrap">
                    ${isNullDate ? `
                        <span class="px-2 py-0.5 rounded text-[10px] font-bold bg-slate-100 text-slate-500 border border-slate-200">NULL</span>
                    ` : `
                        <span class="font-semibold text-slate-800">${c.pelaksanaan}</span>
                        ${c.isFriday ? '<span class="text-[10px] bg-amber-100 text-amber-800 font-bold px-1.5 py-0.5 rounded ml-1">Jumat</span>' : ''}
                    `}
                </td>
                <td class="p-3 text-center whitespace-nowrap min-w-[130px]">
                    ${isSesi00 ? `
                        <div class="inline-flex items-center justify-center gap-1.5 whitespace-nowrap">
                            <span class="inline-block whitespace-nowrap px-2.5 py-1 rounded-md text-xs font-bold bg-slate-100 text-slate-600 border border-slate-300">
                                Sesi 00
                            </span>
                            <span class="inline-block whitespace-nowrap px-2 py-1 rounded-md text-xs font-extrabold bg-slate-400 text-white shadow-xs border border-slate-400" title="Sesi Kumulatif: 00 (NULL)">
                                00
                            </span>
                        </div>
                    ` : `
                        <div class="inline-flex items-center justify-center gap-1.5 whitespace-nowrap">
                            <span class="inline-block whitespace-nowrap px-2.5 py-1 rounded-md text-xs font-bold ${sesiColorBadge}">
                                Sesi ${c.sesi}
                            </span>
                            <span class="inline-block whitespace-nowrap px-2 py-1 rounded-md text-xs font-extrabold bg-slate-800 text-white shadow-xs border border-slate-700" title="Sesi Kumulatif: ${cumSesiFormatted}">
                                ${cumSesiFormatted}
                            </span>
                        </div>
                    `}
                </td>
                <!-- Kolom Waktu (WIT) di-hide. Untuk mengaktifkan kembali, hapus class 'hidden' -->
                <td class="p-3 whitespace-nowrap hidden ${isFriSession2 ? 'font-bold text-amber-800' : 'text-slate-700 font-medium'}">
                    ${(!c.waktu || c.waktu === 'NULL' || c.waktu === '-') ? '<span class="px-2 py-0.5 rounded text-[10px] font-bold bg-slate-100 text-slate-500 border border-slate-200">NULL</span>' : c.waktu}
                </td>
                <td class="p-3 text-center whitespace-nowrap">
                    ${(() => {
                        const candidateKey = String(c.nip || c.id || '').trim();
                        const safeNama = String(c.nama || '').replace(/'/g, "\\'");
                        return `
                            <button onclick="editCandidate('${candidateKey}')" class="p-1 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded mr-1" title="Edit Data">
                                <i data-lucide="edit-2" class="w-3.5 h-3.5"></i>
                            </button>
                            <button onclick="deleteSingleCandidate('${candidateKey}', '${safeNama}')" class="p-1 text-rose-600 hover:text-rose-800 hover:bg-rose-50 rounded" title="Hapus Peserta">
                                <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                            </button>
                        `;
                    })()}
                </td>
            </tr>
        `;
    }).join('');

    if (window.lucide) window.lucide.createIcons();
}

window.deleteSingleCandidate = async (candidateIdOrNip, name) => {
    if (confirm(`Hapus peserta "${name}" dari jadwal ujian?`)) {
        try {
            const safeKey = String(candidateIdOrNip || '').trim();
            await db.deleteCandidate(safeKey, currentExam ? currentExam.id : null);
            currentCandidates = currentCandidates.filter(c => String(c.nip || '').trim() !== safeKey && String(c.id || '').trim() !== safeKey);
            renderDashboardStats();
            populatePelaksanaanFilterDropdown();
            populateSesiFilterDropdown(currentDateFilter);
            applyCandidateFilters();
            showToast(`Peserta "${name}" berhasil dihapus.`, "info");
        } catch (err) {
            console.error(err);
            showToast("Gagal menghapus peserta: " + err.message, "error");
        }
    }
};

window.confirmClearCandidates = () => {
    if (!currentExam) {
        showToast("Pilih instansi ujian aktif terlebih dahulu!", "warning");
        return;
    }
    if (currentCandidates.length === 0) {
        showToast(`Data peserta untuk ${currentExam.instansi} sudah kosong.`, "info");
        return;
    }

    // SELALU munculkan modal PIN 1414 untuk konfirmasi otorisasi tindakan permanen ini
    pendingActionAfterPin = 'CLEAR_CANDIDATES';
    const modal = document.getElementById('modalPinAccess');
    const inputPin = document.getElementById('inputAccessPin');
    const errorMsg = document.getElementById('pinErrorMessage');
    const titleEl = document.getElementById('modalPinTitle');
    const descEl = document.getElementById('modalPinDesc');

    if (titleEl) titleEl.textContent = "Konfirmasi Kosongkan Peserta";
    if (descEl) descEl.textContent = `PERINGATAN: Seluruh (${currentCandidates.length}) peserta untuk "${currentExam.instansi}" akan dihapus permanen dari database. Masukkan PIN User Admin untuk mengonfirmasi.`;

    if (errorMsg) errorMsg.classList.add('hidden');
    if (inputPin) {
        inputPin.value = '';
        inputPin.classList.remove('border-rose-500');
    }

    if (modal) {
        modal.classList.remove('hidden');
        modal.classList.add('flex');
        setTimeout(() => {
            if (inputPin) inputPin.focus();
        }, 100);
    }

    if (window.lucide) window.lucide.createIcons();
};

async function executeClearCandidates() {
    if (!currentExam) return;
    try {
        showToast("Sedang mengosongkan seluruh data peserta...", "info");
        await db.deleteCandidatesByExam(currentExam.id);
        currentCandidates = [];
        filteredCandidates = [];
        renderDashboardStats();
        populatePelaksanaanFilterDropdown();
        populateSesiFilterDropdown('ALL');
        populateKelJabatanFilterDropdown();
        applyCandidateFilters();
        showToast(`Semua data peserta "${currentExam.instansi}" berhasil dikosongkan.`, "success");
    } catch (err) {
        console.error("Gagal mengosongkan peserta:", err);
        showToast("Gagal mengosongkan peserta: " + err.message, "error");
    }
}

/**
 * Setup Modal & Form Tambah/Edit Peserta Manual
 */
function setupManualCandidateForm() {
    const form = document.getElementById('formCandidateManual');
    const inputDate = document.getElementById('inputManualPelaksanaan');
    const selectSesi = document.getElementById('selectManualSesi');
    const inputWaktu = document.getElementById('inputManualWaktu');

    function autoCalculateTime() {
        const d = inputDate.value.trim();
        const s = selectSesi.value;
        if (d && s) {
            inputWaktu.value = getSessionTime(s, d);
        }
    }

    if (inputDate) {
        inputDate.addEventListener('input', autoCalculateTime);
        inputDate.addEventListener('change', autoCalculateTime);
    }
    if (selectSesi) selectSesi.addEventListener('change', autoCalculateTime);

    if (form) {
        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            if (!currentExam) {
                showToast("Harap pilih atau buat ujian terlebih dahulu!", "warning");
                return;
            }

            const id = document.getElementById('editCandidateId').value;
            const nip = document.getElementById('inputManualNip').value.trim();
            const nama = document.getElementById('inputManualNama').value.trim();
            const kelJabatan = document.getElementById('inputManualKelJabatan') ? document.getElementById('inputManualKelJabatan').value.trim() : '';
            const unitKerja = document.getElementById('inputManualUnitKerja').value.trim();
            const jabatan = document.getElementById('inputManualJabatan').value.trim();
            const rawPelaksanaan = document.getElementById('inputManualPelaksanaan').value.trim();
            const parsedPel = parseFlexibleDate(rawPelaksanaan);
            const pelaksanaan = parsedPel ? formatDateDisplay(parsedPel, 'short') : (rawPelaksanaan || '-');
            const sesi = Number(document.getElementById('selectManualSesi').value) || 1;
            const waktu = document.getElementById('inputManualWaktu').value.trim() || getSessionTime(sesi, pelaksanaan);
            const fri = isFriday(pelaksanaan);

            const row = {
                examId: currentExam.id,
                nip,
                nama,
                kelJabatan: kelJabatan || '-',
                unitKerja,
                jabatan,
                pelaksanaan,
                sesi,
                waktu,
                isFriday: fri,
                status: 'Terjadwal'
            };

            try {
                if (id) {
                    row.id = String(id).trim();
                    await db.updateCandidate(row);
                    showToast(`Data peserta "${nama}" berhasil diperbarui.`, "success");
                } else {
                    row.no = currentCandidates.length + 1;
                    const newId = await db.addCandidate(row);
                    row.id = String(newId || nip).trim();
                    showToast(`Peserta "${nama}" berhasil ditambahkan.`, "success");
                }

                currentCandidates = await db.getCandidatesByExam(currentExam.id);
                renderDashboardStats();
                populatePelaksanaanFilterDropdown();
                populateSesiFilterDropdown(currentDateFilter);
                populateKelJabatanFilterDropdown();
                applyCandidateFilters();
                closeModalCandidateManual();

            } catch (err) {
                console.error(err);
                showToast("Gagal menyimpan peserta: " + err.message, "error");
            }
        });
    }
}

window.openModalAddCandidate = () => {
    const modal = document.getElementById('modalCandidateManual');
    const title = document.getElementById('modalCandidateTitle');
    const form = document.getElementById('formCandidateManual');
    if (!modal || !form) return;

    title.textContent = "Tambah Peserta Manual";
    form.reset();
    document.getElementById('editCandidateId').value = '';
    if (document.getElementById('inputManualKelJabatan')) {
        document.getElementById('inputManualKelJabatan').value = '';
    }

    const defDate = (currentExam && currentExam.startDate) ? parseFlexibleDate(currentExam.startDate) : new Date();
    const isoDate = dateToISOInput(defDate);
    document.getElementById('inputManualPelaksanaan').value = isoDate;
    document.getElementById('inputManualWaktu').value = getSessionTime(1, isoDate || defDate);

    modal.classList.remove('hidden');
    modal.classList.add('flex');
};

window.editCandidate = (candidateIdOrNip) => {
    const safeLookup = String(candidateIdOrNip || '').trim();
    const cand = currentCandidates.find(c => String(c.nip || '').trim() === safeLookup || String(c.id || '').trim() === safeLookup);
    if (!cand) return;

    const modal = document.getElementById('modalCandidateManual');
    const title = document.getElementById('modalCandidateTitle');
    if (!modal) return;

    title.textContent = "Edit Data Peserta";
    document.getElementById('editCandidateId').value = cand.id;
    document.getElementById('inputManualNip').value = cand.nip;
    document.getElementById('inputManualNama').value = cand.nama;
    if (document.getElementById('inputManualKelJabatan')) {
        document.getElementById('inputManualKelJabatan').value = cand.kelJabatan && cand.kelJabatan !== '-' ? cand.kelJabatan : '';
    }
    document.getElementById('inputManualUnitKerja').value = cand.unitKerja || '';
    document.getElementById('inputManualJabatan').value = cand.jabatan || '';
    document.getElementById('inputManualPelaksanaan').value = dateToISOInput(cand.pelaksanaan);
    document.getElementById('selectManualSesi').value = cand.sesi || 1;
    document.getElementById('inputManualWaktu').value = cand.waktu || getSessionTime(cand.sesi || 1, cand.pelaksanaan);

    modal.classList.remove('hidden');
    modal.classList.add('flex');
};

window.closeModalCandidateManual = () => {
    const modal = document.getElementById('modalCandidateManual');
    if (modal) {
        modal.classList.add('hidden');
        modal.classList.remove('flex');
    }
};

/**
 * Export Peserta Ujian ke Excel
 */
window.exportCurrentCandidates = () => {
    if (!currentExam || filteredCandidates.length === 0) {
        showToast("Tidak ada data peserta untuk diekspor!", "warning");
        return;
    }
    exportCandidatesToExcel(currentExam.instansi, filteredCandidates);
    showToast("File Excel berhasil di-generate!", "success");
};

/**
 * Cetak Lembar Resmi Presensi Peserta Ujian Profiling ASN
 * Menarik kondisi peserta yang sudah ditandai HADIR dan TIDAK HADIR saja.
 * Menyesuaikan filter yang sedang aktif di daftar peserta, dan membuat tabel terpisah untuk setiap sesinya.
 */
window.printOfficialSchedule = () => {
    if (!currentExam) {
        showToast("Pilih instansi ujian terlebih dahulu!", "warning");
        return;
    }

    // 1. Tarik HANYA peserta yang berstatus HADIR dan TIDAK HADIR dari filter aktif
    const presensiCandidates = filteredCandidates.filter(c => c.kehadiran === 'HADIR' || c.kehadiran === 'TIDAK_HADIR');

    if (presensiCandidates.length === 0) {
        showToast("Tidak ada data peserta dengan status Hadir atau Tidak Hadir untuk dicetak pada filter aktif!", "warning");
        return;
    }

    const printContainer = document.getElementById('officialPrintArea');
    if (!printContainer) return;

    const sortedDates = getSortedExamDates();

    // 2. Kelompokkan peserta berdasarkan tanggal pelaksanaan dan sesi ujian
    const sessionsMap = new Map();
    presensiCandidates.forEach(c => {
        const cum = getCumulativeSessionNumber(c, sortedDates) || 0;
        const key = `${c.pelaksanaan || 'Tanpa Tanggal'}___${c.sesi || 0}`;
        if (!sessionsMap.has(key)) {
            sessionsMap.set(key, {
                key: key,
                cumNum: cum,
                cumFormatted: formatCumulativeSessionNumber(cum),
                dailySession: c.sesi || '-',
                date: c.pelaksanaan || '-',
                waktu: c.waktu || '-',
                candidates: []
            });
        }
        sessionsMap.get(key).candidates.push(c);
    });

    // Urutkan grup sesi secara kronologis tanggal lalu nomor sesi
    const sortedSessionGroups = Array.from(sessionsMap.values()).sort((a, b) => {
        const da = parseFlexibleDate(a.date);
        const db = parseFlexibleDate(b.date);
        const ta = da ? da.getTime() : 0;
        const tb = db ? db.getTime() : 0;
        if (ta !== tb) return ta - tb;
        return Number(a.dailySession || 0) - Number(b.dailySession || 0);
    });

    const todayStr = formatDateDisplay(new Date(), 'long');
    let fullHtml = '';

    sortedSessionGroups.forEach((group) => {
        // Urutkan nama peserta A-Z dalam setiap sesi
        const sortedList = [...group.candidates].sort((a, b) => 
            String(a.nama || '').localeCompare(String(b.nama || ''), 'id', { sensitivity: 'base' })
        );

        const dayName = group.date ? getDayNameID(group.date) : '';
        const countHadir = sortedList.filter(c => c.kehadiran === 'HADIR').length;
        const countTidakHadir = sortedList.filter(c => c.kehadiran === 'TIDAK_HADIR').length;
        const totalPesertaSesi = sortedList.length;

        fullHtml += `
            <div class="print-session-page">
                <!-- KOP RESMI BKN -->
                <div style="border-bottom: 2px solid #000; padding-bottom: 8px; margin-bottom: 12px; text-align: center;">
                    <div style="font-size: 11pt; font-weight: bold; letter-spacing: 0.5px;">BADAN KEPEGAWAIAN NEGARA</div>
                    <div style="font-size: 10pt; font-weight: bold;">KANTOR REGIONAL XIV MANOKWARI</div>
                    <div style="font-size: 12pt; font-weight: 800; margin-top: 4px; text-decoration: underline;">DAFTAR PRESENSI PESERTA UJIAN PROFILING ASN</div>
                </div>

                <!-- HEADER KETERANGAN SESI PRESENSI -->
                <table style="width: 100%; font-size: 8.5pt; margin-bottom: 8px; border: none;">
                    <tr>
                        <td style="width: 18%; font-weight: bold; padding: 2px 0;">Instansi</td>
                        <td style="width: 2%; padding: 2px 0;">:</td>
                        <td style="width: 45%; font-weight: bold; padding: 2px 0;">${currentExam.instansi}</td>
                        <td style="width: 15%; font-weight: bold; padding: 2px 0;">Sesi Ujian</td>
                        <td style="width: 2%; padding: 2px 0;">:</td>
                        <td style="width: 18%; font-weight: bold; color: #1e3a8a; padding: 2px 0;">Sesi ${group.dailySession} ${group.cumFormatted ? `(${group.cumFormatted})` : ''}</td>
                    </tr>
                    <tr>
                        <td style="font-weight: bold; padding: 2px 0;">Titik Lokasi</td>
                        <td style="padding: 2px 0;">:</td>
                        <td style="padding: 2px 0;">${currentExam.location}</td>
                        <td style="font-weight: bold; padding: 2px 0;">Waktu Ujian</td>
                        <td style="padding: 2px 0;">:</td>
                        <td style="padding: 2px 0;">${group.waktu}</td>
                    </tr>
                    <tr>
                        <td style="font-weight: bold; padding: 2px 0;">Hari / Tanggal</td>
                        <td style="padding: 2px 0;">:</td>
                        <td style="padding: 2px 0;">${dayName ? `${dayName}, ` : ''}${group.date}</td>
                        <td style="font-weight: bold; padding: 2px 0;">Total Presensi</td>
                        <td style="padding: 2px 0;">:</td>
                        <td style="padding: 2px 0; font-weight: bold;">${totalPesertaSesi} Orang <span style="font-weight: normal; color: #555;">(Hadir: ${countHadir}, Tidak Hadir: ${countTidakHadir})</span></td>
                    </tr>
                </table>

                <!-- TABEL PRESENSI PESERTA SESI INI -->
                <table class="print-table">
                    <thead>
                        <tr>
                            <th style="width: 28px;">No</th>
                            <th style="width: 130px;">NIP</th>
                            <th>Nama Peserta</th>
                            <th style="width: 105px;">Kel. Jabatan</th>
                            <th>Unit Kerja / Jabatan</th>
                            <th style="width: 100px;">Status Presensi</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${sortedList.map((c, idx) => {
                            const isHadir = c.kehadiran === 'HADIR';
                            const statusBadge = isHadir
                                ? '<span style="display: inline-block; padding: 2px 6px; font-weight: 800; color: #047857; background-color: #ecfdf5; border: 1px solid #a7f3d0; border-radius: 4px; font-size: 7.5pt;">HADIR</span>'
                                : '<span style="display: inline-block; padding: 2px 6px; font-weight: 800; color: #b91c1c; background-color: #fef2f2; border: 1px solid #fecaca; border-radius: 4px; font-size: 7.5pt;">TIDAK HADIR</span>';
                            
                            return `
                                <tr>
                                    <td style="text-align: center;">${idx + 1}</td>
                                    <td style="font-family: monospace; text-align: center; font-size: 7.5pt;">${c.nip}</td>
                                    <td style="font-weight: bold;">${c.nama}</td>
                                    <td style="text-align: center; font-size: 7.5pt; font-weight: 600;">${c.kelJabatan && c.kelJabatan !== '-' ? c.kelJabatan : '-'}</td>
                                    <td>${c.unitKerja && c.unitKerja !== 'NULL' ? c.unitKerja : '-'}${c.jabatan ? `<br><span style="font-size: 7.5pt; color: #555;">${c.jabatan}</span>` : ''}</td>
                                    <td style="text-align: center; vertical-align: middle;">
                                        ${statusBadge}
                                    </td>
                                </tr>
                            `;
                        }).join('')}
                    </tbody>
                </table>

                <!-- RINGKASAN PRESENSI & TANDA TANGAN PENANGGUNG JAWAB -->
                <div style="margin-top: 14px; display: flex; justify-content: space-between; align-items: flex-start; page-break-inside: avoid;">
                    <div style="font-size: 8pt; color: #333; padding: 4px 8px; border: 1px dashed #999; border-radius: 4px; max-width: 320px;">
                        <div style="font-weight: bold; margin-bottom: 2px;">Rekapitulasi Sesi Ini:</div>
                        <div>Peserta Hadir: <strong>${countHadir} Orang</strong></div>
                        <div>Peserta Tidak Hadir: <strong>${countTidakHadir} Orang</strong></div>
                        <div>Total Sesi: <strong>${totalPesertaSesi} Orang</strong></div>
                    </div>

                    <div style="width: 240px; text-align: center; font-size: 8.5pt;">
                        <div>Manokwari, ${todayStr}</div>
                        <div style="margin-top: 4px; font-weight: bold;">Koordinator Tim Pelaksana CAT BKN,</div>
                        <div style="height: 44px;"></div>
                        <div style="border-bottom: 1px solid #000; font-weight: bold;">( ..................................................... )</div>
                        <div style="font-size: 7.5pt; color: #555; margin-top: 2px;">NIP. .................................................</div>
                    </div>
                </div>
            </div>
        `;
    });

    printContainer.innerHTML = fullHtml;
    window.print();
};

/**
 * Setup Navigasi Tab & Proteksi PIN Akses ("PIN")
 */
function setupTabNavigation() {
    window.requestSwitchTab = (tabName) => {
        // Tab Jadwal & Peserta serta Dashboard bebas diakses langsung tanpa PIN
        if (tabName === 'daftar-peserta' || tabName === 'dashboard') {
            window.switchTab(tabName);
            return;
        }

        // Jika PIN sudah berhasil di-unlock di sesi ini atau Super Admin, langsung izinkan
        if (isPinAuthorized || isSuperAdmin) {
            window.switchTab(tabName);
            return;
        }

        // Tampilkan modal PIN pop up di tengah layar
        pendingTargetTab = tabName;
        const modal = document.getElementById('modalPinAccess');
        const inputPin = document.getElementById('inputAccessPin');
        const errorMsg = document.getElementById('pinErrorMessage');
        const titleEl = document.getElementById('modalPinTitle');
        const descEl = document.getElementById('modalPinDesc');

        if (titleEl) titleEl.textContent = "Akses Menu Terkunci";
        if (descEl) descEl.textContent = "Masukkan PIN Otorisasi Administrator untuk mengakses menu ini.";

        if (errorMsg) errorMsg.classList.add('hidden');
        if (inputPin) {
            inputPin.value = '';
            inputPin.classList.remove('border-rose-500');
        }

        if (modal) {
            modal.classList.remove('hidden');
            modal.classList.add('flex');
            setTimeout(() => {
                if (inputPin) inputPin.focus();
            }, 100);
        }

        if (window.lucide) window.lucide.createIcons();
    };

    window.verifyPinAndProceed = async (event) => {
        if (event) event.preventDefault();
        const inputPin = document.getElementById('inputAccessPin');
        const errorMsg = document.getElementById('pinErrorMessage');
        const pinVal = inputPin ? inputPin.value.trim() : '';

        if (pinVal === '1414' || pinVal === '141414') {
            isPinAuthorized = true;

            // Simpan aksi dan target tab tertunda sebelum menutup modal
            const actionToExecute = pendingActionAfterPin;
            const targetTabToSwitch = pendingTargetTab;

            window.closeModalPinAccess();

            // 1. Eksekusi Kosongkan Peserta jika aksi tertunda adalah CLEAR_CANDIDATES
            if (actionToExecute === 'CLEAR_CANDIDATES') {
                await executeClearCandidates();
                return;
            }

            showToast("Akses administrator berhasil dibuka!", "success");

            // 2. Buka Pengaturan Cloud jika aksi tertunda adalah OPEN_FIREBASE_CONFIG
            if (actionToExecute === 'OPEN_FIREBASE_CONFIG') {
                window.openModalFirebaseConfig();
                return;
            }

            // 3. Pindah tab jika ada tab target tertunda
            if (targetTabToSwitch) {
                window.switchTab(targetTabToSwitch);
            }
        } else {
            if (errorMsg) errorMsg.classList.remove('hidden');
            if (inputPin) {
                inputPin.classList.add('border-rose-500');
                inputPin.value = '';
                inputPin.focus();
            }
        }
    };

    window.closeModalPinAccess = () => {
        const modal = document.getElementById('modalPinAccess');
        if (modal) {
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        // Kembalikan teks modal PIN ke default
        const titleEl = document.getElementById('modalPinTitle');
        const descEl = document.getElementById('modalPinDesc');
        if (titleEl) titleEl.textContent = "Akses Terkunci";
        if (descEl) descEl.textContent = "Masukkan PIN Otorisasi Administrator untuk mengakses menu ini.";

        pendingTargetTab = null;
        pendingActionAfterPin = null;
    };

    window.switchTab = (tabName) => {
        document.querySelectorAll('.nav-tab').forEach(btn => {
            btn.classList.remove('border-amber-400', 'text-white');
            btn.classList.add('border-transparent', 'text-blue-200');
        });

        const activeNav = document.getElementById(`nav-${tabName}`);
        if (activeNav) {
            activeNav.classList.remove('border-transparent', 'text-blue-200');
            activeNav.classList.add('border-amber-400', 'text-white');
        }

        document.querySelectorAll('.tab-pane').forEach(pane => {
            pane.classList.add('hidden');
        });

        const activePane = document.getElementById(`pane-${tabName}`);
        if (activePane) {
            activePane.classList.remove('hidden');
        }

        if (tabName === 'master-wilker') {
            renderMasterInstansiTableFull();
        } else if (tabName === 'daftar-peserta') {
            applyCandidateFilters();
        } else if (tabName === 'dashboard') {
            renderDashboardStats();
        } else if (tabName === 'create-ujian') {
            renderExamListInCreateTab();
        }

        if (window.lucide) window.lucide.createIcons();
    };
}

// ---------------------- MODAL PILIH UJIAN & PIN INSTANSI ----------------------

/**
 * Memperbarui opsi pilihan instansi di modal PIN jika modal sedang terbuka
 */
function refreshModalSelectExamPicker() {
    const selectPicker = document.getElementById('modalSelectExamPicker');
    const modal = document.getElementById('modalSelectExamWithPin');
    if (selectPicker && modal && !modal.classList.contains('hidden') && allExams.length > 0) {
        const prevVal = selectPicker.value;
        selectPicker.innerHTML = allExams.map(e => {
            const datePart = e.startDate ? ` (${formatDateDisplay(parseFlexibleDate(e.startDate) || e.startDate, 'short')})` : '';
            return `<option value="${e.id}">${e.instansi}${datePart}</option>`;
        }).join('');

        if (prevVal && allExams.some(e => e.id === prevVal)) {
            selectPicker.value = prevVal;
        } else if (currentExam) {
            selectPicker.value = currentExam.id;
        } else {
            selectPicker.value = allExams[0].id;
        }
    }
}

window.openSelectExamWithPinModal = (preselectedExamId = null, canCancel = false) => {
    const modal = document.getElementById('modalSelectExamWithPin');
    const selectPicker = document.getElementById('modalSelectExamPicker');
    const inputPin = document.getElementById('modalInputExamPin');
    const errorBox = document.getElementById('modalExamPinError');
    const btnCancel = document.getElementById('btnCancelExamPinModal');

    if (!modal) return;

    if (errorBox) errorBox.classList.add('hidden');
    if (inputPin) {
        inputPin.value = '';
        inputPin.classList.remove('border-rose-500');
    }

    if (btnCancel) {
        if (canCancel) {
            btnCancel.classList.remove('hidden');
        } else {
            btnCancel.classList.add('hidden');
        }
    }

    if (selectPicker) {
        if (allExams.length === 0) {
            selectPicker.innerHTML = `<option value="">-- Memuat daftar ujian... --</option>`;
        } else {
            selectPicker.innerHTML = allExams.map(e => {
                const datePart = e.startDate ? ` (${formatDateDisplay(parseFlexibleDate(e.startDate) || e.startDate, 'short')})` : '';
                return `<option value="${e.id}">${e.instansi}${datePart}</option>`;
            }).join('');

            if (preselectedExamId && allExams.some(e => e.id === preselectedExamId)) {
                selectPicker.value = preselectedExamId;
            } else if (currentExam) {
                selectPicker.value = currentExam.id;
            } else {
                selectPicker.value = allExams[0].id;
            }
        }
    }

    modal.classList.remove('hidden');
    modal.classList.add('flex');
    if (window.lucide) window.lucide.createIcons();
    setTimeout(() => { if (inputPin) inputPin.focus(); }, 100);
};

window.closeSelectExamWithPinModal = () => {
    const modal = document.getElementById('modalSelectExamWithPin');
    if (modal) {
        modal.classList.add('hidden');
        modal.classList.remove('flex');
    }
};

window.cancelSelectExamWithPinModal = () => {
    window.closeSelectExamWithPinModal();
    // Kembalikan dropdown navbar ke ujian yang sedang aktif
    const selectNav = document.getElementById('selectActiveExamNavbar');
    if (selectNav) {
        selectNav.value = currentExam ? currentExam.id : '';
    }
};

window.verifyExamPinAndUnlock = async (e) => {
    if (e) e.preventDefault();
    const selectPicker = document.getElementById('modalSelectExamPicker');
    const inputPin = document.getElementById('modalInputExamPin');
    const errorBox = document.getElementById('modalExamPinError');
    const errorText = document.getElementById('modalExamPinErrorText');

    const selectedExamId = selectPicker ? selectPicker.value : '';
    const enteredPin = inputPin ? inputPin.value.trim() : '';

    if (!selectedExamId) {
        if (errorBox && errorText) {
            errorText.textContent = "Silakan pilih instansi pelaksanaan ujian!";
            errorBox.classList.remove('hidden');
        }
        return;
    }

    const targetExam = allExams.find(ex => ex.id === selectedExamId);
    if (!targetExam) {
        if (errorBox && errorText) {
            errorText.textContent = "Data ujian tidak ditemukan!";
            errorBox.classList.remove('hidden');
        }
        return;
    }

    const correctPin = getInstansiPin(targetExam.instansi);

    if (enteredPin.toLowerCase() === correctPin.toLowerCase()) {
        // PIN BENAR!
        saveActiveExamSession(targetExam.id, targetExam.instansi);
        window.closeSelectExamWithPinModal();
        showToast(`PIN benar! Mengambil data peserta ${targetExam.instansi}...`, 'info');
        await setActiveExam(targetExam.id);
        showToast(`Ujian ${targetExam.instansi} aktif!`, 'success');
    } else {
        // PIN SALAH!
        if (errorBox && errorText) {
            errorText.textContent = `PIN salah untuk ${targetExam.instansi}!`;
            errorBox.classList.remove('hidden');
        }
        if (inputPin) {
            inputPin.classList.add('border-rose-500');
            inputPin.select();
        }
    }
};

// ---------------------- MODAL PIN SUPER ADMIN ("141414") ----------------------

window.openSuperAdminPinPrompt = () => {
    const modal = document.getElementById('modalSuperAdminPin');
    const input = document.getElementById('inputSuperAdminPin');
    const err = document.getElementById('superAdminPinError');
    if (err) err.classList.add('hidden');
    if (input) {
        input.value = '';
        input.classList.remove('border-rose-500');
    }
    if (modal) {
        modal.classList.remove('hidden');
        modal.classList.add('flex');
        setTimeout(() => { if (input) input.focus(); }, 150);
    }
    if (window.lucide) window.lucide.createIcons();
};

window.closeSuperAdminPinModal = () => {
    const modal = document.getElementById('modalSuperAdminPin');
    if (modal) {
        modal.classList.add('hidden');
        modal.classList.remove('flex');
    }
};

window.verifySuperAdminPin = async (e) => {
    if (e) e.preventDefault();
    const input = document.getElementById('inputSuperAdminPin');
    const err = document.getElementById('superAdminPinError');
    const val = input ? input.value.trim() : '';

    if (val === '141414') {
        setSuperAdminSession(true);
        window.closeSuperAdminPinModal();
        window.closeSelectExamWithPinModal();
        window.closeModalPinAccess();

        // Aktifkan ujian yang ada jika belum aktif
        if (!currentExam && allExams.length > 0) {
            const saved = getSelectedExamId();
            const targetId = (saved && allExams.some(x => x.id === saved)) ? saved : allExams[0].id;
            await setActiveExam(targetId);
        }

        showToast("Mode Super Admin Aktif! Akses penuh dibuka tanpa batasan PIN.", "success");
    } else {
        if (err) err.classList.remove('hidden');
        if (input) {
            input.classList.add('border-rose-500');
            input.value = '';
            input.focus();
        }
    }
};

/**
 * Setup Master Instansi UI
 */
function setupMasterInstansiUI() {
    const formCustom = document.getElementById('formAddMasterInstansiCustom');
    if (formCustom) {
        formCustom.addEventListener('submit', (e) => {
            e.preventDefault();
            const inputName = document.getElementById('newMasterInstansiNameCustom');
            const selectWilker = document.getElementById('newMasterInstansiWilkerCustom');

            if (!inputName || !selectWilker) return;

            const nameValue = toTitleCase(inputName.value.trim());
            const wilkerValue = selectWilker.value;

            if (!nameValue) {
                showToast("Nama instansi tidak boleh kosong!", "warning");
                return;
            }

            const exists = masterInstansiData.some(i => i.name.toLowerCase() === nameValue.toLowerCase());
            if (exists) {
                showToast(`Instansi "${nameValue}" sudah ada!`, "warning");
                return;
            }

            masterInstansiData.push({ name: nameValue, wilker: wilkerValue });
            localStorage.setItem('master_instansi_pi', JSON.stringify(masterInstansiData));

            inputName.value = '';
            populateInstansiDropdown('selectExamInstansi');
            renderMasterInstansiTableFull();
            closeModalAddMasterInstansi();
            showToast(`Instansi "${nameValue}" berhasil ditambahkan!`, "success");
        });
    }
}

function renderMasterInstansiTableFull() {
    const tbody = document.getElementById('tbodyMasterInstansiList');
    const countPB = document.getElementById('countWilkerPB');
    const countPBD = document.getElementById('countWilkerPBD');
    const countVertikal = document.getElementById('countWilkerVertikal');

    const pbList = masterInstansiData.filter(i => i.wilker === 'Papua Barat');
    const pbdList = masterInstansiData.filter(i => i.wilker === 'Papua Barat Daya');
    const vertikalList = masterInstansiData.filter(i => i.wilker === 'Instansi Vertikal');

    if (countPB) countPB.textContent = `${pbList.length} Kabupaten/Prov`;
    if (countPBD) countPBD.textContent = `${pbdList.length} Kota/Kabupaten`;
    if (countVertikal) countVertikal.textContent = `${vertikalList.length} Instansi`;

    if (!tbody) return;

    tbody.innerHTML = masterInstansiData.map((item, idx) => `
        <tr class="hover:bg-slate-50">
            <td class="p-3 font-semibold text-slate-900">${item.name}</td>
            <td class="p-3">
                <span class="text-[11px] font-bold px-2.5 py-0.5 rounded-full ${item.wilker === 'Papua Barat Daya' ? 'bg-emerald-100 text-emerald-800' : (item.wilker === 'Instansi Vertikal' ? 'bg-amber-100 text-amber-800' : 'bg-blue-100 text-blue-800')}">
                    ${item.wilker}
                </span>
            </td>
            <td class="p-3 text-center">
                <button onclick="deleteMasterItem(${idx}, '${item.name}')" class="px-2 py-1 bg-rose-50 hover:bg-rose-100 text-rose-600 font-bold text-[11px] rounded transition">
                    Hapus
                </button>
            </td>
        </tr>
    `).join('');
}

window.deleteMasterItem = (index, name) => {
    if (confirm(`Hapus "${name}" dari master instansi?`)) {
        masterInstansiData.splice(index, 1);
        localStorage.setItem('master_instansi_pi', JSON.stringify(masterInstansiData));
        populateInstansiDropdown('selectExamInstansi');
        renderMasterInstansiTableFull();
        showToast(`"${name}" telah dihapus.`, "info");
    }
};

window.openModalAddMasterInstansi = () => {
    const modal = document.getElementById('modalMasterInstansi');
    if (modal) {
        modal.classList.remove('hidden');
        modal.classList.add('flex');
    }
};

window.closeModalAddMasterInstansi = () => {
    const modal = document.getElementById('modalMasterInstansi');
    if (modal) {
        modal.classList.add('hidden');
        modal.classList.remove('flex');
    }
};

/**
 * Toast Notification Helper
 */
function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    if (!container) return;

    const toast = document.createElement('div');
    const colorClass = type === 'success' ? 'bg-emerald-800 text-white' : 
                       type === 'error' ? 'bg-rose-800 text-white' : 
                       type === 'warning' ? 'bg-amber-800 text-white' : 'bg-slate-900 text-white';

    const iconName = type === 'success' ? 'check-circle-2' : 
                     type === 'error' ? 'alert-octagon' : 
                     type === 'warning' ? 'alert-triangle' : 'info';

    toast.className = `${colorClass} px-4 py-3 rounded-xl shadow-lg text-xs sm:text-sm font-medium flex items-center space-x-2.5 transition-all duration-300 pointer-events-auto max-w-md`;
    toast.innerHTML = `
        <i data-lucide="${iconName}" class="w-4 h-4 flex-shrink-0"></i>
        <span>${message}</span>
    `;

    container.appendChild(toast);
    if (window.lucide) window.lucide.createIcons();

    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateY(10px)';
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

// ---------------------- FIREBASE REALTIME CLOUD INTEGRATION ----------------------

function setupFirebaseIntegration() {
    onConnectionStatusChange((online) => {
        updateCloudStatusUI(online);
        if (online) {
            // Pasang realtime listener untuk ujian secara global
            listenExamsCloud((cloudExams) => {
                if (cloudExams) {
                    allExams = cloudExams;
                    renderExamSelectDropdowns();
                    renderExamListInCreateTab();
                    refreshModalSelectExamPicker();

                    // Cek jika Super Admin sudah aktif di sesi ini
                    if (checkSuperAdminSession()) {
                        isSuperAdmin = true;
                        isPinAuthorized = true;
                        const activeId = getSelectedExamId() || (allExams.length > 0 ? allExams[0].id : null);
                        if (activeId && (!currentExam || currentExam.id !== activeId)) {
                            setActiveExam(activeId);
                        }
                        return;
                    }

                    // Sinkronkan ujian aktif HANYA jika sesi PIN terotorisasi valid
                    const session = getActiveExamSession();
                    if (session && session.examId && allExams.some(e => e.id === session.examId)) {
                        if (!currentExam || currentExam.id !== session.examId) {
                            setActiveExam(session.examId);
                        } else {
                            currentExam = allExams.find(e => e.id === session.examId);
                            renderDashboardExamInfo();
                        }
                    } else if (session && !allExams.some(e => e.id === session.examId)) {
                        clearActiveExamSession();
                        setActiveExam(null);
                        window.openSelectExamWithPinModal(null, false);
                    } else if (!session && !isSuperAdmin) {
                        // Belum ada PIN terverifikasi: jangan load kandidat apapun!
                        currentExam = null;
                        currentCandidates = [];
                        renderDashboardExamInfo();
                        renderDashboardStats();
                        applyCandidateFilters();

                        const modal = document.getElementById('modalSelectExamWithPin');
                        if (modal && modal.classList.contains('hidden')) {
                            window.openSelectExamWithPinModal(null, false);
                        }
                    }
                }
            });
        }
    });

    // Inisialisasi service dengan config yang tersimpan atau di-inject
    const initialized = initFirebaseService();
    updateCloudStatusUI(initialized);

    setupFirebaseConfigForm();
}

function updateCloudStatusUI(online) {
    const isOnline = Boolean(online && isCloudActive());

    // 1. Badge di Header Navbar (Hanya muncul jika Online / terhubung)
    const badge = document.getElementById('cloudStatusBadge');
    if (badge) {
        if (isOnline) {
            badge.classList.remove('hidden');
            badge.classList.add('flex');
        } else {
            badge.classList.add('hidden');
            badge.classList.remove('flex');
        }
    }

    // 2. Badge di Modal Pengaturan Cloud
    const mDot = document.getElementById('modalCloudStatusDot');
    const mTitle = document.getElementById('modalCloudStatusTitle');
    const mDesc = document.getElementById('modalCloudStatusDesc');
    const btnDisc = document.getElementById('btnDisconnectCloud');

    if (mDot) {
        mDot.className = `w-3 h-3 rounded-full ${isOnline ? 'bg-emerald-500' : 'bg-slate-400'}`;
    }
    if (mTitle) {
        mTitle.textContent = isOnline ? 'Terhubung ke Firebase Realtime Database' : 'Menunggu Konfigurasi Cloud';
        mTitle.className = `text-xs font-bold ${isOnline ? 'text-emerald-800' : 'text-slate-800'}`;
    }
    if (mDesc) {
        const config = getFirebaseConfig();
        mDesc.textContent = isOnline 
            ? `Proyek: ${config?.projectId || 'Aktif'} (Sinkronisasi Realtime Aktif)`
            : 'Belum terhubung ke database online.';
    }
    if (btnDisc) {
        if (isOnline) {
            btnDisc.classList.remove('hidden');
        } else {
            btnDisc.classList.add('hidden');
        }
    }
}

function setupFirebaseConfigForm() {
    const form = document.getElementById('formFirebaseConfig');
    if (form) {
        form.addEventListener('submit', (e) => {
            e.preventDefault();
            const textarea = document.getElementById('inputFirebaseConfigJson');
            const inputVal = textarea ? textarea.value.trim() : '';
            if (!inputVal) {
                showToast("Silakan masukkan konfigurasi Firebase!", "warning");
                return;
            }

            try {
                let cleaned = inputVal.trim();
                const match = cleaned.match(/\{[\s\S]*\}/);
                if (match) {
                    cleaned = match[0];
                }
                const parsed = (new Function(`return ${cleaned};`))();

                if (!parsed || !parsed.apiKey || !parsed.databaseURL) {
                    showToast("Konfigurasi wajib memiliki 'apiKey' dan 'databaseURL'!", "error");
                    return;
                }

                saveFirebaseConfig(parsed);
                showToast("Konfigurasi Firebase berhasil disimpan dan terhubung!", "success");
                window.closeModalFirebaseConfig();

                // Refresh data ujian aktif
                loadInitialData();
            } catch (err) {
                console.error("Gagal parsing konfigurasi Firebase:", err);
                showToast("Format konfigurasi tidak valid! Pastikan format JSON atau objek JS benar.", "error");
            }
        });
    }
}

window.requestOpenFirebaseConfig = () => {
    if (isPinAuthorized || isSuperAdmin) {
        window.openModalFirebaseConfig();
    } else {
        pendingActionAfterPin = 'OPEN_FIREBASE_CONFIG';
        const modal = document.getElementById('modalPinAccess');
        const inputPin = document.getElementById('inputAccessPin');
        const errorMsg = document.getElementById('pinErrorMessage');
        const titleEl = document.getElementById('modalPinTitle');
        const descEl = document.getElementById('modalPinDesc');

        if (titleEl) titleEl.textContent = "Pengaturan Database Cloud";
        if (descEl) descEl.textContent = "Masukkan PIN Otorisasi Administrator untuk membuka pengaturan Database Cloud.";

        if (errorMsg) errorMsg.classList.add('hidden');
        if (inputPin) {
            inputPin.value = '';
            inputPin.classList.remove('border-rose-500');
        }
        if (modal) {
            modal.classList.remove('hidden');
            modal.classList.add('flex');
            setTimeout(() => { if (inputPin) inputPin.focus(); }, 100);
        }
        if (window.lucide) window.lucide.createIcons();
    }
};

window.openModalFirebaseConfig = () => {
    const modal = document.getElementById('modalFirebaseConfig');
    const textarea = document.getElementById('inputFirebaseConfigJson');
    const currentConf = getFirebaseConfig();

    if (textarea && currentConf) {
        textarea.value = JSON.stringify(currentConf, null, 2);
    }

    if (modal) {
        modal.classList.remove('hidden');
        modal.classList.add('flex');
    }
    if (window.lucide) window.lucide.createIcons();
};

window.closeModalFirebaseConfig = () => {
    const modal = document.getElementById('modalFirebaseConfig');
    if (modal) {
        modal.classList.add('hidden');
        modal.classList.remove('flex');
    }
};

window.disconnectCloudFirebase = () => {
    if (confirm("Yakin ingin memutuskan koneksi Firebase Cloud?")) {
        removeFirebaseConfig();
        const textarea = document.getElementById('inputFirebaseConfigJson');
        if (textarea) textarea.value = '';
        updateCloudStatusUI(false);
        showToast("Koneksi Firebase diputuskan.", "info");
    }
};

// ==================== REKAPITULASI LAPORAN KEHADIRAN (MODAL 2 TAB & COPY) ====================

/**
 * Buka modal rekapitulasi laporan kehadiran
 */
window.openModalRekapKehadiran = () => {
    const modal = document.getElementById('modalRekapKehadiran');
    if (!modal) return;

    // Tampilkan nama instansi aktif pada badge modal
    const badgeInstansi = document.getElementById('rekapModalInstansiBadge');
    if (badgeInstansi) {
        badgeInstansi.textContent = currentExam ? currentExam.instansi : 'Belum Ada Ujian Terpilih';
    }

    // Tampilkan info filter tanggal yang sedang aktif
    const filterInfo = document.getElementById('rekapModalFilterInfo');
    const uniqueDates = getSortedExamDates();
    const isAllDates = selectedDashboardDates.size === 0 || selectedDashboardDates.size === uniqueDates.length;
    if (filterInfo) {
        if (uniqueDates.length === 0) {
            filterInfo.textContent = "Filter Tanggal Aktif: Belum ada jadwal ujian";
        } else if (isAllDates) {
            filterInfo.textContent = `Filter Tanggal Aktif: Semua Tanggal Pelaksanaan (${uniqueDates.length} Hari)`;
        } else {
            filterInfo.textContent = `Filter Tanggal Aktif: ${Array.from(selectedDashboardDates).join(', ')} (${selectedDashboardDates.size} Hari Terpilih)`;
        }
    }

    // Render tabel Tab 1 dan Tab 2
    renderRekapModalTables();

    // Set default ke Tab 1
    switchRekapModalTab('kel-jabatan');

    modal.classList.remove('hidden');
    modal.classList.add('flex');

    if (window.lucide) window.lucide.createIcons();
};

/**
 * Tutup modal rekapitulasi laporan kehadiran
 */
window.closeModalRekapKehadiran = () => {
    const modal = document.getElementById('modalRekapKehadiran');
    if (modal) {
        modal.classList.add('hidden');
        modal.classList.remove('flex');
    }
};

/**
 * Beralih tab di dalam modal rekapitulasi
 */
window.switchRekapModalTab = (tabName) => {
    const btnKel = document.getElementById('tabBtnRekapKelJabatan');
    const btnSesi = document.getElementById('tabBtnRekapSesi');
    const paneKel = document.getElementById('paneRekapKelJabatan');
    const paneSesi = document.getElementById('paneRekapSesi');

    if (tabName === 'kel-jabatan') {
        if (btnKel) {
            btnKel.className = "px-4 py-2.5 text-xs sm:text-sm font-bold border-b-2 border-bkn-700 text-bkn-800 bg-white rounded-t-lg shadow-2xs transition flex items-center gap-2 cursor-pointer";
        }
        if (btnSesi) {
            btnSesi.className = "px-4 py-2.5 text-xs sm:text-sm font-bold border-b-2 border-transparent text-slate-600 hover:text-slate-900 hover:bg-white/50 rounded-t-lg transition flex items-center gap-2 cursor-pointer";
        }
        if (paneKel) paneKel.classList.remove('hidden');
        if (paneSesi) paneSesi.classList.add('hidden');
    } else {
        if (btnKel) {
            btnKel.className = "px-4 py-2.5 text-xs sm:text-sm font-bold border-b-2 border-transparent text-slate-600 hover:text-slate-900 hover:bg-white/50 rounded-t-lg transition flex items-center gap-2 cursor-pointer";
        }
        if (btnSesi) {
            btnSesi.className = "px-4 py-2.5 text-xs sm:text-sm font-bold border-b-2 border-bkn-700 text-bkn-800 bg-white rounded-t-lg shadow-2xs transition flex items-center gap-2 cursor-pointer";
        }
        if (paneKel) paneKel.classList.add('hidden');
        if (paneSesi) paneSesi.classList.remove('hidden');
    }

    if (window.lucide) window.lucide.createIcons();
};

/**
 * Menghitung dan merender tabel rekapitulasi pada kedua tab
 */
function renderRekapModalTables() {
    const tableKelContainer = document.getElementById('tableRekapKelJabatanContainer');
    const tableSesiContainer = document.getElementById('tableRekapSesiContainer');

    // 1. Saring kandidat sesuai filter tanggal aktif di Dashboard
    const uniqueDates = getSortedExamDates();
    const isAllDates = selectedDashboardDates.size === 0 || selectedDashboardDates.size === uniqueDates.length;
    const activeDates = isAllDates
        ? uniqueDates
        : uniqueDates.filter(d => selectedDashboardDates.has(d));

    const candidatesToRekap = isAllDates
        ? currentCandidates
        : currentCandidates.filter(c => selectedDashboardDates.has(c.pelaksanaan));

    if (candidatesToRekap.length === 0 || activeDates.length === 0) {
        const emptyMsg = `<div class="p-8 text-center text-slate-400 text-xs italic">Belum ada data peserta untuk tanggal yang dipilih.</div>`;
        if (tableKelContainer) tableKelContainer.innerHTML = emptyMsg;
        if (tableSesiContainer) tableSesiContainer.innerHTML = emptyMsg;
        currentRekapData = { tab1Rows: [], tab2Rows: [] };
        return;
    }

    // ==========================================
    // TAB 1: REKAP KELOMPOK JABATAN PER TANGGAL
    // Ketentuan:
    // - 5 Pilar: JPT Pratama, Administrator, Pengawas, Jab. Fungsional, Pelaksana
    // - Eselon V dimasukkan ke kelompok Pengawas
    // - Jab. Fungsional gabungan semua ahli pertama, muda, madya, mahir, penyelia, terampil
    // - Kolom: Tanggal | Hadir (5 Pilar) | Tidak Hadir (5 Pilar) | Total
    // ==========================================
    const tab1Rows = [];
    const grandTab1 = {
        hadir: { jpt: 0, admin: 0, pengawas: 0, jf: 0, pelaksana: 0, kosong: 0 },
        tidakHadir: { jpt: 0, admin: 0, pengawas: 0, jf: 0, pelaksana: 0, kosong: 0 },
        total: 0
    };

    let hasAnyKosong = false;

    activeDates.forEach(dateStr => {
        const candsDate = candidatesToRekap.filter(c => c.pelaksanaan === dateStr);
        const rowData = {
            date: dateStr,
            hadir: { jpt: 0, admin: 0, pengawas: 0, jf: 0, pelaksana: 0, kosong: 0 },
            tidakHadir: { jpt: 0, admin: 0, pengawas: 0, jf: 0, pelaksana: 0, kosong: 0 },
            total: candsDate.length
        };

        candsDate.forEach(c => {
            const cls = classifyCandidateKelompok(c);
            const isHadir = c.kehadiran === 'HADIR';
            const isTidakHadir = c.kehadiran === 'TIDAK_HADIR';
            const targetObj = isHadir ? rowData.hadir : (isTidakHadir ? rowData.tidakHadir : null);

            // Tentukan pilar
            let pilar = 'pelaksana';
            if (cls.category === 'JPT_PRATAMA') {
                pilar = 'jpt';
            } else if (cls.category === 'ADMINISTRATOR') {
                pilar = 'admin';
            } else if (cls.category === 'PENGAWAS' || cls.category === 'ESELON_V') {
                // Eselon V dimasukkan ke kelompok Pengawas sesuai permintaan
                pilar = 'pengawas';
            } else if (cls.category === 'FUNGSIONAL') {
                // Jabatan Fungsional gabungan semua jenjang
                pilar = 'jf';
            } else if (cls.category === 'PELAKSANA') {
                pilar = 'pelaksana';
            } else if (cls.category === 'KOSONG') {
                pilar = 'kosong';
                hasAnyKosong = true;
            } else {
                pilar = 'pelaksana';
            }

            if (targetObj) {
                targetObj[pilar]++;
            }
        });

        // Akumulasi grand total
        ['jpt', 'admin', 'pengawas', 'jf', 'pelaksana', 'kosong'].forEach(k => {
            grandTab1.hadir[k] += rowData.hadir[k];
            grandTab1.tidakHadir[k] += rowData.tidakHadir[k];
        });
        grandTab1.total += rowData.total;

        tab1Rows.push(rowData);
    });

    currentRekapData.tab1Rows = tab1Rows;
    currentRekapData.grandTab1 = grandTab1;
    currentRekapData.hasAnyKosong = hasAnyKosong;

    // Render HTML Tab 1
    if (tableKelContainer) {
        let theadHtml = `
            <thead class="bg-slate-900 text-white text-[10px] sm:text-[11px] uppercase font-bold sticky top-0 z-10 select-none">
                <tr>
                    <th rowspan="2" class="p-1 sm:p-2 text-center border-r border-slate-700" style="width: 11%;">Tanggal</th>
                    <th colspan="5" class="p-1 sm:p-1.5 text-center bg-emerald-800 text-emerald-100 border-r border-slate-700 tracking-wider">
                        KEHADIRAN (HADIR)
                    </th>
                    <th colspan="5" class="p-1 sm:p-1.5 text-center bg-rose-900 text-rose-100 border-r border-slate-700 tracking-wider">
                        TIDAK HADIR
                    </th>
                    <th rowspan="2" class="p-1 sm:p-2 text-center border-r border-slate-700" style="width: 8.5%;">Total</th>
                    <th rowspan="2" class="p-1 sm:p-2 text-center" style="width: 8.5%;">Aksi</th>
                </tr>
                <tr class="text-[9px] sm:text-[10px] font-semibold">
                    <!-- HADIR (5 PILAR) -->
                    <th class="p-1 sm:p-1.5 text-center bg-emerald-900/90 border-r border-slate-700 truncate" style="width: 7.2%;" title="JPT Pratama">JPT</th>
                    <th class="p-1 sm:p-1.5 text-center bg-emerald-900/90 border-r border-slate-700 truncate" style="width: 7.2%;" title="Administrator">Admin</th>
                    <th class="p-1 sm:p-1.5 text-center bg-emerald-900/90 border-r border-slate-700 truncate" style="width: 7.2%;" title="Pengawas (Termasuk Eselon V)">Pengawas</th>
                    <th class="p-1 sm:p-1.5 text-center bg-emerald-900/90 border-r border-slate-700 truncate" style="width: 7.2%;" title="Jabatan Fungsional (Semua Jenjang)">JF</th>
                    <th class="p-1 sm:p-1.5 text-center bg-emerald-900/90 border-r border-slate-700 truncate" style="width: 7.2%;" title="Pelaksana">Pelaksana</th>
                    <!-- TIDAK HADIR (5 PILAR) -->
                    <th class="p-1 sm:p-1.5 text-center bg-rose-950 border-r border-slate-700 truncate" style="width: 7.2%;" title="JPT Pratama">JPT</th>
                    <th class="p-1 sm:p-1.5 text-center bg-rose-950 border-r border-slate-700 truncate" style="width: 7.2%;" title="Administrator">Admin</th>
                    <th class="p-1 sm:p-1.5 text-center bg-rose-950 border-r border-slate-700 truncate" style="width: 7.2%;" title="Pengawas (Termasuk Eselon V)">Pengawas</th>
                    <th class="p-1 sm:p-1.5 text-center bg-rose-950 border-r border-slate-700 truncate" style="width: 7.2%;" title="Jabatan Fungsional (Semua Jenjang)">JF</th>
                    <th class="p-1 sm:p-1.5 text-center bg-rose-950 border-r border-slate-700 truncate" style="width: 7.2%;" title="Pelaksana">Pelaksana</th>
                </tr>
            </thead>
        `;

        let tbodyHtml = tab1Rows.map((r, idx) => `
            <tr class="hover:bg-slate-50 transition border-b border-slate-100 text-[10px] sm:text-[11px]">
                <td class="p-1 sm:p-1.5 text-center font-bold text-slate-800 border-r border-slate-200 bg-slate-50/50 truncate">${r.date}</td>
                <!-- Hadir -->
                <td class="p-1 sm:p-1.5 text-center text-emerald-800 font-semibold border-r border-slate-100">${r.hadir.jpt}</td>
                <td class="p-1 sm:p-1.5 text-center text-emerald-800 font-semibold border-r border-slate-100">${r.hadir.admin}</td>
                <td class="p-1 sm:p-1.5 text-center text-emerald-800 font-semibold border-r border-slate-100">${r.hadir.pengawas}</td>
                <td class="p-1 sm:p-1.5 text-center text-emerald-800 font-bold border-r border-slate-100">${r.hadir.jf}</td>
                <td class="p-1 sm:p-1.5 text-center text-emerald-800 font-semibold border-r border-slate-200">${r.hadir.pelaksana}</td>
                <!-- Tidak Hadir -->
                <td class="p-1 sm:p-1.5 text-center text-rose-800 font-semibold border-r border-slate-100">${r.tidakHadir.jpt}</td>
                <td class="p-1 sm:p-1.5 text-center text-rose-800 font-semibold border-r border-slate-100">${r.tidakHadir.admin}</td>
                <td class="p-1 sm:p-1.5 text-center text-rose-800 font-semibold border-r border-slate-100">${r.tidakHadir.pengawas}</td>
                <td class="p-1 sm:p-1.5 text-center text-rose-800 font-bold border-r border-slate-100">${r.tidakHadir.jf}</td>
                <td class="p-1 sm:p-1.5 text-center text-rose-800 font-semibold border-r border-slate-200">${r.tidakHadir.pelaksana}</td>
                <!-- Total -->
                <td class="p-1 sm:p-1.5 text-center font-bold text-slate-900 border-r border-slate-200 bg-slate-50/50">${r.total}</td>
                <!-- Aksi Copy Baris -->
                <td class="p-1 sm:p-1.5 text-center">
                    <button type="button" 
                            onclick="copyRowRekapKelJabatan(${idx})" 
                            class="w-full py-1 px-1 bg-slate-100 hover:bg-emerald-600 text-slate-700 hover:text-white rounded text-[10px] font-bold transition flex items-center justify-center gap-1 mx-auto cursor-pointer" 
                            title="Copy baris ${r.date}">
                        <i data-lucide="copy" class="w-3 h-3 flex-shrink-0"></i>
                        <span class="hidden sm:inline">Copy</span>
                    </button>
                </td>
            </tr>
        `).join('');

        // Baris Grand Total
        let tfootHtml = `
            <tfoot class="bg-slate-100 font-bold text-[10px] sm:text-[11px] border-t-2 border-slate-300 text-slate-800">
                <tr>
                    <td class="p-1 sm:p-1.5 text-center font-extrabold uppercase border-r border-slate-200">TOTAL</td>
                    <!-- Total Hadir -->
                    <td class="p-1 sm:p-1.5 text-center text-emerald-800 font-extrabold border-r border-slate-200">${grandTab1.hadir.jpt}</td>
                    <td class="p-1 sm:p-1.5 text-center text-emerald-800 font-extrabold border-r border-slate-200">${grandTab1.hadir.admin}</td>
                    <td class="p-1 sm:p-1.5 text-center text-emerald-800 font-extrabold border-r border-slate-200">${grandTab1.hadir.pengawas}</td>
                    <td class="p-1 sm:p-1.5 text-center text-emerald-800 font-extrabold border-r border-slate-200">${grandTab1.hadir.jf}</td>
                    <td class="p-1 sm:p-1.5 text-center text-emerald-800 font-extrabold border-r border-slate-200">${grandTab1.hadir.pelaksana}</td>
                    <!-- Total Tidak Hadir -->
                    <td class="p-1 sm:p-1.5 text-center text-rose-800 font-extrabold border-r border-slate-200">${grandTab1.tidakHadir.jpt}</td>
                    <td class="p-1 sm:p-1.5 text-center text-rose-800 font-extrabold border-r border-slate-200">${grandTab1.tidakHadir.admin}</td>
                    <td class="p-1 sm:p-1.5 text-center text-rose-800 font-extrabold border-r border-slate-200">${grandTab1.tidakHadir.pengawas}</td>
                    <td class="p-1 sm:p-1.5 text-center text-rose-800 font-extrabold border-r border-slate-200">${grandTab1.tidakHadir.jf}</td>
                    <td class="p-1 sm:p-1.5 text-center text-rose-800 font-extrabold border-r border-slate-200">${grandTab1.tidakHadir.pelaksana}</td>
                    <!-- Total Peserta -->
                    <td class="p-1 sm:p-1.5 text-center font-extrabold text-slate-900 border-r border-slate-200">${grandTab1.total}</td>
                    <td class="p-1 sm:p-1.5 text-center">
                        <button type="button" 
                                onclick="copyAllRekapKelJabatan()" 
                                class="w-full py-1 px-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[9px] sm:text-[10px] font-bold transition mx-auto cursor-pointer" title="Copy Semua Data">
                            Copy
                        </button>
                    </td>
                </tr>
            </tfoot>
        `;

        tableKelContainer.innerHTML = `
            <table class="w-full table-fixed text-left border-collapse">
                ${theadHtml}
                <tbody class="divide-y divide-slate-100 bg-white">
                    ${tbodyHtml}
                </tbody>
                ${tfootHtml}
            </table>
        `;
    }

    // ==========================================
    // TAB 2: REKAP KEHADIRAN PER SESI
    // Kolom: Tanggal | Sesi | Hadir | Tidak Hadir | Total | Aksi
    // ==========================================
    const tab2Rows = [];
    let grandHadirSesi = 0;
    let grandTidakHadirSesi = 0;
    let grandTotalSesi = 0;

    activeDates.forEach(dateStr => {
        const candsDate = candidatesToRekap.filter(c => c.pelaksanaan === dateStr);
        // Sesi yang ada di hari ini: 1, 2, 3 (dan bila ada null/00)
        const sessions = [1, 2, 3];
        const hasSession0 = candsDate.some(c => !c.sesi || c.sesi === 'NULL' || c.sesi === '00' || c.sesi === 0);
        if (hasSession0) sessions.unshift(0);

        sessions.forEach(s => {
            const candsSession = candsDate.filter(c => {
                if (s === 0) return !c.sesi || c.sesi === 'NULL' || c.sesi === '00' || c.sesi === 0;
                return Number(c.sesi) === s;
            });

            // Hanya tampilkan sesi jika ada pesertanya atau jika s in 1..3
            if (candsSession.length > 0 || (s >= 1 && s <= 3)) {
                const h = candsSession.filter(c => c.kehadiran === 'HADIR').length;
                const th = candsSession.filter(c => c.kehadiran === 'TIDAK_HADIR').length;
                const tot = candsSession.length;

                tab2Rows.push({
                    date: dateStr,
                    sesi: s === 0 ? 'Belum Terjadwal (00)' : `Sesi ${s}`,
                    sesiNum: s,
                    hadir: h,
                    tidakHadir: th,
                    total: tot
                });

                grandHadirSesi += h;
                grandTidakHadirSesi += th;
                grandTotalSesi += tot;
            }
        });
    });

    currentRekapData.tab2Rows = tab2Rows;
    currentRekapData.grandTab2 = {
        hadir: grandHadirSesi,
        tidakHadir: grandTidakHadirSesi,
        total: grandTotalSesi
    };

    // Render HTML Tab 2
    if (tableSesiContainer) {
        let theadSesi = `
            <thead class="bg-slate-900 text-white text-xs uppercase font-bold sticky top-0 z-10 select-none">
                <tr>
                    <th class="p-2.5 text-center border-r border-slate-700" style="width: 22%;">Tanggal</th>
                    <th class="p-2.5 text-center border-r border-slate-700" style="width: 22%;">Sesi</th>
                    <th class="p-2.5 text-center border-r border-slate-700 text-emerald-300" style="width: 18%;">Hadir</th>
                    <th class="p-2.5 text-center border-r border-slate-700 text-rose-300" style="width: 18%;">Tidak Hadir</th>
                    <th class="p-2.5 text-center border-r border-slate-700 font-bold" style="width: 10%;">Total</th>
                    <th class="p-2.5 text-center" style="width: 10%;">Aksi</th>
                </tr>
            </thead>
        `;

        let tbodySesi = tab2Rows.map((r, idx) => `
            <tr class="hover:bg-slate-50 transition border-b border-slate-100 text-xs">
                <td class="p-2.5 text-center font-bold text-slate-800 border-r border-slate-200 bg-slate-50/50 truncate">${r.date}</td>
                <td class="p-2.5 text-center font-semibold text-slate-700 border-r border-slate-200">${r.sesi}</td>
                <td class="p-2.5 text-center text-emerald-700 font-bold text-sm border-r border-slate-200">${r.hadir}</td>
                <td class="p-2.5 text-center text-rose-700 font-bold text-sm border-r border-slate-200">${r.tidakHadir}</td>
                <td class="p-2.5 text-center font-bold text-slate-900 border-r border-slate-200 bg-slate-50/50">${r.total}</td>
                <td class="p-2.5 text-center">
                    <button type="button" 
                            onclick="copyRowRekapSesi(${idx})" 
                            class="w-full py-1.5 px-2 bg-slate-100 hover:bg-emerald-600 text-slate-700 hover:text-white rounded text-[11px] font-bold transition flex items-center justify-center gap-1 mx-auto cursor-pointer" 
                            title="Copy baris ini">
                        <i data-lucide="copy" class="w-3 h-3 flex-shrink-0"></i>
                        <span>Copy</span>
                    </button>
                </td>
            </tr>
        `).join('');

        let tfootSesi = `
            <tfoot class="bg-slate-100 font-bold text-xs border-t-2 border-slate-300 text-slate-800">
                <tr>
                    <td colspan="2" class="p-2.5 text-center font-extrabold uppercase border-r border-slate-200">TOTAL KESELURUHAN</td>
                    <td class="p-2.5 text-center text-emerald-800 font-extrabold text-sm border-r border-slate-200">${grandHadirSesi}</td>
                    <td class="p-2.5 text-center text-rose-800 font-extrabold text-sm border-r border-slate-200">${grandTidakHadirSesi}</td>
                    <td class="p-2.5 text-center font-extrabold text-slate-900 border-r border-slate-200">${grandTotalSesi}</td>
                    <td class="p-2.5 text-center">
                        <button type="button" 
                                onclick="copyAllRekapSesi()" 
                                class="w-full py-1.5 px-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[10px] font-bold transition mx-auto cursor-pointer" title="Copy Semua Sesi">
                            Copy Total
                        </button>
                    </td>
                </tr>
            </tfoot>
        `;

        tableSesiContainer.innerHTML = `
            <table class="w-full table-fixed text-left border-collapse">
                ${theadSesi}
                <tbody class="divide-y divide-slate-100 bg-white">
                    ${tbodySesi}
                </tbody>
                ${tfootSesi}
            </table>
        `;
    }

    if (window.lucide) window.lucide.createIcons();
}

// ==================== FUNGSI CLIPBOARD COPY REKAPITULASI ====================

/**
 * Menyalin teks ke clipboard pengguna dan memunculkan toast
 */
function copyTextToClipboard(text, successMsg = "Data berhasil disalin ke clipboard!") {
    if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(text).then(() => {
            showToast(successMsg, "success");
        }).catch(() => {
            fallbackCopyText(text, successMsg);
        });
    } else {
        fallbackCopyText(text, successMsg);
    }
}

/**
 * Fallback metode copy clipboard untuk browser jadul / insecure context
 */
function fallbackCopyText(text, successMsg) {
    const textArea = document.createElement("textarea");
    textArea.value = text;
    textArea.style.position = "fixed";
    textArea.style.left = "-999999px";
    textArea.style.top = "-999999px";
    document.body.appendChild(textArea);
    textArea.focus();
    textArea.select();
    try {
        const successful = document.execCommand('copy');
        if (successful) {
            showToast(successMsg, "success");
        } else {
            showToast("Gagal menyalin, peramban membatasi akses clipboard.", "error");
        }
    } catch (err) {
        showToast("Gagal menyalin ke clipboard: " + err.message, "error");
    }
    document.body.removeChild(textArea);
}

/**
 * Copy 1 baris Rekap Kelompok Jabatan (TSV Excel-Ready)
 */
window.copyRowRekapKelJabatan = (rowIndex) => {
    const row = currentRekapData.tab1Rows[rowIndex];
    if (!row) return;

    const headerStr = `Tanggal\tHadir JPT Pratama\tHadir Administrator\tHadir Pengawas\tHadir Jab. Fungsional\tHadir Pelaksana\tTidak Hadir JPT Pratama\tTidak Hadir Administrator\tTidak Hadir Pengawas\tTidak Hadir Jab. Fungsional\tTidak Hadir Pelaksana\tTotal`;
    const dataStr = `${row.date}\t${row.hadir.jpt}\t${row.hadir.admin}\t${row.hadir.pengawas}\t${row.hadir.jf}\t${row.hadir.pelaksana}\t${row.tidakHadir.jpt}\t${row.tidakHadir.admin}\t${row.tidakHadir.pengawas}\t${row.tidakHadir.jf}\t${row.tidakHadir.pelaksana}\t${row.total}`;

    const textToCopy = `${headerStr}\n${dataStr}`;
    copyTextToClipboard(textToCopy, `Data rekap tanggal ${row.date} berhasil disalin! Format siap di-paste ke Excel.`);
};

/**
 * Copy seluruh data Rekap Kelompok Jabatan (Semua Baris + Total)
 */
window.copyAllRekapKelJabatan = () => {
    const rows = currentRekapData.tab1Rows || [];
    if (rows.length === 0) {
        showToast("Tidak ada data rekapitulasi untuk disalin.", "info");
        return;
    }

    const headerStr = `Tanggal\tHadir JPT Pratama\tHadir Administrator\tHadir Pengawas\tHadir Jab. Fungsional\tHadir Pelaksana\tTidak Hadir JPT Pratama\tTidak Hadir Administrator\tTidak Hadir Pengawas\tTidak Hadir Jab. Fungsional\tTidak Hadir Pelaksana\tTotal`;
    const lines = [headerStr];

    rows.forEach(r => {
        const line = `${r.date}\t${r.hadir.jpt}\t${r.hadir.admin}\t${r.hadir.pengawas}\t${r.hadir.jf}\t${r.hadir.pelaksana}\t${r.tidakHadir.jpt}\t${r.tidakHadir.admin}\t${r.tidakHadir.pengawas}\t${r.tidakHadir.jf}\t${r.tidakHadir.pelaksana}\t${r.total}`;
        lines.push(line);
    });

    // Tambahkan baris TOTAL
    const g = currentRekapData.grandTab1;
    if (g) {
        const totalLine = `TOTAL\t${g.hadir.jpt}\t${g.hadir.admin}\t${g.hadir.pengawas}\t${g.hadir.jf}\t${g.hadir.pelaksana}\t${g.tidakHadir.jpt}\t${g.tidakHadir.admin}\t${g.tidakHadir.pengawas}\t${g.tidakHadir.jf}\t${g.tidakHadir.pelaksana}\t${g.total}`;
        lines.push(totalLine);
    }

    copyTextToClipboard(lines.join('\n'), "Seluruh tabel Rekap Kelompok Jabatan berhasil disalin! Format siap di-paste ke Excel.");
};

/**
 * Copy 1 baris Rekap Sesi
 */
window.copyRowRekapSesi = (rowIndex) => {
    const row = currentRekapData.tab2Rows[rowIndex];
    if (!row) return;

    const headerStr = `Tanggal\tSesi\tHadir\tTidak Hadir\tTotal`;
    const dataStr = `${row.date}\t${row.sesi}\t${row.hadir}\t${row.tidakHadir}\t${row.total}`;

    copyTextToClipboard(`${headerStr}\n${dataStr}`, `Data ${row.sesi} (${row.date}) berhasil disalin!`);
};

/**
 * Copy seluruh tabel Rekap Sesi (Semua Baris + Total)
 */
window.copyAllRekapSesi = () => {
    const rows = currentRekapData.tab2Rows || [];
    if (rows.length === 0) {
        showToast("Tidak ada data sesi untuk disalin.", "info");
        return;
    }

    const headerStr = `Tanggal\tSesi\tHadir\tTidak Hadir\tTotal`;
    const lines = [headerStr];

    rows.forEach(r => {
        lines.push(`${r.date}\t${r.sesi}\t${r.hadir}\t${r.tidakHadir}\t${r.total}`);
    });

    const g = currentRekapData.grandTab2;
    if (g) {
        lines.push(`TOTAL KESELURUHAN\t-\t${g.hadir}\t${g.tidakHadir}\t${g.total}`);
    }

    copyTextToClipboard(lines.join('\n'), "Seluruh tabel Rekap Sesi berhasil disalin! Format siap di-paste ke Excel.");
};

