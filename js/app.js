/**
 * app.js
 * Controller Utama Aplikasi "Profiling ASN"
 * Mengintegrasikan IndexedDB, Excel Handler, Session Rules, dan Manajemen Wilayah Papua Barat & PB Daya
 */

import { masterInstansiData, toTitleCase, getInstansiPin } from '../masterInstansi.js';
import * as db from './db.js';
import { parseFlexibleDate, isFriday, getSessionTime, formatDateDisplay, getDayNameID, formatCumulativeSessionNumber, calculateCumulativeSessionNumber, convertCumulativeSessionToDaily } from './sessionRules.js';
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
    updateAttendanceInCloud,
    bulkUpdatePathsInCloud
} from './firebaseService.js';
import { parseAuditExcel, compareAuditDataWithDatabase } from './auditManager.js';

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
    setupScrollToTopButton();
    setupAuditUI();

    // Isi dropdown instansi
    populateInstansiDropdown('selectExamInstansi');

    // Cek otorisasi PIN yang tersimpan di sesi browser
    if (sessionStorage.getItem('is_admin_pin_authorized') === 'true') {
        isPinAuthorized = true;
    }

    // Jika belum ada sesi PIN yang terotorisasi, langsung munculkan pop-up PIN tanpa menunggu network
    if (!checkSuperAdminSession() && !getActiveExamSession()) {
        window.openSelectExamWithPinModal(null, false);
    }

    // Muat data ujian (hanya metadata ujian, TANPA data peserta)
    await loadInitialData();

    // Buka tab yang dipilih pengguna dari URL hash atau sessionStorage (default: 'daftar-peserta')
    const hashTab = window.location.hash.replace('#', '');
    const validTabs = ['daftar-peserta', 'dashboard', 'create-ujian', 'upload-excel', 'master-wilker'];
    const savedTab = validTabs.includes(hashTab) 
        ? hashTab 
        : (sessionStorage.getItem('active_tab') || 'daftar-peserta');
    switchTab(savedTab);

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
            let cloudSyncDebounceTimer = null;
            activeCandidatesUnsubscribe = listenCandidatesCloud(currentExam.id, (cloudCandidates) => {
                if (cloudCandidates) {
                    const prevCount = currentCandidates ? currentCandidates.length : 0;
                    currentCandidates = cloudCandidates;

                    // Jika jumlah peserta berubah (tambah/hapus): perbarui dropdown dan tabel penuh
                    if (cloudCandidates.length !== prevCount) {
                        populatePelaksanaanFilterDropdown();
                        populateSesiFilterDropdown(currentDateFilter || 'ALL');
                        populateKelJabatanFilterDropdown();
                        renderDashboardStats();
                        applyCandidateFilters();
                        return;
                    }

                    // Jika hanya perubahan status presensi: sinkronkan sel yang terlihat tanpa merender ulang seluruh DOM
                    updateVisibleAttendanceCells();
                    clearTimeout(cloudSyncDebounceTimer);
                    cloudSyncDebounceTimer = setTimeout(() => {
                        updateFloatingAttendanceBubble();
                        renderDashboardStats();
                    }, 250);
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
    if (window.updateAuditTabExamInfo) window.updateAuditTabExamInfo();
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

let isDashboardDatesInitialized = false;

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
 * Reset filter tanggal dashboard agar seluruh tanggal kembali terpilih
 */
window.resetDashboardDateFilter = () => {
    const uniqueDates = getSortedExamDates();
    selectedDashboardDates = new Set(uniqueDates);
    updateDashboardDateFilterUI();
    renderDashboardStats();
    showToast("Filter tanggal dashboard berhasil di-reset ke semua tanggal.", "info");
};

/**
 * Handler saat checkbox tanggal individual diubah
 */
window.onToggleSingleDashboardDate = (dateVal, isChecked) => {
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
    const selectedCount = selectedDashboardDates.size;
    const isAllSelected = totalDays > 0 && selectedCount === totalDays;

    if (masterCb) {
        masterCb.checked = isAllSelected;
        masterCb.indeterminate = selectedCount > 0 && selectedCount < totalDays;
    }

    if (countBadge) {
        countBadge.textContent = `${selectedCount} Hari`;
    }

    if (labelEl) {
        if (totalDays === 0) {
            labelEl.textContent = '-- Belum Ada Tanggal Ujian --';
        } else if (isAllSelected) {
            labelEl.textContent = `Semua Tanggal Pelaksanaan (${totalDays} Hari)`;
        } else if (selectedCount === 0) {
            labelEl.textContent = '0 Tanggal Terpilih (Statistik Kosong)';
        } else if (selectedCount === 1) {
            const onlyDate = Array.from(selectedDashboardDates)[0];
            labelEl.textContent = `1 Tanggal: ${onlyDate}`;
        } else {
            labelEl.textContent = `${selectedCount} Tanggal Terpilih`;
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

    // Inisialisasi awal jika belum pernah di-set
    if (!isDashboardDatesInitialized && uniqueDates.length > 0) {
        selectedDashboardDates = new Set(uniqueDates);
        isDashboardDatesInitialized = true;
    } else {
        // Hapus tanggal lama yang sudah tidak ada
        const validSelected = new Set();
        selectedDashboardDates.forEach(d => {
            if (uniqueDates.includes(d)) validSelected.add(d);
        });
        selectedDashboardDates = validSelected;
    }

    if (uniqueDates.length === 0) {
        listContainer.innerHTML = `<p class="text-slate-400 italic text-[11px] py-2 text-center">Belum ada tanggal pelaksanaan.</p>`;
        updateDashboardDateFilterUI();
        return;
    }

    let html = '';
    uniqueDates.forEach(d => {
        const isChecked = selectedDashboardDates.has(d);
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
    const isAllDatesSelected = uniqueDates.length > 0 && selectedDashboardDates.size === uniqueDates.length;
    const dashboardCandidates = selectedDashboardDates.size === 0
        ? []
        : (isAllDatesSelected ? currentCandidates : currentCandidates.filter(c => selectedDashboardDates.has(c.pelaksanaan)));

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
            // Struktur penampung statistik (menyimpan NIP kandidat per kategori untuk efisiensi memori)
            const createStatHolder = () => ({ hadir: 0, tidakHadir: 0, belum: 0, total: 0, candidateNips: [] });
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
                    const nipStr = String(c.nip || c.id || '').trim();
                    if (nipStr) statObj.candidateNips.push(nipStr);
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

            // Simpan ke window agar bisa diakses saat baris diklik untuk membuka modal
            window._activeKelJabatanStats = categories;

            // Helper render satu baris data tabel (interaktif & dapat diklik)
            const renderRow = (label, stat, options = {}) => {
                const { isParent = false, isChild = false, isWarning = false, dataKey = '' } = options;
                const pct = stat.total > 0 ? Math.round((stat.hadir / stat.total) * 100) : 0;
                
                let rowBg = 'hover:bg-amber-50/60 transition cursor-pointer group';
                if (isWarning) {
                    rowBg = 'bg-rose-50/75 hover:bg-rose-100/80 border-l-4 border-l-red-900 font-medium transition cursor-pointer group';
                } else if (isParent) {
                    rowBg = 'bg-indigo-50/40 hover:bg-indigo-100/70 font-bold border-t border-b border-indigo-100/70 transition cursor-pointer group';
                } else if (isChild) {
                    rowBg = 'bg-slate-50/50 hover:bg-slate-100/80 text-slate-600 transition cursor-pointer group';
                }

                return `
                    <tr class="${rowBg}" onclick="window.openModalDetailKelompokJabatan('${dataKey}', '${encodeURIComponent(label)}')" title="Klik untuk melihat rincian peserta (${label})">
                        <td class="p-2.5 ${isChild ? 'pl-8 text-xs font-semibold' : 'font-bold'} ${isWarning ? 'text-rose-900 flex items-center gap-1.5' : (isParent ? 'text-indigo-950 flex items-center gap-1.5' : 'text-slate-800')}">
                            ${isWarning ? '<i data-lucide="alert-circle" class="w-3.5 h-3.5 text-rose-700 inline-block flex-shrink-0"></i>' : ''}
                            ${isChild ? '<span class="text-slate-400 font-bold mr-1">↳</span>' : ''}
                            <span class="group-hover:text-indigo-900 transition-colors">${label}</span>
                            ${isParent ? '<span class="text-[10px] bg-indigo-100 text-indigo-800 font-extrabold px-1.5 py-0.5 rounded uppercase tracking-wider">Semua Jenjang</span>' : ''}
                            ${isWarning ? '<span class="text-[10px] bg-red-900 text-red-100 px-1.5 py-0.5 rounded font-bold uppercase shadow-xs">Peserta Belum Terdaftar</span>' : ''}
                            <i data-lucide="external-link" class="w-3 h-3 text-slate-400 group-hover:text-indigo-600 opacity-0 group-hover:opacity-100 transition-all ml-auto inline-block"></i>
                        </td>
                        <td class="p-2.5 text-center text-emerald-700 font-bold text-sm group-hover:bg-emerald-50/50 transition">${stat.hadir}</td>
                        <td class="p-2.5 text-center text-rose-700 font-bold text-sm group-hover:bg-rose-50/50 transition">${stat.tidakHadir}</td>
                        <td class="p-2.5 text-center text-amber-700 font-semibold group-hover:bg-amber-50/50 transition">${stat.belum}</td>
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
                rowsHtml += renderRow(categories.JPT_PRATAMA.label, categories.JPT_PRATAMA.stat, { dataKey: 'JPT_PRATAMA' });
            }

            // 2. Administrator
            if (categories.ADMINISTRATOR.stat.total > 0) {
                rowsHtml += renderRow(categories.ADMINISTRATOR.label, categories.ADMINISTRATOR.stat, { dataKey: 'ADMINISTRATOR' });
            }

            // 3. Pengawas
            if (categories.PENGAWAS.stat.total > 0) {
                rowsHtml += renderRow(categories.PENGAWAS.label, categories.PENGAWAS.stat, { dataKey: 'PENGAWAS' });
            }

            // 4. Eselon V
            if (categories.ESELON_V.stat.total > 0) {
                rowsHtml += renderRow(categories.ESELON_V.label, categories.ESELON_V.stat, { dataKey: 'ESELON_V' });
            }

            // 5. Jabatan Fungsional (Parent & Child Rows)
            if (categories.FUNGSIONAL.stat.total > 0) {
                // Render Parent Row
                rowsHtml += renderRow(categories.FUNGSIONAL.label, categories.FUNGSIONAL.stat, { isParent: true, dataKey: 'FUNGSIONAL' });
                
                // Susunan baku Child: terampil, mahir, penyelia, ahli pertama, ahli muda, ahli madya
                const standardJfOrder = ['Terampil', 'Mahir', 'Penyelia', 'Ahli Pertama', 'Ahli Muda', 'Ahli Madya', 'Fungsional Lainnya'];
                standardJfOrder.forEach(subName => {
                    const childStat = categories.FUNGSIONAL.children[subName];
                    if (childStat && childStat.total > 0) {
                        rowsHtml += renderRow(subName, childStat, { isChild: true, dataKey: `FUNGSIONAL:${subName}` });
                    }
                });
            }

            // 6. Pelaksana
            if (categories.PELAKSANA.stat.total > 0) {
                rowsHtml += renderRow(categories.PELAKSANA.label, categories.PELAKSANA.stat, { dataKey: 'PELAKSANA' });
            }

            // 7. Belum Terdata / Kosong (Peserta Belum Terdaftar) -> DI BAWAH BARIS PELAKSANA
            if (categories.KOSONG.stat.total > 0) {
                rowsHtml += renderRow(categories.KOSONG.label, categories.KOSONG.stat, { isWarning: true, dataKey: 'KOSONG' });
            }

            // 8. Kelompok Lainnya (jika ada)
            Object.keys(categories.LAINNYA).forEach(otherLabel => {
                const stat = categories.LAINNYA[otherLabel];
                if (stat.total > 0) {
                    rowsHtml += renderRow(otherLabel, stat, { dataKey: `LAINNYA:${otherLabel}` });
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

        // De-duplikasi internal file Excel (jika ada NIP yang berulang dalam file Excel yang diunggah)
        const seenNipsInExcel = new Set();
        const distinctExcelCandidates = [];
        parseResult.candidates.forEach(c => {
            const cleanNip = String(c.nip || '').replace(/['"`\s]/g, '').trim();
            if (cleanNip && !seenNipsInExcel.has(cleanNip)) {
                seenNipsInExcel.add(cleanNip);
                distinctExcelCandidates.push(c);
            }
        });

        // JIKA DATABASE SUDAH MEMILIKI DATA PESERTA:
        // Otomatis pisahkan peserta baru dan peserta yang sudah terdaftar
        if (existingCandidates && existingCandidates.length > 0) {
            const existingNipSet = new Set(
                existingCandidates.map(c => String(c.nip || '').replace(/['"`\s]/g, '').trim())
            );

            const newCandidates = [];
            const alreadyExistingCandidates = [];

            distinctExcelCandidates.forEach(c => {
                const cleanNip = String(c.nip || '').replace(/['"`\s]/g, '').trim();
                if (existingNipSet.has(cleanNip)) {
                    alreadyExistingCandidates.push(c);
                } else {
                    newCandidates.push(c);
                }
            });

            // Beri nomor urut ulang dan penanda peserta baru
            newCandidates.forEach((c, idx) => {
                c.no = idx + 1;
                c.isNewCandidate = true;
            });

            alreadyExistingCandidates.forEach((c, idx) => {
                c.no = idx + 1;
                c.isNewCandidate = false;
            });

            previewParsedData = {
                examId: targetExamId,
                candidates: newCandidates, // Tampilkan HANYA peserta baru sesuai permintaan user!
                newCandidates: newCandidates,
                alreadyExistingCandidates: alreadyExistingCandidates,
                allExcelCandidates: distinctExcelCandidates,
                isOnlyNewParticipants: true,
                existingCount: alreadyExistingCandidates.length,
                totalExcelRows: distinctExcelCandidates.length,
                summary: {
                    totalRows: newCandidates.length,
                    newCount: newCandidates.length,
                    existingCount: alreadyExistingCandidates.length,
                    skippedRows: parseResult.summary.skippedRows,
                    sesi1: newCandidates.filter(c => Number(c.sesi) === 1).length,
                    sesi2: newCandidates.filter(c => Number(c.sesi) === 2).length,
                    sesi3: newCandidates.filter(c => Number(c.sesi) === 3).length,
                    nullScheduleRows: newCandidates.filter(c => !c.sesi || c.sesi === 'NULL' || c.sesi === '00' || c.sesi === 0 || c.sesi === '0').length,
                    fridayRows: newCandidates.filter(c => c.isFriday).length
                }
            };

            renderExcelPreview(previewParsedData);

            if (newCandidates.length > 0) {
                showToast(`Ditemukan ${newCandidates.length} peserta baru (${alreadyExistingCandidates.length} peserta lama di database otomatis dilewati).`, "success");
            } else {
                showToast(`Tidak ada peserta baru. Seluruh ${alreadyExistingCandidates.length} peserta di file Excel ini sudah terdaftar di database.`, "info");
            }
            return;
        }

        // JIKA DATABASE MASIH KOSONG (Upload awal):
        distinctExcelCandidates.forEach((c, idx) => {
            c.no = idx + 1;
            c.isNewCandidate = true;
        });

        previewParsedData = {
            examId: targetExamId,
            candidates: distinctExcelCandidates,
            allExcelCandidates: distinctExcelCandidates,
            newCandidates: distinctExcelCandidates,
            alreadyExistingCandidates: [],
            isOnlyNewParticipants: false,
            existingCount: 0,
            totalExcelRows: distinctExcelCandidates.length,
            summary: {
                totalRows: distinctExcelCandidates.length,
                newCount: distinctExcelCandidates.length,
                existingCount: 0,
                skippedRows: parseResult.summary.skippedRows,
                sesi1: distinctExcelCandidates.filter(c => Number(c.sesi) === 1).length,
                sesi2: distinctExcelCandidates.filter(c => Number(c.sesi) === 2).length,
                sesi3: distinctExcelCandidates.filter(c => Number(c.sesi) === 3).length,
                nullScheduleRows: distinctExcelCandidates.filter(c => !c.sesi || c.sesi === 'NULL' || c.sesi === '00' || c.sesi === 0 || c.sesi === '0').length,
                fridayRows: distinctExcelCandidates.filter(c => c.isFriday).length
            }
        };

        renderExcelPreview(previewParsedData);
        showToast(`Berhasil membaca ${distinctExcelCandidates.length} data peserta baru!`, "success");

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
 * Render Kartu Komparasi dan Checkbox Duplikasi Data Peserta (Default: Semua Dicentang)
 */
function renderDuplicateGroupsInModal(dupAnalysis) {
    const container = document.getElementById('containerDuplicateGroups');
    if (!container) return;

    let html = '';

    // 1. DUPLIKAT DENGAN DATABASE EKSISTING (Peserta yang sudah terdaftar di database)
    if (dupAnalysis.existingDuplicates && dupAnalysis.existingDuplicates.length > 0) {
        html += `
            <div class="space-y-2">
                <div class="flex items-center justify-between px-1">
                    <span class="text-xs font-bold text-rose-800 uppercase tracking-wide flex items-center gap-1.5">
                        <i data-lucide="database" class="w-4 h-4 text-rose-600"></i>
                        Data Sudah Ada di Database (${dupAnalysis.existingDuplicates.length} Peserta)
                    </span>
                    <span class="text-[11px] text-slate-500 italic">Default: dicentang untuk mengganti data lama dengan data baru</span>
                </div>
                <div class="space-y-2.5">
        `;

        dupAnalysis.existingDuplicates.forEach((group, gIdx) => {
            const dbItem = group.existingItem;
            const incomingItem = group.incomingItems && group.incomingItems[0] ? group.incomingItems[0] : null;

            html += `
                <div class="bg-white border-2 border-rose-100 hover:border-rose-200 rounded-xl p-3 shadow-2xs transition">
                    <div class="flex flex-col sm:flex-row sm:items-center justify-between pb-2 mb-2.5 border-b border-slate-100 gap-1.5">
                        <div class="flex items-center gap-2">
                            <span class="w-5 h-5 rounded-full bg-rose-100 text-rose-800 text-[10px] font-extrabold flex items-center justify-center flex-shrink-0">
                                ${gIdx + 1}
                            </span>
                            <div>
                                <span class="font-bold text-slate-900 text-sm">${group.nama || dbItem.nama}</span>
                                <span class="font-mono text-xs text-slate-500 ml-2">NIP: ${group.nip}</span>
                            </div>
                        </div>
                        <span class="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold bg-rose-50 text-rose-700 border border-rose-200 w-fit">
                            NIP Ganda di Database
                        </span>
                    </div>

                    <!-- Grid Komparasi: Data Database Saat Ini vs Data Excel Baru -->
                    <div class="grid grid-cols-1 md:grid-cols-12 gap-2 text-xs">
                        <!-- Kolom Data Lama Database -->
                        <div class="md:col-span-6 bg-slate-50/80 border border-slate-200 rounded-lg p-2.5">
                            <div class="font-bold text-slate-700 text-[11px] uppercase mb-1.5 flex items-center gap-1">
                                <span class="w-2 h-2 rounded-full bg-slate-400"></span>
                                Data Database Saat Ini (Lama)
                            </div>
                            <div class="space-y-1 text-slate-600 text-[11px]">
                                <div><span class="text-slate-400">Kel. Jabatan:</span> <strong class="text-slate-700">${dbItem.kelJabatan || '-'}</strong></div>
                                <div><span class="text-slate-400">Jabatan:</span> ${dbItem.jabatan || '-'}</div>
                                <div><span class="text-slate-400">Unit Kerja:</span> ${dbItem.unitKerja || '-'}</div>
                                <div><span class="text-slate-400">Sesi:</span> <strong class="text-indigo-700">${dbItem.sesi ? `Sesi ${dbItem.sesi}` : 'NULL'}</strong> | <span class="text-slate-400">Status:</span> <span class="font-bold ${dbItem.kehadiran === 'HADIR' ? 'text-emerald-700' : 'text-slate-600'}">${dbItem.kehadiran || 'BELUM'}</span></div>
                            </div>
                        </div>

                        <!-- Kolom Data Baru Excel (DEFAULT: CHECKED ALL) -->
                        <div class="md:col-span-6 bg-blue-50/70 border border-blue-200 rounded-lg p-2.5 flex flex-col justify-between">
                            <div>
                                <div class="font-bold text-blue-900 text-[11px] uppercase mb-1.5 flex items-center justify-between">
                                    <span class="flex items-center gap-1">
                                        <span class="w-2 h-2 rounded-full bg-blue-500"></span>
                                        Data Baru dari File Excel
                                    </span>
                                    <span class="text-[10px] text-blue-600 font-semibold lowercase">(${incomingItem.duplicateSource || 'Excel Baru'})</span>
                                </div>
                                <div class="space-y-1 text-slate-700 text-[11px]">
                                    <div><span class="text-slate-400">Kel. Jabatan:</span> <strong class="text-blue-900">${incomingItem.kelJabatan || '-'}</strong></div>
                                    <div><span class="text-slate-400">Jabatan:</span> ${incomingItem.jabatan || '-'}</div>
                                    <div><span class="text-slate-400">Unit Kerja:</span> ${incomingItem.unitKerja || '-'}</div>
                                    <div><span class="text-slate-400">Sesi:</span> <strong class="text-indigo-800">${incomingItem.sesi ? `Sesi ${incomingItem.sesi}` : 'NULL'}</strong> ${incomingItem.pelaksanaan ? `(${incomingItem.pelaksanaan})` : ''}</div>
                                </div>
                            </div>

                            <!-- Checkbox Ganti Data (DEFAULT: CHECKED) -->
                            <div class="mt-2.5 pt-2 border-t border-blue-200/80">
                                <label class="flex items-center space-x-2 cursor-pointer select-none">
                                    <input type="checkbox" 
                                           class="dup-checkbox w-4 h-4 text-bkn-600 rounded cursor-pointer" 
                                           data-nip="${group.nip}" 
                                           data-key="${incomingItem.uniqueKey}" 
                                           data-type="existing" 
                                           checked 
                                           onchange="updateDuplicateSelectedCount()">
                                    <span class="font-bold text-xs text-blue-950">Ganti data database dengan data Excel baru ini</span>
                                </label>
                            </div>
                        </div>
                    </div>
                </div>
            `;
        });

        html += `
                </div>
            </div>
        `;
    }

    // 2. DUPLIKAT INTERNAL EXCEL (NIP muncul > 1 kali di dalam file Excel yang sama)
    if (dupAnalysis.internalDuplicates && dupAnalysis.internalDuplicates.length > 0) {
        html += `
            <div class="space-y-2 pt-3 border-t border-slate-200">
                <div class="flex items-center justify-between px-1">
                    <span class="text-xs font-bold text-blue-800 uppercase tracking-wide flex items-center gap-1.5">
                        <i data-lucide="copy" class="w-4 h-4 text-blue-600"></i>
                        Duplikasi Internal di File Excel (${dupAnalysis.internalDuplicates.length} Peserta)
                    </span>
                    <span class="text-[11px] text-slate-500 italic">Pilih 1 baris yang ingin dimasukkan ke database</span>
                </div>
                <div class="space-y-2.5">
        `;

        dupAnalysis.internalDuplicates.forEach((group, gIdx) => {
            html += `
                <div class="bg-white border-2 border-blue-100 hover:border-blue-200 rounded-xl p-3 shadow-2xs transition">
                    <div class="flex items-center justify-between pb-2 mb-2 border-b border-slate-100">
                        <div class="flex items-center gap-2">
                            <span class="w-5 h-5 rounded-full bg-blue-100 text-blue-800 text-[10px] font-extrabold flex items-center justify-center flex-shrink-0">
                                ${gIdx + 1}
                            </span>
                            <div>
                                <span class="font-bold text-slate-900 text-sm">${group.nama}</span>
                                <span class="font-mono text-xs text-slate-500 ml-2">NIP: ${group.nip}</span>
                            </div>
                        </div>
                        <span class="inline-flex items-center px-2 py-0.5 rounded-full text-[10px] font-bold bg-amber-50 text-amber-800 border border-amber-200">
                            Muncul ${group.count} Kali di Excel
                        </span>
                    </div>

                    <div class="space-y-1.5">
                        ${group.items.map((item, idx) => `
                            <label class="flex items-center justify-between p-2 rounded-lg border border-slate-200 hover:bg-slate-50 cursor-pointer text-xs">
                                <div class="flex items-center space-x-2.5">
                                    <input type="checkbox" 
                                           class="dup-checkbox w-4 h-4 text-bkn-600 rounded cursor-pointer" 
                                           data-nip="${group.nip}" 
                                           data-key="${item.uniqueKey}" 
                                           data-type="internal" 
                                           ${idx === 0 ? 'checked' : ''} 
                                           onchange="handleInternalDuplicateSingleCheck(this); updateDuplicateSelectedCount();">
                                    <div>
                                        <span class="font-bold text-slate-800">${item.duplicateSource}</span>:
                                        <span class="text-slate-600 ml-1">${item.kelJabatan || '-'} | Sesi ${item.sesi || 'NULL'} | ${item.jabatan || '-'}</span>
                                    </div>
                                </div>
                            </label>
                        `).join('')}
                    </div>
                </div>
            `;
        });

        html += `
                </div>
            </div>
        `;
    }

    container.innerHTML = html;
    if (window.lucide) window.lucide.createIcons();
}

/**
 * Hitung jumlah item duplikat yang dipilih dan update status tombol & checkbox ALL
 */
function updateDuplicateSelectedCount() {
    const checkboxes = document.querySelectorAll('.dup-checkbox');
    const checkedBoxes = document.querySelectorAll('.dup-checkbox:checked');
    const totalCount = checkboxes.length;
    const checkedCount = checkedBoxes.length;

    const labelCount = document.getElementById('labelSelectedDuplicateCount');
    if (labelCount) {
        labelCount.textContent = checkedCount;
    }

    const checkAll = document.getElementById('checkSelectAllDuplicates');
    if (checkAll) {
        checkAll.checked = (checkedCount === totalCount && totalCount > 0);
        checkAll.indeterminate = (checkedCount > 0 && checkedCount < totalCount);
    }

    const btnApply = document.getElementById('btnApplyDuplicates');
    if (btnApply) {
        const hasClean = activeDuplicateState && activeDuplicateState.dupAnalysis && activeDuplicateState.dupAnalysis.cleanCandidates.length > 0;
        if (checkedCount === 0 && !hasClean) {
            btnApply.disabled = true;
            btnApply.classList.add('opacity-50', 'cursor-not-allowed');
        } else {
            btnApply.disabled = false;
            btnApply.classList.remove('opacity-50', 'cursor-not-allowed');
        }
    }
}
window.updateDuplicateSelectedCount = updateDuplicateSelectedCount;

/**
 * Memastikan hanya 1 baris yang dipilih untuk duplikasi internal NIP yang sama
 */
window.handleInternalDuplicateSingleCheck = (currentCheckbox) => {
    if (!currentCheckbox.checked) return;
    const nip = currentCheckbox.getAttribute('data-nip');
    document.querySelectorAll(`.dup-checkbox[data-type="internal"][data-nip="${nip}"]`).forEach(cb => {
        if (cb !== currentCheckbox) {
            cb.checked = false;
        }
    });
};

/**
 * Toggle Centang / Uncheck Semua
 */
window.toggleSelectAllDuplicates = (isChecked) => {
    const seenInternalNips = new Set();
    document.querySelectorAll('.dup-checkbox').forEach(cb => {
        const type = cb.getAttribute('data-type');
        const nip = cb.getAttribute('data-nip');

        if (type === 'internal') {
            if (isChecked) {
                if (!seenInternalNips.has(nip)) {
                    cb.checked = true;
                    seenInternalNips.add(nip);
                } else {
                    cb.checked = false;
                }
            } else {
                cb.checked = false;
            }
        } else {
            cb.checked = isChecked;
        }
    });
    updateDuplicateSelectedCount();
};

/**
 * Otomatis pilih baris pertama untuk setiap NIP internal
 */
window.autoSelectFirstDuplicates = () => {
    window.toggleSelectAllDuplicates(true);
    showToast("Semua data baru dicentang untuk menggantikan data lama.", "info");
};
window.autoSelectFirstExcelDuplicates = window.autoSelectFirstDuplicates;

/**
 * Ganti Data Database dengan File Excel Baru (Centang Semua)
 */
window.autoSelectLatestExcelDuplicates = () => {
    window.toggleSelectAllDuplicates(true);
    showToast("Semua data baru dicentang untuk menggantikan data lama.", "info");
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
    const noticeContainer = document.getElementById('previewNoticeContainer');
    const btnSave = document.getElementById('btnSaveExcelToDB');

    if (!previewCard) return;

    previewCard.classList.remove('hidden');

    const s = result.summary;

    // 1. Render Summary Badges
    if (summaryContainer) {
        if (result.isOnlyNewParticipants) {
            summaryContainer.innerHTML = `
                <div class="bg-emerald-50 p-3 rounded-lg border border-emerald-200">
                    <div class="text-[10px] uppercase font-bold text-emerald-700">Peserta Baru Ditemukan</div>
                    <div class="text-xl font-bold text-emerald-900 mt-0.5">${result.newCandidates ? result.newCandidates.length : s.totalRows} Peserta</div>
                </div>
                <div class="bg-slate-100 p-3 rounded-lg border border-slate-200">
                    <div class="text-[10px] uppercase font-bold text-slate-600">Sudah Ada di Database</div>
                    <div class="text-xl font-bold text-slate-800 mt-0.5">${result.existingCount || 0} Peserta</div>
                </div>
                <div class="bg-blue-50 p-3 rounded-lg border border-blue-200">
                    <div class="text-[10px] uppercase font-bold text-blue-700">Total Baris File Excel</div>
                    <div class="text-xl font-bold text-blue-900 mt-0.5">${result.totalExcelRows || s.totalRows} Baris</div>
                </div>
                <div class="bg-indigo-50 p-3 rounded-lg border border-indigo-200">
                    <div class="text-[10px] uppercase font-bold text-indigo-700">Status Tampilan</div>
                    <div class="text-sm font-bold text-indigo-900 mt-1">Hanya Peserta Baru</div>
                </div>
            `;
        } else {
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
    }

    // 2. Render Notice Banner & Toggle Buttons
    if (noticeContainer) {
        if (result.isOnlyNewParticipants) {
            const newCount = result.newCandidates ? result.newCandidates.length : 0;
            if (newCount > 0) {
                noticeContainer.innerHTML = `
                    <div class="p-3.5 bg-emerald-50 border border-emerald-200 rounded-xl flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs">
                        <div class="flex items-start sm:items-center space-x-2.5 text-emerald-950">
                            <div class="w-7 h-7 rounded-lg bg-emerald-600 text-white flex items-center justify-center flex-shrink-0 mt-0.5 sm:mt-0">
                                <i data-lucide="user-plus" class="w-4 h-4"></i>
                            </div>
                            <div>
                                <span class="font-bold text-sm block sm:inline">Ditemukan ${newCount} Peserta Baru!</span>
                                <span class="text-emerald-800 ml-0 sm:ml-1">Tabel di bawah otomatis hanya menampilkan baris peserta baru. Data ${result.existingCount} peserta yang sudah ada di database dilewati agar tidak terganggu.</span>
                            </div>
                        </div>
                        <div class="flex items-center space-x-1.5 self-end sm:self-center flex-shrink-0">
                            <button type="button" id="btnPreviewFilterNew" onclick="switchPreviewCandidateView('new')" class="px-3 py-1.5 text-xs font-bold rounded-lg bg-emerald-700 text-white shadow-sm transition">
                                Peserta Baru (${newCount})
                            </button>
                            <button type="button" id="btnPreviewFilterAll" onclick="switchPreviewCandidateView('all')" class="px-3 py-1.5 text-xs font-medium rounded-lg bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 transition">
                                Semua di Excel (${result.totalExcelRows})
                            </button>
                        </div>
                    </div>
                `;
            } else {
                noticeContainer.innerHTML = `
                    <div class="p-3.5 bg-amber-50 border border-amber-200 rounded-xl flex items-center space-x-3 text-xs text-amber-900">
                        <div class="w-7 h-7 rounded-lg bg-amber-500 text-white flex items-center justify-center flex-shrink-0">
                            <i data-lucide="info" class="w-4 h-4"></i>
                        </div>
                        <div>
                            <span class="font-bold text-sm block">Tidak Ada Peserta Baru</span>
                            <span class="text-amber-800">Seluruh <strong>${result.existingCount} peserta</strong> di dalam file Excel ini sudah terdaftar di database. Tidak ada peserta baru yang perlu ditambahkan.</span>
                        </div>
                    </div>
                `;
            }
        } else {
            noticeContainer.innerHTML = '';
        }
    }

    // 3. Render Table Rows (Default: hanya peserta baru jika mode isOnlyNewParticipants)
    renderPreviewTableRows(result.candidates, false);

    // 4. Update Tombol Simpan
    if (btnSave) {
        if (result.isOnlyNewParticipants) {
            const newCount = result.newCandidates ? result.newCandidates.length : 0;
            if (newCount > 0) {
                btnSave.disabled = false;
                btnSave.classList.remove('opacity-50', 'cursor-not-allowed');
                btnSave.innerHTML = `<i data-lucide="user-plus" class="w-4 h-4 mr-1 inline"></i><span>Simpan ${newCount} Peserta Baru ke Database</span>`;
            } else {
                btnSave.disabled = true;
                btnSave.classList.add('opacity-50', 'cursor-not-allowed');
                btnSave.innerHTML = `<i data-lucide="check" class="w-4 h-4 mr-1 inline"></i><span>Semua Sudah Terdaftar di Database</span>`;
            }
        } else {
            btnSave.disabled = false;
            btnSave.classList.remove('opacity-50', 'cursor-not-allowed');
            btnSave.innerHTML = `<i data-lucide="database" class="w-4 h-4 mr-1 inline"></i><span>Simpan ke Database</span>`;
        }
    }

    if (window.lucide) window.lucide.createIcons();
}

/**
 * Render Baris Tabel Preview Excel
 */
function renderPreviewTableRows(list, isShowingAll = false) {
    const tbody = document.getElementById('tbodyExcelPreview');
    if (!tbody) return;

    if (!list || list.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="9" class="p-8 text-center text-slate-500 font-medium italic">
                    Tidak ada peserta baru yang ditemukan dari file ini (seluruh peserta sudah ada di database).
                </td>
            </tr>
        `;
        return;
    }

    const previewList = list.slice(0, 100);
    tbody.innerHTML = previewList.map((c, idx) => {
        const isNullSchedule = !c.pelaksanaan || c.pelaksanaan === 'NULL' || !c.sesi || c.sesi === 'NULL' || c.sesi === '00' || c.sesi === 0;
        const isFri = !isNullSchedule && c.isFriday && Number(c.sesi) === 2;
        const isNullUnit = !c.unitKerja || c.unitKerja === 'NULL' || c.unitKerja === '-';
        return `
            <tr class="${isFri ? 'bg-amber-50/60 font-medium' : 'hover:bg-slate-50'}">
                <td class="p-2.5 text-center text-slate-500">${c.no || (idx + 1)}</td>
                <td class="p-2.5 font-mono text-slate-900 whitespace-nowrap">${c.nip}</td>
                <td class="p-2.5 font-semibold text-slate-900">
                    <div class="flex items-center space-x-1.5">
                        <span>${c.nama}</span>
                        ${c.isNewCandidate ? `
                            <span class="inline-flex items-center px-1.5 py-0.2 bg-emerald-100 text-emerald-800 text-[10px] font-bold rounded border border-emerald-200">BARU</span>
                        ` : (isShowingAll ? `
                            <span class="inline-flex items-center px-1.5 py-0.2 bg-slate-100 text-slate-600 text-[10px] font-medium rounded border border-slate-200">SUDAH ADA</span>
                        ` : '')}
                    </div>
                </td>
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

    if (list.length > 100) {
        tbody.innerHTML += `
            <tr>
                <td colspan="9" class="p-3 text-center text-xs text-slate-500 bg-slate-50 font-medium italic">
                    ... dan ${list.length - 100} peserta lainnya akan dimasukkan ke database saat disimpan.
                </td>
            </tr>
        `;
    }

    if (window.lucide) window.lucide.createIcons();
}

/**
 * Switch Tampilan Preview (Hanya Peserta Baru vs Semua Baris di Excel)
 */
window.switchPreviewCandidateView = (viewType) => {
    if (!previewParsedData) return;
    const btnNew = document.getElementById('btnPreviewFilterNew');
    const btnAll = document.getElementById('btnPreviewFilterAll');

    if (viewType === 'all') {
        renderPreviewTableRows(previewParsedData.allExcelCandidates, true);
        if (btnNew && btnAll) {
            btnNew.className = "px-3 py-1.5 text-xs font-medium rounded-lg bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 transition";
            btnAll.className = "px-3 py-1.5 text-xs font-bold rounded-lg bg-bkn-700 text-white shadow-sm transition";
        }
    } else {
        renderPreviewTableRows(previewParsedData.newCandidates, false);
        if (btnNew && btnAll) {
            btnNew.className = "px-3 py-1.5 text-xs font-bold rounded-lg bg-emerald-700 text-white shadow-sm transition";
            btnAll.className = "px-3 py-1.5 text-xs font-medium rounded-lg bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 transition";
        }
    }
};

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
        showToast("Tidak ada data peserta baru untuk disimpan!", "warning");
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
        } else if (previewParsedData.isOnlyNewParticipants) {
            // Simpan HANYA peserta baru yang belum ada di database!
            // Peserta lama yang sudah ada di database sama sekali tidak dihapus / tidak diubah
            const candidatesToAdd = previewParsedData.newCandidates && previewParsedData.newCandidates.length > 0 
                ? previewParsedData.newCandidates 
                : previewParsedData.candidates;

            if (!candidatesToAdd || candidatesToAdd.length === 0) {
                showToast("Tidak ada peserta baru untuk disimpan!", "warning");
                return;
            }

            const count = await db.bulkAddCandidates(previewParsedData.examId, candidatesToAdd);
            if (isCloudActive()) {
                await bulkAddCandidatesToCloud(previewParsedData.examId, candidatesToAdd);
            }
            showToast(`Sukses! ${count} peserta baru berhasil ditambahkan ke dalam database. Data lama tetap aman.`, "success");
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
        btn.className = 'filter-sesi-btn px-2 sm:px-2.5 py-1 text-[11px] font-semibold rounded-md bg-slate-100 text-slate-700 hover:bg-slate-200 transition';
    });

    const activeBtn = session === 'ALL' 
        ? document.getElementById('btnFilterSesiAll') 
        : document.getElementById(`btnFilterSesi${session}`);

    if (activeBtn) {
        activeBtn.className = 'filter-sesi-btn px-2.5 py-1 text-[11px] font-semibold rounded-md bg-bkn-800 text-white shadow-2xs transition';
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
                btn.className = 'filter-sesi-btn px-2 sm:px-2.5 py-1 text-[11px] font-semibold rounded-md bg-slate-100 text-slate-700 hover:bg-slate-200 transition';
            });
            const btnAll = document.getElementById('btnFilterSesiAll');
            if (btnAll) btnAll.className = 'filter-sesi-btn px-2.5 py-1 text-[11px] font-semibold rounded-md bg-bkn-800 text-white shadow-2xs transition';
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
                btn.className = 'filter-sesi-btn px-2 sm:px-2.5 py-1 text-[11px] font-semibold rounded-md bg-slate-100 text-slate-700 hover:bg-slate-200 transition';
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
            btn.className = 'filter-sesi-btn px-2 sm:px-2.5 py-1 text-[11px] font-semibold rounded-md bg-slate-100 text-slate-700 hover:bg-slate-200 transition';
        });
        const btnAll = document.getElementById('btnFilterSesiAll');
        if (btnAll) btnAll.className = 'filter-sesi-btn px-2.5 py-1 text-[11px] font-semibold rounded-md bg-bkn-800 text-white shadow-2xs transition';
    } else if (sessionValue === '00' || sessionValue === 0 || sessionValue === '0') {
        setCandidateFilterSession('00');
        return;
    } else {
        currentCumulativeSessionFilter = Number(sessionValue);
        currentSessionFilter = 'ALL';
        if (inputTyping) inputTyping.value = formatCumulativeSessionNumber(sessionValue);
        document.querySelectorAll('.filter-sesi-btn').forEach(btn => {
            btn.className = 'filter-sesi-btn px-2 sm:px-2.5 py-1 text-[11px] font-semibold rounded-md bg-slate-100 text-slate-700 hover:bg-slate-200 transition';
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
        btn.className = 'filter-sesi-btn px-2 sm:px-2.5 py-1 text-[11px] font-semibold rounded-md bg-slate-100 text-slate-700 hover:bg-slate-200 transition';
    });
    const btnAll = document.getElementById('btnFilterSesiAll');
    if (btnAll) btnAll.className = 'filter-sesi-btn px-2.5 py-1 text-[11px] font-semibold rounded-md bg-bkn-800 text-white shadow-2xs transition';

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
 * Hari biasa: 3 sesi per hari (Sesi 1, 2, 3)
 * Hari Jumat: 2 sesi per hari (Sesi 1, 2)
 */
function getCumulativeSessionNumber(c, sortedDates) {
    return calculateCumulativeSessionNumber(c, sortedDates);
}

/**
 * Helper untuk merender HTML sel presensi secara instan dengan native SVG (Zero-lag, 60fps)
 */
function getAttendanceCellContent(cand) {
    const candidateKey = String(cand.nip || cand.id || '').trim();
    const rawKel = String(cand.kelJabatan || '').trim();
    const isKelEmpty = !rawKel || rawKel === '-' || rawKel === 'NULL';

    if (isKelEmpty) {
        return `
            <span class="inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-[10px] font-semibold bg-slate-100 text-slate-500 border border-slate-200 select-none" title="Presensi tidak tersedia karena peserta belum terdaftar di sistem.">
                <svg class="w-3 h-3 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke-width="2"/><path d="M4.93 4.93l14.14 14.14" stroke-width="2"/></svg>
                <span>Tidak ada</span>
            </span>
        `;
    }

    if (cand.kehadiran === 'HADIR') {
        return `
            <button onclick="toggleAttendance('${candidateKey}', 'RESET')" class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] sm:text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300 hover:bg-emerald-200 transition shadow-2xs cursor-pointer active:scale-95" title="Status: Hadir. Klik untuk ubah/batal">
                <svg class="w-3 h-3 stroke-[3]" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7"/></svg>
                <span>Hadir</span>
            </button>
        `;
    }

    if (cand.kehadiran === 'TIDAK_HADIR') {
        return `
            <button onclick="toggleAttendance('${candidateKey}', 'RESET')" class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] sm:text-[11px] font-bold bg-rose-100 text-rose-800 border border-rose-300 hover:bg-rose-200 transition shadow-2xs cursor-pointer active:scale-95" title="Status: Tidak Hadir. Klik untuk ubah/batal">
                <svg class="w-3 h-3 stroke-[3]" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12"/></svg>
                <span>Tidak Hadir</span>
            </button>
        `;
    }

    return `
        <div class="inline-flex items-center justify-center gap-1">
            <button onclick="toggleAttendance('${candidateKey}', 'HADIR')" class="p-1 rounded-md bg-emerald-50 hover:bg-emerald-600 hover:text-white text-emerald-600 border border-emerald-300 transition shadow-2xs cursor-pointer active:scale-95" title="Tandai Hadir">
                <svg class="w-3.5 h-3.5 stroke-[2.5]" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7"/></svg>
            </button>
            <button onclick="toggleAttendance('${candidateKey}', 'TIDAK_HADIR')" class="p-1 rounded-md bg-rose-50 hover:bg-rose-600 hover:text-white text-rose-600 border border-rose-300 transition shadow-2xs cursor-pointer active:scale-95" title="Tandai Tidak Hadir">
                <svg class="w-3.5 h-3.5 stroke-[2.5]" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12"/></svg>
            </button>
        </div>
    `;
}

/**
 * Format timestamp kehadiran ke tampilan detail jam menit: "10 Sept 2026, pukul 12:15 WIT"
 */
function formatAttendanceTimeDetail(cand) {
    if (!cand) return '';
    const rawTs = cand.attendanceTimestamp || (cand.kehadiran === 'HADIR' && cand.loginTime ? cand.loginTime : cand.updatedAt);
    if (!rawTs) return '';

    const d = new Date(rawTs);
    if (!isNaN(d.getTime())) {
        const day = String(d.getDate()).padStart(2, '0');
        const monthsShort = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agust', 'Sept', 'Okt', 'Nov', 'Des'];
        const mon = monthsShort[d.getMonth()];
        const y = d.getFullYear();
        const hh = String(d.getHours()).padStart(2, '0');
        const mm = String(d.getMinutes()).padStart(2, '0');
        return `${day} ${mon} ${y}, pukul ${hh}:${mm} WIT`;
    }
    return String(rawTs);
}

/**
 * Helper untuk merender isi kolom Aksi baris peserta
 * Khusus jika data peserta sudah ditandai hadir atau tidak hadir:
 * tombol edit dan hapus otomatis hilang dan kolom aksi diganti status verified ceklis lengkap dengan detail jam menit kapan disubmit.
 * Tombol edit dan hapus akan muncul kembali jika status kehadiran di-reset.
 */
function getActionCellContent(cand) {
    const isVerified = cand && (cand.kehadiran === 'HADIR' || cand.kehadiran === 'TIDAK_HADIR');
    if (isVerified) {
        const isHadir = cand.kehadiran === 'HADIR';
        const timeDetail = formatAttendanceTimeDetail(cand);
        const timeText = timeDetail ? ` pada ${timeDetail}` : '';
        const statusLabel = isHadir ? 'Kehadiran' : 'Ketidakhadiran';
        const tooltipTitle = `Terverifikasi: ${statusLabel} sudah tercatat ${timeText}.`;

        return `
            <span class="inline-flex items-center justify-center gap-1 px-2 py-0.5 rounded-full text-[10px] sm:text-[11px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-300 shadow-2xs select-none cursor-help" title="${tooltipTitle}">
                <svg class="w-3.5 h-3.5 text-emerald-600 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                    <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clip-rule="evenodd" />
                </svg>
                <span>Verified</span>
            </span>
        `;
    }

    const candidateKey = String(cand.nip || cand.id || '').trim();
    const safeNama = String(cand.nama || '').replace(/'/g, "\\'");
    return `
        <button onclick="editCandidate('${candidateKey}')" class="p-1 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded mr-0.5 cursor-pointer" title="Edit Data">
            <svg class="w-3.5 h-3.5 inline" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M15.232 5.232l3.536 3.536m-2.036-5.036a2.5 2.5 0 113.536 3.536L6.5 21.036H3v-3.572L16.732 3.732z"/></svg>
        </button>
        <button onclick="deleteSingleCandidate('${candidateKey}', '${safeNama}')" class="p-1 text-rose-600 hover:text-rose-800 hover:bg-rose-50 rounded cursor-pointer" title="Hapus Peserta">
            <svg class="w-3.5 h-3.5 inline" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/></svg>
        </button>
    `;
}

/**
 * Sinkronkan sel kehadiran yang terlihat tanpa render ulang seluruh DOM tabel
 */
function updateVisibleAttendanceCells() {
    if (!currentCandidates) return;
    currentCandidates.forEach(cand => {
        const candidateKey = String(cand.nip || cand.id || '').trim();
        const cell = document.getElementById('att-cell-' + candidateKey);
        if (cell) {
            const currentStatus = cell.getAttribute('data-status');
            const newStatus = cand.kehadiran || 'NULL';
            if (currentStatus !== newStatus) {
                cell.setAttribute('data-status', newStatus);
                cell.innerHTML = getAttendanceCellContent(cand);
            }
        }
        const actionCell = document.getElementById('action-cell-' + candidateKey);
        if (actionCell) {
            actionCell.innerHTML = getActionCellContent(cand);
        }
    });
}

/**
 * Toggle Status Kehadiran Peserta (HADIR, TIDAK_HADIR, RESET) - Ultra-Fast 0ms Latency
 */
let rapidAttendanceTimer = null;
let lastAttendanceToastTime = 0;

window.toggleAttendance = (candidateNipOrId, action) => {
    const lookup = String(candidateNipOrId || '').trim();
    const cand = currentCandidates.find(c => String(c.nip || '').trim() === lookup || String(c.id || '').trim() === lookup);
    if (!cand) {
        console.warn("Peserta tidak ditemukan untuk presensi:", candidateNipOrId);
        return;
    }

    // 1. UPDATE MEMORI LOKAL INSTAN (0ms)
    const newStatus = action === 'RESET' ? null : action;
    const nowIso = new Date().toISOString();
    cand.kehadiran = newStatus;
    cand.attendanceTimestamp = newStatus ? nowIso : null;
    cand.updatedAt = nowIso;

    // 2. UPDATE DOM SEL INI SAJA SECARA LOKAL (INSTAN, TANPA RENDER ULANG TABEL!)
    const candidateKey = String(cand.nip || cand.id || '').trim();
    const cell = document.getElementById('att-cell-' + candidateKey);
    if (cell) {
        cell.setAttribute('data-status', newStatus || 'NULL');
        cell.innerHTML = getAttendanceCellContent(cand);
    }
    const actionCell = document.getElementById('action-cell-' + candidateKey);
    if (actionCell) {
        actionCell.innerHTML = getActionCellContent(cand);
    }

    // 3. UPDATE FLOATING ATTENDANCE BUBBLE SECARA INSTAN
    updateFloatingAttendanceBubble();

    // 4. TOAST RINGAN TANPA MENUMPUK/MEMBEBANI BROWSER
    const now = Date.now();
    if (now - lastAttendanceToastTime > 600) {
        const statusText = cand.kehadiran === 'HADIR' ? 'Hadir' : (cand.kehadiran === 'TIDAK_HADIR' ? 'Tidak Hadir' : 'Direset');
        showToast(`Status ${cand.nama}: ${statusText}`, cand.kehadiran === 'HADIR' ? 'success' : (cand.kehadiran === 'TIDAK_HADIR' ? 'error' : 'info'));
        lastAttendanceToastTime = now;
    }

    // 5. SINKRONKAN KE DATABASE CLOUD SECARA BACKGROUND (FIRE-AND-FORGET, NON-BLOCKING!)
    if (isCloudActive() && currentExam) {
        updateAttendanceInCloud(currentExam.id, candidateKey, cand.kehadiran).catch(err => {
            console.error("Gagal sinkron presensi ke cloud:", err);
        });
    }

    // 6. JIKA KOLOM SORTING AKTIF ADALAH 'KEHADIRAN':
    // Debounce re-sort 800ms agar posisi baris tidak meloncat saat user sedang menandai cepat
    clearTimeout(rapidAttendanceTimer);
    rapidAttendanceTimer = setTimeout(() => {
        if (currentSortColumn === 'kehadiran') {
            applyCandidateFilters();
        }
    }, 800);
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

function applyCandidateFilters(resetPage = true) {
    if (resetPage) {
        candidateCurrentPage = 1;
    }
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
    updateFloatingAttendanceBubble();
}

/**
 * Update data pada Floating Attendance Bubble (Hadir, Tidak Hadir, Belum Presensi, Total & Bar Persentase)
 */
function updateFloatingAttendanceBubble() {
    const bubble = document.getElementById('floatingAttendanceBubble');
    if (!bubble) return;

    if (!currentExam || currentCandidates.length === 0) {
        bubble.classList.add('hidden');
        return;
    }
    bubble.classList.remove('hidden');

    const total = filteredCandidates.length;
    const hadir = filteredCandidates.filter(c => c.kehadiran === 'HADIR').length;
    const tidakHadir = filteredCandidates.filter(c => c.kehadiran === 'TIDAK_HADIR').length;
    const belum = Math.max(0, total - hadir - tidakHadir);

    const formatCommaPct = (num) => {
        return (num || 0).toFixed(2).replace('.', ',');
    };

    const numPctHadir = total > 0 ? ((hadir / total) * 100) : 0;
    const numPctTidakHadir = total > 0 ? ((tidakHadir / total) * 100) : 0;
    const numPctBelum = total > 0 ? ((belum / total) * 100) : 0;

    const pctHadir = formatCommaPct(numPctHadir);
    const pctTidakHadir = formatCommaPct(numPctTidakHadir);
    const pctBelum = formatCommaPct(numPctBelum);

    const elHadir = document.getElementById('bubbleCountHadir');
    const elPctHadir = document.getElementById('bubblePctHadir');
    const elTidakHadir = document.getElementById('bubbleCountTidakHadir');
    const elPctTidakHadir = document.getElementById('bubblePctTidakHadir');
    const elBelum = document.getElementById('bubbleCountBelum');
    const elPctBelum = document.getElementById('bubblePctBelum');
    const elTotal = document.getElementById('bubbleCountTotal');
    const elPctTotal = document.getElementById('bubblePctTotal');
    const elMiniCount = document.getElementById('bubbleMiniCount');

    if (elHadir) elHadir.textContent = hadir;
    if (elPctHadir) elPctHadir.textContent = `(${pctHadir}%)`;
    if (elTidakHadir) elTidakHadir.textContent = tidakHadir;
    if (elPctTidakHadir) elPctTidakHadir.textContent = `(${pctTidakHadir}%)`;
    if (elBelum) elBelum.textContent = belum;
    if (elPctBelum) elPctBelum.textContent = `(${pctBelum}%)`;
    if (elTotal) elTotal.textContent = total;
    if (elPctTotal) elPctTotal.textContent = total > 0 ? '(100,00%)' : '(0,00%)';
    if (elMiniCount) elMiniCount.textContent = hadir;

    // Multi-segmented vertical bar (proporsi memanjang vertikal & lebih lebar)
    const barHadir = document.getElementById('bubbleBarHadir');
    const barTidakHadir = document.getElementById('bubbleBarTidakHadir');
    const barBelum = document.getElementById('bubbleBarBelum');

    if (barHadir) {
        barHadir.style.height = `${numPctHadir.toFixed(2)}%`;
        barHadir.title = `Hadir: ${hadir} (${pctHadir}%)`;
        const txt = barHadir.querySelector('.bubbleBarText');
        if (txt) {
            txt.textContent = numPctHadir >= 16 ? `${pctHadir}%` : (numPctHadir >= 8 ? `${numPctHadir.toFixed(1).replace('.', ',')}%` : '');
        }
    }
    if (barTidakHadir) {
        barTidakHadir.style.height = `${numPctTidakHadir.toFixed(2)}%`;
        barTidakHadir.title = `Tidak Hadir: ${tidakHadir} (${pctTidakHadir}%)`;
        const txt = barTidakHadir.querySelector('.bubbleBarText');
        if (txt) {
            txt.textContent = numPctTidakHadir >= 16 ? `${pctTidakHadir}%` : (numPctTidakHadir >= 8 ? `${numPctTidakHadir.toFixed(1).replace('.', ',')}%` : '');
        }
    }
    if (barBelum) {
        barBelum.style.height = total === 0 ? '100%' : `${numPctBelum.toFixed(2)}%`;
        barBelum.title = `Belum Presensi: ${belum} (${pctBelum}%)`;
        const txt = barBelum.querySelector('.bubbleBarText');
        if (txt) {
            txt.textContent = total > 0 && numPctBelum >= 16 ? `${pctBelum}%` : (total > 0 && numPctBelum >= 8 ? `${numPctBelum.toFixed(1).replace('.', ',')}%` : '');
        }
    }

    // Micro-legend di bawah bar (2 angka di belakang koma)
    const legHadir = document.getElementById('bubbleLegendHadir');
    const legTidakHadir = document.getElementById('bubbleLegendTidakHadir');
    const legBelum = document.getElementById('bubbleLegendBelum');

    if (legHadir) legHadir.textContent = `${pctHadir}%`;
    if (legTidakHadir) legTidakHadir.textContent = `${pctTidakHadir}%`;
    if (legBelum) legBelum.textContent = `${pctBelum}%`;
}
window.updateFloatingAttendanceBubble = updateFloatingAttendanceBubble;

/**
 * Minimize / Expand Floating Attendance Bubble dengan transisi mulus
 */
window.toggleFloatingBubbleCollapse = (isCollapse) => {
    const expanded = document.getElementById('bubbleExpandedView');
    const collapsed = document.getElementById('bubbleCollapsedView');
    if (!expanded || !collapsed) return;

    if (isCollapse) {
        // Animasi keluar untuk expanded (slide ke kiri & fade out)
        expanded.classList.remove('opacity-100', 'translate-x-0');
        expanded.classList.add('opacity-0', '-translate-x-full', 'pointer-events-none');

        setTimeout(() => {
            expanded.classList.add('hidden');
            collapsed.classList.remove('hidden');
            
            requestAnimationFrame(() => {
                collapsed.classList.remove('opacity-0', '-translate-x-full');
                collapsed.classList.add('opacity-100', 'translate-x-0');
            });
            if (window.lucide) window.lucide.createIcons();
        }, 180);
    } else {
        // Animasi keluar untuk collapsed
        collapsed.classList.remove('opacity-100', 'translate-x-0');
        collapsed.classList.add('opacity-0', '-translate-x-full');

        setTimeout(() => {
            collapsed.classList.add('hidden');
            expanded.classList.remove('hidden');
            
            requestAnimationFrame(() => {
                expanded.classList.remove('opacity-0', '-translate-x-full', 'pointer-events-none');
                expanded.classList.add('opacity-100', 'translate-x-0');
            });
            if (window.lucide) window.lucide.createIcons();
        }, 150);
    }
};

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

let candidateCurrentPage = 1;
let candidatePageSize = 50;
let isVirtualScrollActive = false;
let virtualScrollTicking = false;
let lastVirtualStartIndex = -1;
let lastVirtualEndIndex = -1;

const VIRTUAL_ROW_HEIGHT = 41;
const VIRTUAL_BUFFER = 15;
const VIRTUAL_WINDOW = 45;

window.changeCandidatePageSize = (size) => {
    candidatePageSize = size === 'ALL' ? 'ALL' : Number(size);
    candidateCurrentPage = 1;
    renderCandidateListTable();
};

window.changeCandidatePage = (page) => {
    candidateCurrentPage = Math.max(1, Number(page) || 1);
    renderCandidateListTable();
    const tableEl = document.getElementById('tableCandidateContainer') || document.getElementById('tbodyCandidateList');
    if (tableEl) {
        tableEl.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
};

function onCandidateVirtualScroll() {
    if (candidatePageSize !== 'ALL' || filteredCandidates.length <= 60) return;
    if (virtualScrollTicking) return;

    virtualScrollTicking = true;
    requestAnimationFrame(() => {
        virtualScrollTicking = false;
        renderVirtualCandidateSlice();
    });
}

function renderCandidateRowHtml(c, globalIdx, sortedDates) {
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

    const candidateKey = String(c.nip || c.id || '').trim();
    const safeNama = String(c.nama || '').replace(/'/g, "\\'");

    return `
        <tr id="cand-row-${candidateKey}" class="${rowBgClass} transition text-[11px] sm:text-xs">
            <td class="p-2 text-center text-slate-500 font-medium">${globalIdx}</td>
            <td id="att-cell-${candidateKey}" data-status="${c.kehadiran || 'NULL'}" class="p-1.5 text-center whitespace-nowrap">
                ${getAttendanceCellContent(c)}
            </td>
            <td class="p-2 font-mono font-medium text-slate-900 truncate" title="${c.nip}">${c.nip}</td>
            <td class="p-2 font-bold text-slate-900 break-words line-clamp-2" title="${c.nama}">${c.nama}</td>
            <td class="p-2 truncate" title="${isKelEmpty ? 'Peserta Belum Terdaftar' : c.kelJabatan}">
                ${isKelEmpty ? `
                    <span class="inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-[10px] font-bold bg-red-900 text-red-100 border border-red-950 shadow-xs whitespace-nowrap">
                        <svg class="w-3 h-3 text-red-200 inline" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
                        <span>Belum Terdaftar</span>
                    </span>
                ` : `
                    <span class="font-semibold text-blue-800 bg-blue-50/60 px-1.5 py-0.5 rounded border border-blue-200/50 text-[10px] sm:text-[11px] truncate block">${c.kelJabatan}</span>
                `}
            </td>
            <td class="p-2 text-slate-600 truncate" title="${c.unitKerja || '-'}">
                ${isNullUnit ? `
                    <span class="px-1.5 py-0.5 rounded text-[9px] font-bold bg-slate-100 text-slate-500 border border-slate-200">NULL</span>
                ` : (c.unitKerja || '-')}
            </td>
            <td class="p-2 text-slate-600 truncate" title="${c.jabatan || '-'}">${c.jabatan || '-'}</td>
            <td class="p-2 text-center whitespace-nowrap">
                ${isNullDate ? `
                    <span class="px-1.5 py-0.5 rounded text-[9px] font-bold bg-slate-100 text-slate-500 border border-slate-200">NULL</span>
                ` : `
                    <span class="font-semibold text-slate-800 text-[11px]">${c.pelaksanaan}</span>
                    ${c.isFriday ? '<span class="text-[9px] bg-amber-100 text-amber-800 font-bold px-1 rounded ml-0.5">Jumat</span>' : ''}
                `}
            </td>
            <td class="p-2 text-center whitespace-nowrap">
                ${isSesi00 ? `
                    <div class="inline-flex items-center justify-center gap-1">
                        <span class="px-1.5 py-0.5 rounded text-[10px] font-bold bg-slate-100 text-slate-600 border border-slate-300">
                            S00
                        </span>
                        <span class="px-1 py-0.5 rounded text-[10px] font-extrabold bg-slate-400 text-white shadow-xs" title="Sesi Kumulatif: 00">
                            00
                        </span>
                    </div>
                ` : `
                    <div class="inline-flex items-center justify-center gap-1">
                        <span class="px-1.5 py-0.5 rounded text-[10px] font-bold ${sesiColorBadge}">
                            S${c.sesi}
                        </span>
                        <span class="px-1 py-0.5 rounded text-[10px] font-extrabold bg-slate-800 text-white shadow-xs border border-slate-700" title="Sesi Kumulatif: ${cumSesiFormatted}">
                            ${cumSesiFormatted}
                        </span>
                    </div>
                `}
            </td>
            <!-- Kolom Waktu (WIT) di-hide -->
            <td class="p-2 whitespace-nowrap hidden ${isFriSession2 ? 'font-bold text-amber-800' : 'text-slate-700 font-medium'}">
                ${(!c.waktu || c.waktu === 'NULL' || c.waktu === '-') ? '<span class="px-1.5 py-0.5 rounded text-[9px] font-bold bg-slate-100 text-slate-500 border border-slate-200">NULL</span>' : c.waktu}
            </td>
            <td id="action-cell-${candidateKey}" class="p-2 text-center whitespace-nowrap">
                ${getActionCellContent(c)}
            </td>
        </tr>
    `;
}

function renderVirtualCandidateSlice() {
    const tbody = document.getElementById('tbodyCandidateList');
    if (!tbody || candidatePageSize !== 'ALL') return;

    const total = filteredCandidates.length;
    if (total <= 60) return;

    const rect = tbody.getBoundingClientRect();
    const scrolledPast = Math.max(0, -rect.top);
    const firstVisible = Math.floor(scrolledPast / VIRTUAL_ROW_HEIGHT);
    const startIndex = Math.max(0, firstVisible - VIRTUAL_BUFFER);
    const endIndex = Math.min(total, firstVisible + VIRTUAL_WINDOW + VIRTUAL_BUFFER);

    if (lastVirtualStartIndex !== -1 && Math.abs(startIndex - lastVirtualStartIndex) < 4 && endIndex === lastVirtualEndIndex) {
        return;
    }

    lastVirtualStartIndex = startIndex;
    lastVirtualEndIndex = endIndex;

    const topSpacer = startIndex * VIRTUAL_ROW_HEIGHT;
    const bottomSpacer = (total - endIndex) * VIRTUAL_ROW_HEIGHT;
    const sortedDates = getSortedExamDates();

    const rowsHtml = filteredCandidates.slice(startIndex, endIndex).map((c, i) => {
        return renderCandidateRowHtml(c, startIndex + i + 1, sortedDates);
    }).join('');

    let fullHtml = '';
    if (topSpacer > 0) {
        fullHtml += `<tr id="virtual-spacer-top" style="height:${topSpacer}px;border:none;"><td colspan="10" style="padding:0;height:${topSpacer}px;border:none;line-height:0;font-size:0;"></td></tr>`;
    }
    fullHtml += rowsHtml;
    if (bottomSpacer > 0) {
        fullHtml += `<tr id="virtual-spacer-bottom" style="height:${bottomSpacer}px;border:none;"><td colspan="10" style="padding:0;height:${bottomSpacer}px;border:none;line-height:0;font-size:0;"></td></tr>`;
    }

    tbody.innerHTML = fullHtml;
}

function renderCandidateListTable() {
    const tbody = document.getElementById('tbodyCandidateList');
    const countBadge = document.getElementById('countTableVisible');
    const paginationInfo = document.getElementById('tablePaginationInfo');
    const paginationControls = document.getElementById('candidatePaginationControls');

    if (!tbody) return;

    if (countBadge) countBadge.textContent = `${filteredCandidates.length} Data`;

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
        if (isVirtualScrollActive) {
            window.removeEventListener('scroll', onCandidateVirtualScroll);
            isVirtualScrollActive = false;
        }
        tbody.innerHTML = `
            <tr>
                <td colspan="10" class="p-8 text-center text-slate-400">
                    <i data-lucide="lock" class="w-8 h-8 mx-auto mb-2 text-slate-300"></i>
                    <p class="font-bold text-slate-700 text-sm">Pilih Instansi Ujian & Masukkan PIN</p>
                    <p class="text-xs text-slate-500 mt-1">Data peserta hanya akan dimuat setelah instansi dipilih dan PIN berhasil diverifikasi.</p>
                </td>
            </tr>
        `;
        if (paginationInfo) paginationInfo.textContent = `Menampilkan 0 dari 0 peserta`;
        if (paginationControls) paginationControls.innerHTML = '';
        if (window.lucide) window.lucide.createIcons();
        return;
    }

    if (currentCandidates.length === 0) {
        if (isVirtualScrollActive) {
            window.removeEventListener('scroll', onCandidateVirtualScroll);
            isVirtualScrollActive = false;
        }
        tbody.innerHTML = `
            <tr>
                <td colspan="10" class="p-8 text-center text-slate-400">
                    <i data-lucide="inbox" class="w-8 h-8 mx-auto mb-2 text-slate-300"></i>
                    <p class="font-bold text-slate-700 text-sm">Belum Ada Data Peserta</p>
                    <p class="text-xs text-slate-500 mt-1">Belum ada peserta yang diunggah untuk instansi <strong>${currentExam.instansi}</strong>.</p>
                </td>
            </tr>
        `;
        if (paginationInfo) paginationInfo.textContent = `Menampilkan 0 dari 0 peserta`;
        if (paginationControls) paginationControls.innerHTML = '';
        if (window.lucide) window.lucide.createIcons();
        return;
    }

    if (filteredCandidates.length === 0) {
        if (isVirtualScrollActive) {
            window.removeEventListener('scroll', onCandidateVirtualScroll);
            isVirtualScrollActive = false;
        }
        tbody.innerHTML = `
            <tr>
                <td colspan="10" class="p-8 text-center text-slate-400">
                    <i data-lucide="search-x" class="w-8 h-8 mx-auto mb-2 text-slate-300"></i>
                    <p class="font-bold text-slate-700 text-sm">Tidak Ada Data Peserta</p>
                    <p class="text-xs text-slate-500 mt-1">Tidak ada peserta yang cocok dengan kriteria filter pencarian saat ini.</p>
                </td>
            </tr>
        `;
        if (paginationInfo) paginationInfo.textContent = `Menampilkan 0 dari 0 peserta`;
        if (paginationControls) paginationControls.innerHTML = '';
        if (window.lucide) window.lucide.createIcons();
        return;
    }

    const totalFiltered = filteredCandidates.length;
    const isAll = candidatePageSize === 'ALL';
    const sortedDates = getSortedExamDates();

    // =========================================================================
    // KASUS 1: MODE "SEMUA" DENGAN VIRTUAL SCROLLING (HEMAT MEMORI & BEBAS FREEZE)
    // =========================================================================
    if (isAll) {
        if (totalFiltered > 60) {
            if (!isVirtualScrollActive) {
                window.addEventListener('scroll', onCandidateVirtualScroll, { passive: true });
                isVirtualScrollActive = true;
            }
            lastVirtualStartIndex = -1;
            lastVirtualEndIndex = -1;
            renderVirtualCandidateSlice();
        } else {
            if (isVirtualScrollActive) {
                window.removeEventListener('scroll', onCandidateVirtualScroll);
                isVirtualScrollActive = false;
            }
            tbody.innerHTML = filteredCandidates.map((c, idx) => {
                return renderCandidateRowHtml(c, idx + 1, sortedDates);
            }).join('');
        }

        if (paginationInfo) {
            const totalStr = totalFiltered.toLocaleString('id-ID');
            paginationInfo.innerHTML = `Menampilkan <span class="font-bold text-slate-800">Semua (${totalStr})</span> peserta`;
        }

        if (paginationControls) {
            paginationControls.innerHTML = `
                <div class="flex items-center gap-1.5 bg-emerald-50 border border-emerald-200 px-3 py-1 rounded-lg text-emerald-800 text-xs font-semibold shadow-2xs">
                    <span class="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>
                    <span>Virtual Scroll Aktif (Hemat RAM)</span>
                </div>
            `;
        }
        return;
    }

    // =========================================================================
    // KASUS 2: MODE PAGINASI (50, 100, 250)
    // =========================================================================
    if (isVirtualScrollActive) {
        window.removeEventListener('scroll', onCandidateVirtualScroll);
        isVirtualScrollActive = false;
    }

    const effectivePageSize = Number(candidatePageSize) || 50;
    const totalPages = Math.max(1, Math.ceil(totalFiltered / effectivePageSize));

    if (candidateCurrentPage > totalPages) candidateCurrentPage = totalPages;
    if (candidateCurrentPage < 1) candidateCurrentPage = 1;

    const startIndex = (candidateCurrentPage - 1) * effectivePageSize;
    const endIndex = Math.min(startIndex + effectivePageSize, totalFiltered);
    const displayedCandidates = filteredCandidates.slice(startIndex, endIndex);

    // Update info footer
    if (paginationInfo) {
        const startStr = (startIndex + 1).toLocaleString('id-ID');
        const endStr = endIndex.toLocaleString('id-ID');
        const totalStr = totalFiltered.toLocaleString('id-ID');
        paginationInfo.innerHTML = `Menampilkan <span class="font-bold text-slate-800">${startStr} - ${endStr}</span> dari <span class="font-bold text-slate-800">${totalStr}</span> peserta`;
    }

    // Render kontrol navigasi halaman
    if (paginationControls) {
        if (totalPages <= 1) {
            paginationControls.innerHTML = `<span class="text-slate-400 text-xs px-2 py-1 font-medium bg-slate-100/70 rounded">Semua data ditampilkan</span>`;
        } else {
            let optionsHtml = '';
            for (let p = 1; p <= totalPages; p++) {
                optionsHtml += `<option value="${p}" ${p === candidateCurrentPage ? 'selected' : ''}>Hal ${p} / ${totalPages}</option>`;
            }

            const isFirst = candidateCurrentPage === 1;
            const isLast = candidateCurrentPage === totalPages;

            paginationControls.innerHTML = `
                <div class="flex items-center gap-1 bg-white p-1 rounded-lg border border-slate-200 shadow-2xs">
                    <button onclick="changeCandidatePage(1)" ${isFirst ? 'disabled' : ''} class="px-2 py-1 text-[11px] font-bold rounded ${isFirst ? 'text-slate-300 cursor-not-allowed' : 'text-slate-700 hover:bg-slate-100 cursor-pointer active:scale-95'}" title="Halaman Pertama">
                        ⇤
                    </button>
                    <button onclick="changeCandidatePage(${candidateCurrentPage - 1})" ${isFirst ? 'disabled' : ''} class="px-2 py-1 text-[11px] font-bold rounded ${isFirst ? 'text-slate-300 cursor-not-allowed' : 'text-slate-700 hover:bg-slate-100 cursor-pointer active:scale-95'}" title="Halaman Sebelumnya">
                        ‹ Prev
                    </button>
                    <select onchange="changeCandidatePage(Number(this.value))" class="text-[11px] font-bold bg-slate-50 border border-slate-300 rounded px-2 py-1 text-slate-800 focus:ring-1 focus:ring-bkn-600 outline-none cursor-pointer">
                        ${optionsHtml}
                    </select>
                    <button onclick="changeCandidatePage(${candidateCurrentPage + 1})" ${isLast ? 'disabled' : ''} class="px-2 py-1 text-[11px] font-bold rounded ${isLast ? 'text-slate-300 cursor-not-allowed' : 'text-slate-700 hover:bg-slate-100 cursor-pointer active:scale-95'}" title="Halaman Berikutnya">
                        Next ›
                    </button>
                    <button onclick="changeCandidatePage(${totalPages})" ${isLast ? 'disabled' : ''} class="px-2 py-1 text-[11px] font-bold rounded ${isLast ? 'text-slate-300 cursor-not-allowed' : 'text-slate-700 hover:bg-slate-100 cursor-pointer active:scale-95'}" title="Halaman Terakhir">
                        ⇥
                    </button>
                </div>
            `;
        }
    }

    // Render baris data tabel hanya untuk halaman aktif dengan native SVG
    tbody.innerHTML = displayedCandidates.map((c, localIdx) => {
        return renderCandidateRowHtml(c, startIndex + localIdx + 1, sortedDates);
    }).join('');
}

window.deleteSingleCandidate = async (candidateIdOrNip, name) => {
    const safeKey = String(candidateIdOrNip || '').trim();
    const cand = currentCandidates.find(c => String(c.nip || '').trim() === safeKey || String(c.id || '').trim() === safeKey);
    if (cand && (cand.kehadiran === 'HADIR' || cand.kehadiran === 'TIDAK_HADIR')) {
        showToast("Data peserta sudah terverifikasi (Hadir/Tidak Hadir) dan tidak dapat dihapus. Reset status kehadiran terlebih dahulu jika ingin menghapus.", "warning");
        return;
    }

    if (confirm(`Hapus peserta "${name}" dari jadwal ujian?`)) {
        try {
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

/**
 * Mendapatkan daftar tanggal riil pelaksanaan ujian yang tersedia (startDate s.d. endDate & data peserta)
 */
function getAvailableExamDates() {
    const dateSet = new Set();

    // 1. Dari jadwal create ujian pertama kali (startDate s.d. endDate)
    if (currentExam && currentExam.startDate) {
        const start = parseFlexibleDate(currentExam.startDate);
        const end = currentExam.endDate ? parseFlexibleDate(currentExam.endDate) : start;
        if (start && end) {
            const cur = new Date(start.getTime());
            let limit = 0;
            while (cur <= end && limit < 60) {
                dateSet.add(formatDateDisplay(cur, 'short'));
                cur.setDate(cur.getDate() + 1);
                limit++;
            }
        } else if (start) {
            dateSet.add(formatDateDisplay(start, 'short'));
        }
    }

    // 2. Tambahkan tanggal yang sudah ada di data peserta
    currentCandidates.forEach(c => {
        if (c.pelaksanaan && c.pelaksanaan !== 'NULL' && c.pelaksanaan !== '-') {
            dateSet.add(c.pelaksanaan);
        }
    });

    // 3. Urutkan secara kronologis
    const sorted = Array.from(dateSet).sort((a, b) => {
        const da = parseFlexibleDate(a);
        const db = parseFlexibleDate(b);
        return (da ? da.getTime() : 0) - (db ? db.getTime() : 0);
    });

    return sorted;
}

/**
 * Mengisi dropdown pilihan tanggal pelaksanaan pada modal tambah/edit peserta
 */
function populateManualCandidateDateOptions(selectedDate = '') {
    const select = document.getElementById('inputManualPelaksanaan');
    if (!select) return;

    const dates = getAvailableExamDates();
    if (dates.length === 0) {
        select.innerHTML = `<option value="">-- Belum ada tanggal ujian --</option>`;
        return;
    }

    let html = `<option value="">-- Pilih Tanggal Pelaksanaan --</option>`;
    dates.forEach(d => {
        const fri = isFriday(d);
        const label = `${d}${fri ? ' (Jumat)' : ''}`;
        html += `<option value="${d}">${label}</option>`;
    });

    select.innerHTML = html;

    // Set nilai terpilih
    if (selectedDate && dates.includes(selectedDate)) {
        select.value = selectedDate;
    } else if (selectedDate) {
        const parsedSel = parseFlexibleDate(selectedDate);
        const matched = dates.find(d => {
            const p = parseFlexibleDate(d);
            return p && parsedSel && p.getTime() === parsedSel.getTime();
        });
        if (matched) {
            select.value = matched;
        } else {
            const newOpt = document.createElement('option');
            newOpt.value = selectedDate;
            newOpt.textContent = selectedDate;
            select.appendChild(newOpt);
            select.value = selectedDate;
        }
    } else if (dates.length > 0) {
        select.value = dates[0];
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

    populateManualCandidateDateOptions();
    const selectDate = document.getElementById('inputManualPelaksanaan');
    const curDate = selectDate ? selectDate.value : '';
    document.getElementById('selectManualSesi').value = '1';
    document.getElementById('inputManualWaktu').value = getSessionTime(1, curDate);

    modal.classList.remove('hidden');
    modal.classList.add('flex');
    if (window.lucide) window.lucide.createIcons();
};

window.editCandidate = (candidateIdOrNip) => {
    const safeLookup = String(candidateIdOrNip || '').trim();
    const cand = currentCandidates.find(c => String(c.nip || '').trim() === safeLookup || String(c.id || '').trim() === safeLookup);
    if (!cand) return;

    if (cand.kehadiran === 'HADIR' || cand.kehadiran === 'TIDAK_HADIR') {
        showToast("Data peserta sudah terverifikasi (Hadir/Tidak Hadir). Reset status kehadiran terlebih dahulu jika ingin mengedit.", "warning");
        return;
    }

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

    populateManualCandidateDateOptions(cand.pelaksanaan);
    document.getElementById('selectManualSesi').value = cand.sesi || 1;
    const curDate = document.getElementById('inputManualPelaksanaan').value || cand.pelaksanaan;
    document.getElementById('inputManualWaktu').value = cand.waktu || getSessionTime(cand.sesi || 1, curDate);

    modal.classList.remove('hidden');
    modal.classList.add('flex');
    if (window.lucide) window.lucide.createIcons();
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
        // Urutkan peserta: HADIR dahulu, kemudian TIDAK HADIR, masing-masing tetap diurutkan Nama A-Z (Ascending)
        const sortedList = [...group.candidates].sort((a, b) => {
            const pA = a.kehadiran === 'HADIR' ? 1 : 2;
            const pB = b.kehadiran === 'HADIR' ? 1 : 2;
            if (pA !== pB) {
                return pA - pB;
            }
            return String(a.nama || '').localeCompare(String(b.nama || ''), 'id', { sensitivity: 'base' });
        });

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
                        <div>Total Peserta: <strong>${totalPesertaSesi} Orang</strong></div>
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
        // Scroll ke paling atas dengan halus
        window.scrollTo({ top: 0, behavior: 'smooth' });

        // Jika tab membutuhkan PIN dan belum terotorisasi, tampilkan modal PIN
        const needsPin = (tabName === 'create-ujian' || tabName === 'upload-excel' || tabName === 'master-wilker' || tabName === 'audit');
        if (needsPin && !isPinAuthorized && !isSuperAdmin) {
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
            return;
        }

        // Buka tab tujuan secara instan dan simpan posisinya
        window.switchTab(tabName);
    };

    window.verifyPinAndProceed = async (event) => {
        if (event) event.preventDefault();
        const inputPin = document.getElementById('inputAccessPin');
        const errorMsg = document.getElementById('pinErrorMessage');
        const pinVal = inputPin ? inputPin.value.trim() : '';

        if (pinVal === '1414' || pinVal === '141414') {
            isPinAuthorized = true;
            sessionStorage.setItem('is_admin_pin_authorized', 'true');

            // Simpan aksi dan target tab tertunda sebelum menutup modal
            const actionToExecute = pendingActionAfterPin;
            const targetTabToSwitch = pendingTargetTab;

            window.closeModalPinAccess();

            // 1. Eksekusi Kosongkan Peserta jika aksi tertunda adalah CLEAR_CANDIDATES
            if (actionToExecute === 'CLEAR_CANDIDATES') {
                await executeClearCandidates();
                return;
            }

            // 2. Buka Pengaturan Cloud jika aksi tertunda adalah OPEN_FIREBASE_CONFIG
            if (actionToExecute === 'OPEN_FIREBASE_CONFIG') {
                window.openModalFirebaseConfig();
                return;
            }

            // 3. Pindah ke tab tujuan secara instan
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
        // Simpan tab aktif di sessionStorage dan URL Hash agar saat user F5 (refresh manual), tetap di tab ini
        sessionStorage.setItem('active_tab', tabName);
        if (window.history && window.history.replaceState) {
            window.history.replaceState(null, null, `#${tabName}`);
        }
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

        // Tampilkan/sembunyikan floating bubble presensi khusus di tab daftar-peserta
        const bubble = document.getElementById('floatingAttendanceBubble');
        if (bubble) {
            if (tabName === 'daftar-peserta') {
                bubble.classList.remove('hidden');
                updateFloatingAttendanceBubble();
            } else {
                bubble.classList.add('hidden');
            }
        }

        if (tabName === 'master-wilker') {
            renderMasterInstansiTableFull();
        } else if (tabName === 'daftar-peserta') {
            applyCandidateFilters();
        } else if (tabName === 'dashboard') {
            renderDashboardStats();
        } else if (tabName === 'create-ujian') {
            renderExamListInCreateTab();
        } else if (tabName === 'audit') {
            if (window.updateAuditTabExamInfo) window.updateAuditTabExamInfo();
        }

        if (window.lucide) window.lucide.createIcons();
    };
}

/**
 * Setup Floating Tombol Naik ke Atas
 * Muncul otomatis saat halaman di-scroll ke bawah > 300px
 */
function setupScrollToTopButton() {
    const btn = document.getElementById('btnScrollToTop');
    if (!btn) return;

    window.addEventListener('scroll', () => {
        if (window.scrollY > 300) {
            btn.classList.remove('opacity-0', 'translate-y-6', 'pointer-events-none');
            btn.classList.add('opacity-100', 'translate-y-0', 'pointer-events-auto');
        } else {
            btn.classList.remove('opacity-100', 'translate-y-0', 'pointer-events-auto');
            btn.classList.add('opacity-0', 'translate-y-6', 'pointer-events-none');
        }
    }, { passive: true });

    window.scrollToTopPage = () => {
        window.scrollTo({
            top: 0,
            behavior: 'smooth'
        });
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
        const dayIdx = Math.max(0, uniqueDates.indexOf(dateStr));
        const rowData = {
            date: dateStr,
            dayIdx: dayIdx,
            hadir: { jpt: 0, admin: 0, pengawas: 0, jf: 0, pelaksana: 0, kosong: 0 },
            tidakHadir: { jpt: 0, admin: 0, pengawas: 0, jf: 0, pelaksana: 0, kosong: 0 },
            total: candsDate.length,
            sessionBreakdown: []
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

        // Rincian Akumulasi Sesi per Tanggal (Jumat otomatis 2 sesi)
        const sessions = isFriday(dateStr) ? [1, 2] : [1, 2, 3];
        const hasSession0 = candsDate.some(c => !c.sesi || c.sesi === 'NULL' || c.sesi === '00' || c.sesi === 0);
        if (hasSession0) sessions.unshift(0);
        if (isFriday(dateStr) && candsDate.some(c => Number(c.sesi) === 3)) {
            sessions.push(3);
        }

        sessions.forEach(s => {
            const candsSession = candsDate.filter(c => {
                if (s === 0) return !c.sesi || c.sesi === 'NULL' || c.sesi === '00' || c.sesi === 0;
                return Number(c.sesi) === s;
            });

            if (candsSession.length > 0 || (s >= 1 && s <= (isFriday(dateStr) ? 2 : 3))) {
                const cumSesiNum = calculateCumulativeSessionNumber({ pelaksanaan: dateStr, sesi: s }, uniqueDates) || ((dayIdx * 3) + s);
                const label = s === 0 ? 'Belum Terjadwal (00)' : `Sesi ${s} [${cumSesiNum}]`;

                const sessData = {
                    sesiNum: s,
                    cumNum: s === 0 ? 0 : cumSesiNum,
                    label: label,
                    hadir: { jpt: 0, admin: 0, pengawas: 0, jf: 0, pelaksana: 0, kosong: 0 },
                    tidakHadir: { jpt: 0, admin: 0, pengawas: 0, jf: 0, pelaksana: 0, kosong: 0 },
                    totalHadir: 0,
                    totalTidakHadir: 0,
                    total: candsSession.length
                };

                candsSession.forEach(c => {
                    const cls = classifyCandidateKelompok(c);
                    const isHadir = c.kehadiran === 'HADIR';
                    const isTidakHadir = c.kehadiran === 'TIDAK_HADIR';
                    const targetObj = isHadir ? sessData.hadir : (isTidakHadir ? sessData.tidakHadir : null);

                    let pilar = 'pelaksana';
                    if (cls.category === 'JPT_PRATAMA') pilar = 'jpt';
                    else if (cls.category === 'ADMINISTRATOR') pilar = 'admin';
                    else if (cls.category === 'PENGAWAS' || cls.category === 'ESELON_V') pilar = 'pengawas';
                    else if (cls.category === 'FUNGSIONAL') pilar = 'jf';
                    else if (cls.category === 'PELAKSANA') pilar = 'pelaksana';
                    else if (cls.category === 'KOSONG') pilar = 'kosong';

                    if (targetObj) targetObj[pilar]++;
                });

                sessData.totalHadir = ['jpt', 'admin', 'pengawas', 'jf', 'pelaksana', 'kosong'].reduce((acc, k) => acc + sessData.hadir[k], 0);
                sessData.totalTidakHadir = ['jpt', 'admin', 'pengawas', 'jf', 'pelaksana', 'kosong'].reduce((acc, k) => acc + sessData.tidakHadir[k], 0);

                rowData.sessionBreakdown.push(sessData);
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

        let tbodyHtml = tab1Rows.map((r, idx) => {
            const childSubRowsHtml = (r.sessionBreakdown || []).map((sb, sIdx) => {
                const totHadir = sb.totalHadir !== undefined ? sb.totalHadir : (['jpt', 'admin', 'pengawas', 'jf', 'pelaksana', 'kosong'].reduce((acc, k) => acc + (sb.hadir[k] || 0), 0));
                const totTidakHadir = sb.totalTidakHadir !== undefined ? sb.totalTidakHadir : (['jpt', 'admin', 'pengawas', 'jf', 'pelaksana', 'kosong'].reduce((acc, k) => acc + (sb.tidakHadir[k] || 0), 0));

                return `
                <tr class="hover:bg-blue-100/80 transition-colors border-b border-slate-300 text-[9px] sm:text-[10px] bg-slate-50/80">
                    <td class="p-1 sm:p-1.5 text-center font-bold text-slate-900 border-r border-slate-300 bg-slate-200/90 truncate">
                        ${sb.label}
                    </td>
                    <!-- Hadir (Shade Hijau Lebih Gelap Sedikit) -->
                    <td class="p-1 sm:p-1.5 text-center text-emerald-950 font-bold border-r border-slate-200 bg-emerald-100/70">${sb.hadir.jpt}</td>
                    <td class="p-1 sm:p-1.5 text-center text-emerald-950 font-bold border-r border-slate-200 bg-emerald-100/70">${sb.hadir.admin}</td>
                    <td class="p-1 sm:p-1.5 text-center text-emerald-950 font-bold border-r border-slate-200 bg-emerald-100/70">${sb.hadir.pengawas}</td>
                    <td class="p-1 sm:p-1.5 text-center text-emerald-950 font-extrabold border-r border-slate-200 bg-emerald-200/80">${sb.hadir.jf}</td>
                    <td class="p-1 sm:p-1.5 text-center text-emerald-950 font-bold border-r border-slate-200 bg-emerald-100/70">${sb.hadir.pelaksana}</td>
                    <!-- Kolom Tambahan: Total Hadir Sesi -->
                    <td class="p-1 sm:p-1.5 text-center text-emerald-950 font-black border-r border-slate-400 bg-emerald-200/90 shadow-2xs">${totHadir}</td>

                    <!-- Tidak Hadir (Shade Merah Lebih Gelap Sedikit) -->
                    <td class="p-1 sm:p-1.5 text-center text-rose-950 font-bold border-r border-slate-200 bg-rose-100/70">${sb.tidakHadir.jpt}</td>
                    <td class="p-1 sm:p-1.5 text-center text-rose-950 font-bold border-r border-slate-200 bg-rose-100/70">${sb.tidakHadir.admin}</td>
                    <td class="p-1 sm:p-1.5 text-center text-rose-950 font-bold border-r border-slate-200 bg-rose-100/70">${sb.tidakHadir.pengawas}</td>
                    <td class="p-1 sm:p-1.5 text-center text-rose-950 font-extrabold border-r border-slate-200 bg-rose-200/80">${sb.tidakHadir.jf}</td>
                    <td class="p-1 sm:p-1.5 text-center text-rose-950 font-bold border-r border-slate-200 bg-rose-100/70">${sb.tidakHadir.pelaksana}</td>
                    <!-- Kolom Tambahan: Total Tidak Hadir Sesi -->
                    <td class="p-1 sm:p-1.5 text-center text-rose-950 font-black border-r border-slate-400 bg-rose-200/90 shadow-2xs">${totTidakHadir}</td>

                    <!-- Total -->
                    <td class="p-1 sm:p-1.5 text-center font-extrabold text-slate-900 border-r border-slate-300 bg-slate-200">${sb.total}</td>
                    <!-- Aksi -->
                    <td class="p-1 sm:p-1.5 text-center">
                        <button type="button" 
                                onclick="event.stopPropagation(); copySessionRekapKelJabatan(${idx}, ${sIdx})" 
                                class="w-full py-0.5 px-1 bg-white hover:bg-slate-800 border border-slate-300 hover:border-slate-800 text-slate-700 hover:text-white rounded text-[9px] font-bold transition cursor-pointer flex items-center justify-center gap-0.5 shadow-2xs" 
                                title="Copy baris ${sb.label}">
                            <i data-lucide="copy" class="w-2.5 h-2.5 flex-shrink-0"></i>
                            <span class="hidden sm:inline">Copy</span>
                        </button>
                    </td>
                </tr>
                `;
            }).join('');

            return `
                <tr class="hover:bg-blue-50/60 transition-colors border-b border-slate-100 text-[10px] sm:text-[11px] cursor-pointer group select-none"
                    onclick="toggleRekapSesiChildRow(${idx})"
                    title="Klik untuk melihat / menutup rincian akumulasi sesi ${r.date}">
                    <td class="p-1 sm:p-1.5 font-bold text-slate-800 border-r border-slate-200 bg-slate-50/70">
                        <div class="flex items-center justify-between gap-1 px-1">
                            <span class="truncate">${r.date}</span>
                            <span class="inline-flex items-center justify-center w-4 h-4 rounded hover:bg-slate-200 transition">
                                <i id="rekap-chevron-${idx}" data-lucide="chevron-right" class="w-3.5 h-3.5 text-slate-400 group-hover:text-blue-600 transition-all duration-200"></i>
                            </span>
                        </div>
                    </td>
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
                    <td class="p-1 sm:p-1.5 text-center" onclick="event.stopPropagation()">
                        <button type="button" 
                                onclick="event.stopPropagation(); copyRowRekapKelJabatan(${idx})" 
                                class="w-full py-1 px-1 bg-slate-100 hover:bg-emerald-600 text-slate-700 hover:text-white rounded text-[10px] font-bold transition flex items-center justify-center gap-1 mx-auto cursor-pointer" 
                                title="Copy baris ${r.date}">
                            <i data-lucide="copy" class="w-3 h-3 flex-shrink-0"></i>
                            <span class="hidden sm:inline">Copy</span>
                        </button>
                    </td>
                </tr>

                <!-- Sub-tabel Rincian Akumulasi Sesi (Tersembunyi Awalnya - Shade Lebih Gelap Sedikit) -->
                <tr id="rekap-child-row-${idx}" class="hidden bg-slate-200/90 border-b-2 border-slate-300">
                    <td colspan="13" class="p-2 sm:p-3">
                        <div class="bg-slate-100 border-2 border-slate-300 rounded-xl p-3 shadow-xs">
                            <div class="flex items-center justify-between mb-2 px-1">
                                <span class="text-[10px] sm:text-xs font-extrabold text-slate-800 flex items-center gap-1.5">
                                    <i data-lucide="layers" class="w-3.5 h-3.5 text-blue-700"></i>
                                    Rincian Sesi Pelaksanaan (${r.date})
                                </span>
                                <span class="text-[9px] font-bold text-slate-700 bg-slate-200 px-2 py-0.5 rounded-full border border-slate-300">Format: Sesi [Akumulasi]</span>
                            </div>
                            <table class="w-full table-fixed text-left border-collapse border border-slate-400 rounded-lg overflow-hidden shadow-2xs">
                                <thead class="text-white text-[9px] sm:text-[10px] uppercase font-extrabold select-none shadow-xs border-b border-slate-700">
                                    <tr>
                                        <th class="p-1 sm:p-1.5 text-center bg-slate-900 text-amber-300 border-r border-slate-700" style="width: 10%;">Sesi</th>
                                        <th class="p-1 sm:p-1.5 text-center bg-emerald-900 text-emerald-100 border-r border-emerald-800 truncate" style="width: 6%;" title="Hadir JPT Pratama">JPT</th>
                                        <th class="p-1 sm:p-1.5 text-center bg-emerald-900 text-emerald-100 border-r border-emerald-800 truncate" style="width: 6%;" title="Hadir Administrator">Admin</th>
                                        <th class="p-1 sm:p-1.5 text-center bg-emerald-900 text-emerald-100 border-r border-emerald-800 truncate" style="width: 6.5%;" title="Hadir Pengawas">Pengawas</th>
                                        <th class="p-1 sm:p-1.5 text-center bg-emerald-900 text-emerald-100 border-r border-emerald-800 truncate" style="width: 5.5%;" title="Hadir JF">JF</th>
                                        <th class="p-1 sm:p-1.5 text-center bg-emerald-900 text-emerald-100 border-r border-slate-300 truncate" style="width: 6.5%;" title="Hadir Pelaksana">Pelaksana</th>
                                        <th class="p-1 sm:p-1.5 text-center bg-emerald-950 text-emerald-200 border-r border-slate-400 truncate font-black" style="width: 7.5%;" title="Total Hadir Sesi Ini"><span class="hidden sm:inline">Tot Hadir</span><span class="sm:hidden">Hdr</span></th>
                                        <th class="p-1 sm:p-1.5 text-center bg-rose-900 text-rose-100 border-r border-rose-800 truncate" style="width: 6%;" title="Tidak Hadir JPT Pratama">JPT</th>
                                        <th class="p-1 sm:p-1.5 text-center bg-rose-900 text-rose-100 border-r border-rose-800 truncate" style="width: 6%;" title="Tidak Hadir Administrator">Admin</th>
                                        <th class="p-1 sm:p-1.5 text-center bg-rose-900 text-rose-100 border-r border-rose-800 truncate" style="width: 6.5%;" title="Tidak Hadir Pengawas">Pengawas</th>
                                        <th class="p-1 sm:p-1.5 text-center bg-rose-900 text-rose-100 border-r border-rose-800 truncate" style="width: 5.5%;" title="Tidak Hadir JF">JF</th>
                                        <th class="p-1 sm:p-1.5 text-center bg-rose-900 text-rose-100 border-r border-slate-300 truncate" style="width: 6.5%;" title="Tidak Hadir Pelaksana">Pelaksana</th>
                                        <th class="p-1 sm:p-1.5 text-center bg-rose-950 text-rose-200 border-r border-slate-400 truncate font-black" style="width: 7.5%;" title="Total Tidak Hadir Sesi Ini"><span class="hidden sm:inline">Tot T.Hadir</span><span class="sm:hidden">TH</span></th>
                                        <th class="p-1 sm:p-1.5 text-center bg-slate-900 text-white border-r border-slate-700" style="width: 7.5%;">Total</th>
                                        <th class="p-1 sm:p-1.5 text-center bg-slate-900 text-white" style="width: 6.5%;">Aksi</th>
                                    </tr>
                                </thead>
                                <tbody class="divide-y divide-slate-300 bg-slate-50">
                                    ${childSubRowsHtml}
                                </tbody>
                            </table>
                        </div>
                    </td>
                </tr>
            `;
        }).join('');

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

/**
 * Toggle tampil/sembunyi sub-tabel akumulasi sesi pada Tab 1 Rekapitulasi
 */
window.toggleRekapSesiChildRow = (idx) => {
    const childRow = document.getElementById(`rekap-child-row-${idx}`);
    const chevron = document.getElementById(`rekap-chevron-${idx}`);
    if (!childRow) return;

    const isHidden = childRow.classList.contains('hidden');
    if (isHidden) {
        childRow.classList.remove('hidden');
        if (chevron) {
            chevron.classList.add('rotate-90');
            chevron.classList.add('text-blue-600');
        }
        if (window.lucide) window.lucide.createIcons();
    } else {
        childRow.classList.add('hidden');
        if (chevron) {
            chevron.classList.remove('rotate-90');
            chevron.classList.remove('text-blue-600');
        }
    }
};

/**
 * Copy 1 baris rincian sesi dari Tab 1 Rekapitulasi (TSV Excel-Ready)
 */
window.copySessionRekapKelJabatan = (rowIdx, sIdx) => {
    const row = currentRekapData.tab1Rows && currentRekapData.tab1Rows[rowIdx];
    if (!row || !row.sessionBreakdown || !row.sessionBreakdown[sIdx]) return;
    const sb = row.sessionBreakdown[sIdx];
    const totHadir = sb.totalHadir !== undefined ? sb.totalHadir : (['jpt', 'admin', 'pengawas', 'jf', 'pelaksana', 'kosong'].reduce((acc, k) => acc + (sb.hadir[k] || 0), 0));
    const totTidakHadir = sb.totalTidakHadir !== undefined ? sb.totalTidakHadir : (['jpt', 'admin', 'pengawas', 'jf', 'pelaksana', 'kosong'].reduce((acc, k) => acc + (sb.tidakHadir[k] || 0), 0));

    const headerStr = `Tanggal\tSesi\tHadir JPT Pratama\tHadir Administrator\tHadir Pengawas\tHadir Jab. Fungsional\tHadir Pelaksana\tTotal Hadir\tTidak Hadir JPT Pratama\tTidak Hadir Administrator\tTidak Hadir Pengawas\tTidak Hadir Jab. Fungsional\tTidak Hadir Pelaksana\tTotal Tidak Hadir\tTotal`;
    const dataStr = `${row.date}\t${sb.label}\t${sb.hadir.jpt}\t${sb.hadir.admin}\t${sb.hadir.pengawas}\t${sb.hadir.jf}\t${sb.hadir.pelaksana}\t${totHadir}\t${sb.tidakHadir.jpt}\t${sb.tidakHadir.admin}\t${sb.tidakHadir.pengawas}\t${sb.tidakHadir.jf}\t${sb.tidakHadir.pelaksana}\t${totTidakHadir}\t${sb.total}`;

    copyTextToClipboard(`${headerStr}\n${dataStr}`, `Data ${sb.label} (${row.date}) berhasil disalin! Format siap di-paste ke Excel.`);
};

// =========================================================================
// FITUR 1: MODAL RINCIAN PESERTA KELOMPOK JABATAN (HADIR, TIDAK HADIR, BELUM)
// =========================================================================

let currentDetailKelJabatanList = [];
let currentDetailKelJabatanStatusFilter = 'ALL';
let currentDetailKelJabatanTitle = '';

window.openModalDetailKelompokJabatan = (key, encodedLabel) => {
    const label = encodedLabel ? decodeURIComponent(encodedLabel) : key;
    currentDetailKelJabatanTitle = label;
    currentDetailKelJabatanStatusFilter = 'ALL';

    const categories = window._activeKelJabatanStats;
    let nips = [];

    if (categories) {
        if (key && key.startsWith('FUNGSIONAL:')) {
            const sub = key.replace('FUNGSIONAL:', '');
            nips = (categories.FUNGSIONAL && categories.FUNGSIONAL.children && categories.FUNGSIONAL.children[sub]) 
                ? (categories.FUNGSIONAL.children[sub].candidateNips || categories.FUNGSIONAL.children[sub].candidates || []) 
                : [];
        } else if (key && key.startsWith('LAINNYA:')) {
            const other = key.replace('LAINNYA:', '');
            nips = (categories.LAINNYA && categories.LAINNYA[other]) 
                ? (categories.LAINNYA[other].candidateNips || categories.LAINNYA[other].candidates || []) 
                : [];
        } else if (categories[key]) {
            nips = categories[key].stat ? (categories[key].stat.candidateNips || categories[key].stat.candidates || []) : [];
        }
    }

    // Buat lookup Map cepat dari currentCandidates agar hemat memori & O(1) akses
    const candMap = new Map();
    (currentCandidates || []).forEach(c => {
        const k = String(c.nip || c.id || '').trim();
        if (k) candMap.set(k, c);
    });

    let list = [];
    nips.forEach(item => {
        if (typeof item === 'string') {
            const found = candMap.get(item);
            if (found) list.push(found);
        } else if (item && typeof item === 'object') {
            list.push(item);
        }
    });

    currentDetailKelJabatanList = list;

    // Set judul modal
    const titleEl = document.getElementById('detailKelJabatanTitle');
    if (titleEl) titleEl.textContent = label;

    // Reset input search
    const searchInput = document.getElementById('inputSearchDetailKelJabatan');
    if (searchInput) searchInput.value = '';

    // Hitung badge counter
    const total = list.length;
    const hadir = list.filter(c => c.kehadiran === 'HADIR').length;
    const tidakHadir = list.filter(c => c.kehadiran === 'TIDAK_HADIR').length;
    const belum = Math.max(0, total - hadir - tidakHadir);

    const bTotal = document.getElementById('detailBadgeTotal');
    const bHadir = document.getElementById('detailBadgeHadir');
    const bTidakHadir = document.getElementById('detailBadgeTidakHadir');
    const bBelum = document.getElementById('detailBadgeBelum');

    if (bTotal) bTotal.textContent = total;
    if (bHadir) bHadir.textContent = hadir;
    if (bTidakHadir) bTidakHadir.textContent = tidakHadir;
    if (bBelum) bBelum.textContent = belum;

    window.setDetailKelJabatanStatusFilter('ALL');

    const modal = document.getElementById('modalDetailKelompokJabatan');
    if (modal) {
        modal.classList.remove('hidden');
        modal.classList.add('flex');
    }
    if (window.lucide) window.lucide.createIcons();
};

window.closeModalDetailKelompokJabatan = () => {
    const modal = document.getElementById('modalDetailKelompokJabatan');
    if (modal) {
        modal.classList.add('hidden');
        modal.classList.remove('flex');
    }
    // Bersihkan memori dan DOM table saat modal ditutup
    currentDetailKelJabatanList = [];
    const tbody = document.getElementById('tbodyDetailKelJabatan');
    if (tbody) tbody.innerHTML = '';
};

window.setDetailKelJabatanStatusFilter = (status) => {
    currentDetailKelJabatanStatusFilter = status;

    const btnAll = document.getElementById('btnFilterDetailAll');
    const btnHadir = document.getElementById('btnFilterDetailHadir');
    const btnTidakHadir = document.getElementById('btnFilterDetailTidakHadir');
    const btnBelum = document.getElementById('btnFilterDetailBelum');

    const resetBtn = (btn) => {
        if (!btn) return;
        btn.classList.remove('bg-bkn-700', 'text-white', 'border-bkn-700', 'bg-emerald-700', 'bg-rose-700', 'bg-amber-700');
        btn.classList.add('bg-white', 'border-slate-300');
    };

    resetBtn(btnAll);
    resetBtn(btnHadir);
    resetBtn(btnTidakHadir);
    resetBtn(btnBelum);

    if (status === 'ALL' && btnAll) {
        btnAll.classList.add('bg-bkn-700', 'text-white', 'border-bkn-700');
        btnAll.classList.remove('bg-white', 'border-slate-300');
    } else if (status === 'HADIR' && btnHadir) {
        btnHadir.classList.add('bg-emerald-700', 'text-white', 'border-emerald-700');
        btnHadir.classList.remove('bg-white', 'border-slate-300');
    } else if (status === 'TIDAK_HADIR' && btnTidakHadir) {
        btnTidakHadir.classList.add('bg-rose-700', 'text-white', 'border-rose-700');
        btnTidakHadir.classList.remove('bg-white', 'border-slate-300');
    } else if (status === 'BELUM' && btnBelum) {
        btnBelum.classList.add('bg-amber-700', 'text-white', 'border-amber-700');
        btnBelum.classList.remove('bg-white', 'border-slate-300');
    }

    window.renderDetailKelJabatanTable();
};

window.renderDetailKelJabatanTable = () => {
    const tbody = document.getElementById('tbodyDetailKelJabatan');
    if (!tbody) return;

    const searchInput = document.getElementById('inputSearchDetailKelJabatan');
    const term = searchInput ? searchInput.value.trim().toLowerCase() : '';

    let list = currentDetailKelJabatanList;

    // Filter status presensi
    if (currentDetailKelJabatanStatusFilter === 'HADIR') {
        list = list.filter(c => c.kehadiran === 'HADIR');
    } else if (currentDetailKelJabatanStatusFilter === 'TIDAK_HADIR') {
        list = list.filter(c => c.kehadiran === 'TIDAK_HADIR');
    } else if (currentDetailKelJabatanStatusFilter === 'BELUM') {
        list = list.filter(c => c.kehadiran !== 'HADIR' && c.kehadiran !== 'TIDAK_HADIR');
    }

    // Filter search NIP / Nama
    if (term) {
        list = list.filter(c => 
            String(c.nip || '').toLowerCase().includes(term) ||
            String(c.nama || '').toLowerCase().includes(term) ||
            String(c.jabatan || '').toLowerCase().includes(term) ||
            String(c.unitKerja || '').toLowerCase().includes(term)
        );
    }

    const footerInfo = document.getElementById('detailKelJabatanFooterInfo');
    if (footerInfo) footerInfo.textContent = `Menampilkan ${list.length} dari ${currentDetailKelJabatanList.length} peserta`;

    if (list.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="p-8 text-center text-slate-400 italic">Tidak ada peserta yang cocok dengan kriteria filter.</td></tr>`;
        return;
    }

    tbody.innerHTML = list.map((c, idx) => {
        const isHadir = c.kehadiran === 'HADIR';
        const isTidakHadir = c.kehadiran === 'TIDAK_HADIR';

        let badgeStatus = '';
        if (isHadir) {
            badgeStatus = `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-extrabold bg-emerald-100 text-emerald-800 border border-emerald-300"><span class="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>HADIR</span>`;
        } else if (isTidakHadir) {
            badgeStatus = `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-extrabold bg-rose-100 text-rose-800 border border-rose-300"><span class="w-1.5 h-1.5 rounded-full bg-rose-500"></span>TIDAK HADIR</span>`;
        } else {
            badgeStatus = `<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-extrabold bg-amber-100 text-amber-800 border border-amber-300"><span class="w-1.5 h-1.5 rounded-full bg-amber-400"></span>BELUM</span>`;
        }

        const sesiText = (c.sesi !== undefined && c.sesi !== null && c.sesi !== 'NULL') ? `Sesi ${c.sesi}` : '-';
        const tglText = c.pelaksanaan || '-';

        return `
            <tr class="hover:bg-slate-50 transition">
                <td class="p-2.5 text-center font-bold text-slate-400">${idx + 1}</td>
                <td class="p-2.5 text-center">${badgeStatus}</td>
                <td class="p-2.5 font-mono font-semibold text-slate-800 select-all">${c.nip || '-'}</td>
                <td class="p-2.5 font-bold text-slate-900">${c.nama || '-'}</td>
                <td class="p-2.5 text-slate-700">${c.jabatan || '-'}</td>
                <td class="p-2.5 text-slate-600">${c.unitKerja || '-'}</td>
                <td class="p-2.5 text-center text-[11px]">
                    <span class="font-semibold text-slate-800 block">${tglText}</span>
                    <span class="text-bkn-700 font-bold">${sesiText}</span>
                </td>
            </tr>
        `;
    }).join('');
};

window.copyDetailKelJabatanToClipboard = () => {
    if (!currentDetailKelJabatanList || currentDetailKelJabatanList.length === 0) {
        showToast("Tidak ada data untuk disalin.", "info");
        return;
    }

    const header = ["No", "Status", "NIP", "Nama", "Jabatan", "Unit Kerja", "Tanggal", "Sesi"].join('\t');
    const rows = currentDetailKelJabatanList.map((c, i) => [
        i + 1,
        c.kehadiran || 'BELUM PRESENSI',
        `'${c.nip || ''}`,
        c.nama || '',
        c.jabatan || '',
        c.unitKerja || '',
        c.pelaksanaan || '',
        c.sesi || ''
    ].join('\t'));

    const text = [header, ...rows].join('\n');
    copyTextToClipboard(text, `Data ${currentDetailKelJabatanTitle} (${currentDetailKelJabatanList.length} peserta) berhasil disalin ke clipboard!`);
};

// =========================================================================
// FITUR 2: MODUL AUDIT & SINKRONISASI PASCA UJIAN (CAT BKN)
// =========================================================================

let currentAuditResult = null;
let selectedAuditNips = new Set();
let auditCompareCategoryFilter = 'ALL';

function setupAuditUI() {
    window.updateAuditTabExamInfo = () => {
        const badge = document.getElementById('auditActiveExamName');
        if (badge) {
            if (currentExam) {
                badge.textContent = `Instansi Ujian Aktif: ${currentExam.instansi || currentExam.title}`;
            } else {
                badge.textContent = `Instansi Ujian: Belum Ada Ujian Aktif`;
            }
        }
    };

    window.handleAuditDrop = (e) => {
        e.preventDefault();
        const dropzone = document.getElementById('dropzoneAuditExcel');
        if (dropzone) dropzone.classList.remove('border-bkn-600', 'bg-blue-50/40');

        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            window.processAuditFile(e.dataTransfer.files[0]);
        }
    };

    window.handleAuditFileInput = (e) => {
        if (e.target && e.target.files && e.target.files.length > 0) {
            window.processAuditFile(e.target.files[0]);
        }
    };

    window.processAuditFile = async (file) => {
        if (!currentExam) {
            showToast("Harap pilih Instansi Ujian Aktif terlebih dahulu sebelum melakukan audit!", "error");
            return;
        }

        const validExts = ['.xlsx', '.xls', '.csv'];
        const isExtValid = validExts.some(ext => file.name.toLowerCase().endsWith(ext));
        if (!isExtValid) {
            showToast("Format file tidak didukung. Harap upload file .xlsx, .xls, atau .csv!", "error");
            return;
        }

        const dropzone = document.getElementById('dropzoneAuditExcel');
        const loading = document.getElementById('auditLoadingState');
        const resultPanel = document.getElementById('auditResultSummary');

        if (dropzone) dropzone.classList.add('hidden');
        if (loading) loading.classList.remove('hidden');
        if (resultPanel) resultPanel.classList.add('hidden');

        try {
            const auditRows = await parseAuditExcel(file);
            const sortedDates = getSortedExamDates();
            const comparison = compareAuditDataWithDatabase(auditRows, currentCandidates, { sortedDates });
            currentAuditResult = comparison;

            // Default: pilih semua NIP yang memiliki perbedaan
            selectedAuditNips = new Set(comparison.diffCandidates.map(d => String(d.nip).trim()));

            // Update UI Ringkasan Hasil
            const elTotal = document.getElementById('auditStatTotalExcel');
            const elMatched = document.getElementById('auditStatMatched');
            const elAutoHadir = document.getElementById('auditStatAutoHadir');
            const elDiff = document.getElementById('auditStatDiffCount');
            const elUnmatchedAlert = document.getElementById('auditUnmatchedAlert');
            const elUnmatchedCount = document.getElementById('auditUnmatchedCount');

            if (elTotal) elTotal.textContent = comparison.totalExcel;
            if (elMatched) elMatched.textContent = comparison.matchedCount;
            if (elAutoHadir) elAutoHadir.textContent = comparison.autoHadirCount;
            if (elDiff) elDiff.textContent = comparison.diffCount;

            if (elUnmatchedAlert && elUnmatchedCount) {
                if (comparison.unmatchedCount > 0) {
                    elUnmatchedCount.textContent = comparison.unmatchedCount;
                    elUnmatchedAlert.classList.remove('hidden');
                } else {
                    elUnmatchedAlert.classList.add('hidden');
                }
            }

            if (loading) loading.classList.add('hidden');
            if (resultPanel) resultPanel.classList.remove('hidden');

            if (window.lucide) window.lucide.createIcons();

            if (comparison.diffCount === 0) {
                showToast("Data hasil audit sudah 100% cocok dengan database! Tidak ada perbedaan yang perlu diubah.", "success");
            } else {
                showToast(`Analisis selesai: Terdeteksi ${comparison.diffCount} peserta dengan perbedaan data siap ditinjau.`, "info");
                setTimeout(() => {
                    window.openModalAuditCompare();
                }, 300);
            }
        } catch (err) {
            console.error("Error processing audit file:", err);
            showToast("Gagal menganalisis file audit: " + err.message, "error");
            if (loading) loading.classList.add('hidden');
            if (dropzone) dropzone.classList.remove('hidden');
        }
    };

    window.resetAuditUpload = () => {
        currentAuditResult = null;
        selectedAuditNips.clear();
        const dropzone = document.getElementById('dropzoneAuditExcel');
        const loading = document.getElementById('auditLoadingState');
        const resultPanel = document.getElementById('auditResultSummary');
        const fileInput = document.getElementById('inputAuditExcelFile');

        if (fileInput) fileInput.value = '';
        if (dropzone) dropzone.classList.remove('hidden');
        if (loading) loading.classList.add('hidden');
        if (resultPanel) resultPanel.classList.add('hidden');
    };

    window.openModalAuditCompare = () => {
        if (!currentAuditResult || currentAuditResult.diffCandidates.length === 0) {
            showToast("Tidak ada perbedaan data untuk ditampilkan.", "info");
            return;
        }

        auditCompareCategoryFilter = 'ALL';
        const searchInput = document.getElementById('inputSearchAuditCompare');
        if (searchInput) searchInput.value = '';

        const badgeTotal = document.getElementById('badgeAuditDiffTotal');
        if (badgeTotal) badgeTotal.textContent = `${currentAuditResult.diffCount} Perbedaan`;

        // Update kategori counter
        const countAll = currentAuditResult.diffCandidates.length;
        const countHadir = currentAuditResult.diffCandidates.filter(d => d.hasKehadiranChange).length;
        const countSesi = currentAuditResult.diffCandidates.filter(d => d.hasSesiChange).length;
        const countBio = currentAuditResult.diffCandidates.filter(d => d.changes.some(c => c.category === 'biodata' || c.category === 'waktu')).length;

        const elAll = document.getElementById('countFilterCompareAll');
        const elHadir = document.getElementById('countFilterCompareKehadiran');
        const elSesi = document.getElementById('countFilterCompareSesi');
        const elBio = document.getElementById('countFilterCompareBiodata');

        if (elAll) elAll.textContent = countAll;
        if (elHadir) elHadir.textContent = countHadir;
        if (elSesi) elSesi.textContent = countSesi;
        if (elBio) elBio.textContent = countBio;

        window.setAuditCompareCategoryFilter('ALL');

        const modal = document.getElementById('modalAuditCompare');
        if (modal) {
            modal.classList.remove('hidden');
            modal.classList.add('flex');
        }

        if (window.lucide) window.lucide.createIcons();
    };

    window.closeModalAuditCompare = () => {
        const modal = document.getElementById('modalAuditCompare');
        if (modal) {
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }
        const tbody = document.getElementById('tbodyAuditCompareList');
        if (tbody) tbody.innerHTML = '';
        if (selectedAuditNips) selectedAuditNips.clear();
    };

    window.setAuditCompareCategoryFilter = (category) => {
        auditCompareCategoryFilter = category;

        const btnAll = document.getElementById('btnFilterCompareAll');
        const btnHadir = document.getElementById('btnFilterCompareKehadiran');
        const btnSesi = document.getElementById('btnFilterCompareSesi');
        const btnBio = document.getElementById('btnFilterCompareBiodata');

        const resetBtn = (btn) => {
            if (!btn) return;
            btn.classList.remove('bg-bkn-700', 'text-white', 'bg-emerald-700', 'bg-indigo-700', 'bg-slate-800');
            btn.classList.add('bg-white');
        };

        resetBtn(btnAll);
        resetBtn(btnHadir);
        resetBtn(btnSesi);
        resetBtn(btnBio);

        if (category === 'ALL' && btnAll) {
            btnAll.classList.add('bg-bkn-700', 'text-white');
            btnAll.classList.remove('bg-white');
        } else if (category === 'kehadiran' && btnHadir) {
            btnHadir.classList.add('bg-emerald-700', 'text-white');
            btnHadir.classList.remove('bg-white');
        } else if (category === 'sesi' && btnSesi) {
            btnSesi.classList.add('bg-indigo-700', 'text-white');
            btnSesi.classList.remove('bg-white');
        } else if (category === 'biodata' && btnBio) {
            btnBio.classList.add('bg-slate-800', 'text-white');
            btnBio.classList.remove('bg-white');
        }

        window.renderAuditCompareTable();
    };

    window.toggleSelectAllAuditDifferences = (isChecked) => {
        if (!currentAuditResult) return;
        if (isChecked) {
            currentAuditResult.diffCandidates.forEach(d => {
                selectedAuditNips.add(String(d.nip).trim());
            });
        } else {
            selectedAuditNips.clear();
        }
        window.renderAuditCompareTable();
    };

    window.toggleAuditItemCheck = (nip, isChecked) => {
        const safeNip = String(nip).trim();
        if (isChecked) {
            selectedAuditNips.add(safeNip);
        } else {
            selectedAuditNips.delete(safeNip);
        }
        window.updateAuditSelectedCount();
    };

    window.updateAuditSelectedCount = () => {
        const total = currentAuditResult ? currentAuditResult.diffCandidates.length : 0;
        const selected = selectedAuditNips.size;

        const checkAll = document.getElementById('checkAllAuditChanges');
        if (checkAll) {
            checkAll.checked = (selected === total && total > 0);
            checkAll.indeterminate = (selected > 0 && selected < total);
        }

        const infoEl = document.getElementById('auditCompareSelectedInfo');
        const btnCount = document.getElementById('btnApplySelectedCount');
        const btnApply = document.getElementById('btnApplyAuditSelected');

        if (infoEl) infoEl.textContent = `${selected} dari ${total} peserta dipilih untuk di-replace`;
        if (btnCount) btnCount.textContent = selected;

        if (btnApply) {
            if (selected === 0) {
                btnApply.disabled = true;
                btnApply.classList.add('opacity-50', 'cursor-not-allowed');
            } else {
                btnApply.disabled = false;
                btnApply.classList.remove('opacity-50', 'cursor-not-allowed');
            }
        }
    };

    window.renderAuditCompareTable = () => {
        const tbody = document.getElementById('tbodyAuditCompareList');
        if (!tbody || !currentAuditResult) return;

        const searchInput = document.getElementById('inputSearchAuditCompare');
        const term = searchInput ? searchInput.value.trim().toLowerCase() : '';

        let list = currentAuditResult.diffCandidates;

        // Filter kategori
        if (auditCompareCategoryFilter === 'kehadiran') {
            list = list.filter(d => d.hasKehadiranChange);
        } else if (auditCompareCategoryFilter === 'sesi') {
            list = list.filter(d => d.hasSesiChange);
        } else if (auditCompareCategoryFilter === 'biodata') {
            list = list.filter(d => d.changes.some(c => c.category === 'biodata' || c.category === 'waktu'));
        }

        // Filter search
        if (term) {
            list = list.filter(d => 
                String(d.nip || '').toLowerCase().includes(term) ||
                String(d.nama || '').toLowerCase().includes(term)
            );
        }

        // Pastikan urutan selalu pegawai dengan jumlah perubahan data terbanyak berada paling atas
        list.sort((a, b) => {
            const countA = a.changes ? a.changes.length : 0;
            const countB = b.changes ? b.changes.length : 0;
            if (countB !== countA) return countB - countA;
            return String(a.nama || '').localeCompare(String(b.nama || ''), 'id', { sensitivity: 'base' });
        });

        if (list.length === 0) {
            tbody.innerHTML = `<tr><td colspan="7" class="p-8 text-center text-slate-400 italic">Tidak ada perubahan data yang cocok dengan kriteria filter.</td></tr>`;
            window.updateAuditSelectedCount();
            return;
        }

        tbody.innerHTML = list.map((item, idx) => {
            const isChecked = selectedAuditNips.has(String(item.nip).trim());

            // Build change summary rows
            const changeRowsHtml = item.changes.map(ch => {
                const isHadir = ch.field === 'kehadiran';
                const isSesi = ch.field === 'sesi';

                return `
                    <div class="py-1.5 border-b border-slate-100 last:border-0 grid grid-cols-12 gap-2 items-center text-xs">
                        <div class="col-span-3 font-bold text-slate-700 flex items-center gap-1.5">
                            ${isHadir ? '<span class="w-2 h-2 rounded-full bg-emerald-500"></span>' : (isSesi ? '<span class="w-2 h-2 rounded-full bg-indigo-500"></span>' : '<span class="w-2 h-2 rounded-full bg-slate-400"></span>')}
                            <span>${ch.label}</span>
                        </div>
                        <div class="col-span-3 text-rose-800 bg-rose-50/70 p-1 rounded font-medium line-through">
                            ${ch.oldValue}
                        </div>
                        <div class="col-span-3 text-emerald-800 bg-emerald-50/70 p-1 rounded font-bold flex items-center gap-1">
                            <i data-lucide="arrow-right" class="w-3 h-3 text-emerald-600 flex-shrink-0"></i>
                            <span>${ch.newValue}</span>
                        </div>
                        <div class="col-span-3 text-[11px] text-slate-500 italic">
                            ${ch.reason || '-'}
                        </div>
                    </div>
                `;
            }).join('');

            return `
                <tr class="hover:bg-blue-50/20 transition ${isChecked ? 'bg-white' : 'bg-slate-50/50 opacity-60'}">
                    <td class="p-3 text-center">
                        <input type="checkbox" 
                               class="audit-item-check rounded text-bkn-700 focus:ring-bkn-500 w-4 h-4 cursor-pointer"
                               ${isChecked ? 'checked' : ''}
                               onchange="window.toggleAuditItemCheck('${item.nip}', this.checked)">
                    </td>
                    <td class="p-3 text-center font-bold text-slate-400">${idx + 1}</td>
                    <td class="p-3">
                        <div class="flex items-center gap-1.5 flex-wrap">
                            <span class="font-bold text-slate-900">${item.nama || '-'}</span>
                            <span class="inline-flex items-center px-1.5 py-0.5 rounded text-[10px] font-bold bg-amber-100 text-amber-900 border border-amber-300">
                                ${item.changes ? item.changes.length : 0} Data Berbeda
                            </span>
                        </div>
                        <div class="font-mono text-[11px] text-slate-500 mt-0.5">${item.nip || '-'}</div>
                    </td>
                    <td colspan="4" class="p-2">
                        <div class="space-y-1">
                            ${changeRowsHtml}
                        </div>
                    </td>
                </tr>
            `;
        }).join('');

        window.updateAuditSelectedCount();
        if (window.lucide) window.lucide.createIcons();
    };

    window.applySelectedAuditDifferencesToDatabase = async () => {
        if (!currentExam) return;
        if (!currentAuditResult || selectedAuditNips.size === 0) {
            showToast("Pilih setidaknya 1 peserta untuk diperbarui!", "warning");
            return;
        }

        const selectedDiffs = currentAuditResult.diffCandidates.filter(d => selectedAuditNips.has(String(d.nip).trim()));
        if (selectedDiffs.length === 0) {
            showToast("Tidak ada peserta terpilih.", "warning");
            return;
        }

        const btnApply = document.getElementById('btnApplyAuditSelected');
        if (btnApply) {
            btnApply.disabled = true;
            btnApply.innerHTML = `<div class="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin mr-2 inline-block"></div> Menerapkan ${selectedDiffs.length} Data...`;
        }

        // Beri jeda 20ms agar browser sempat merender spinner pada tombol sebelum komputasi & network call
        await new Promise(r => setTimeout(r, 20));

        try {
            const updates = {};
            let countKehadiranUpdated = 0;
            let countSesiUpdated = 0;
            const nowIso = new Date().toISOString();

            // Bangun index lookup Map O(1) agar tidak O(N*M) linear scanning
            const candidateMap = new Map();
            currentCandidates.forEach(c => {
                candidateMap.set(String(c.nip).trim(), c);
            });

            selectedDiffs.forEach(item => {
                const nipClean = String(item.nip).trim();
                const safeKey = nipClean.replace(/[.#$[\]/]/g, '_');
                const pathPrefix = `candidates/${currentExam.id}/${safeKey}/`;

                // Update candidate lokal di memori secara instan
                const candInMem = candidateMap.get(nipClean);

                item.changes.forEach(ch => {
                    const rawVal = ch.rawNewValue !== undefined ? ch.rawNewValue : ch.newValue;
                    updates[pathPrefix + ch.field] = rawVal;

                    if (candInMem) {
                        candInMem[ch.field] = rawVal;
                    }

                    if (ch.field === 'sesi') {
                        countSesiUpdated++;

                        // Jika ada targetPelaksanaan hasil konversi sesi akumulasi
                        if (ch.targetPelaksanaan && ch.targetPelaksanaan !== 'NULL') {
                            updates[pathPrefix + 'pelaksanaan'] = ch.targetPelaksanaan;
                            if (candInMem) candInMem.pelaksanaan = ch.targetPelaksanaan;

                            // Perbarui jam pelaksanaan standar sesi
                            const newTime = getSessionTime(rawVal, ch.targetPelaksanaan);
                            updates[pathPrefix + 'waktu'] = newTime;
                            if (candInMem) candInMem.waktu = newTime;

                            const d = parseFlexibleDate(ch.targetPelaksanaan);
                            const isFri = d ? isFriday(d) : false;
                            updates[pathPrefix + 'isFriday'] = isFri;
                            if (candInMem) candInMem.isFriday = isFri;
                        }
                    }

                    if (ch.field === 'kehadiran') {
                        countKehadiranUpdated++;
                        updates[pathPrefix + 'attendanceTimestamp'] = nowIso;
                        if (candInMem) candInMem.attendanceTimestamp = nowIso;
                    }
                });

                // Set status Terjadwal jika sesi valid
                if (item.changes.some(c => c.field === 'sesi')) {
                    updates[pathPrefix + 'status'] = 'Terjadwal';
                    if (candInMem) candInMem.status = 'Terjadwal';
                }

                updates[pathPrefix + 'updatedAt'] = nowIso;
                if (candInMem) candInMem.updatedAt = nowIso;
            });

            // 1. Simpan ke Firebase Realtime Database SECARA BULK ATOMIK (1 Kali Request Cepat!)
            if (isCloudActive() && Object.keys(updates).length > 0) {
                await bulkUpdatePathsInCloud(updates);
            }

            // 2. Tutup modal komparasi review terlebih dahulu
            window.closeModalAuditCompare();

            // 3. Re-kalkulasi dan refresh seluruh tabel & statistik satu kali secara instan
            applyCandidateFilters();
            renderDashboardStats();
            updateFloatingAttendanceBubble();

            // 4. Reset upload file agar siap untuk upload audit baru berikutnya
            window.resetAuditUpload();

            // 5. Buka DIALOG KONFIRMASI SUKSES informatif & elegan!
            window.openModalAuditSuccess(selectedDiffs.length, countKehadiranUpdated, countSesiUpdated);
            showToast(`Sukses memperbarui ${selectedDiffs.length} data peserta!`, "success");

        } catch (err) {
            console.error("Gagal menerapkan perubahan audit:", err);
            showToast("Gagal menerapkan perubahan: " + err.message, "error");
        } finally {
            if (btnApply) {
                btnApply.disabled = false;
                btnApply.innerHTML = `<i data-lucide="check-check" class="w-4 h-4"></i><span>Terapkan Perubahan Terpilih</span>`;
                if (window.lucide) window.lucide.createIcons();
            }
        }
    };

    /**
     * Buka Dialog Konfirmasi Sukses Audit
     */
    window.openModalAuditSuccess = (totalCount, kehadiranCount, sesiCount) => {
        const modal = document.getElementById('modalAuditSuccess');
        const elTotal = document.getElementById('auditSuccessTotalCount');
        const elKehadiran = document.getElementById('auditSuccessKehadiranCount');
        const elSesi = document.getElementById('auditSuccessSesiCount');

        if (elTotal) elTotal.textContent = `${totalCount} Orang`;
        if (elKehadiran) elKehadiran.textContent = `${kehadiranCount} Peserta`;
        if (elSesi) elSesi.textContent = `${sesiCount} Peserta`;

        if (modal) {
            modal.classList.remove('hidden');
            modal.classList.add('flex');
        }

        if (window.lucide) window.lucide.createIcons();
    };

    /**
     * Tutup Dialog Konfirmasi Sukses Audit
     */
    window.closeModalAuditSuccess = (goToPeserta = false) => {
        const modal = document.getElementById('modalAuditSuccess');
        if (modal) {
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }
        if (goToPeserta && typeof switchTab === 'function') {
            switchTab('daftar-peserta');
        }
    };
}

