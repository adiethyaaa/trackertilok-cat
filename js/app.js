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
let currentDashboardDateFilter = 'ALL';
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
 * Event handler saat filter tanggal di Dashboard berubah
 */
window.onDashboardFilterDateChange = (dateVal) => {
    currentDashboardDateFilter = dateVal || 'ALL';
    renderDashboardStats();
};

/**
 * Mengisi opsi dropdown filter tanggal di Dashboard
 */
function populateDashboardFilterTanggalDropdown() {
    const select = document.getElementById('selectDashboardFilterTanggal');
    if (!select) return;

    const uniqueDates = getSortedExamDates();
    const prevVal = currentDashboardDateFilter;

    let html = `<option value="ALL">-- Semua Tanggal Pelaksanaan (${uniqueDates.length} Hari) --</option>`;
    uniqueDates.forEach(d => {
        const fri = isFriday(d);
        html += `<option value="${d}">${d} ${fri ? '(Jumat - Sesi 2: 13.00)' : ''}</option>`;
    });

    select.innerHTML = html;

    if (uniqueDates.includes(prevVal) || prevVal === 'ALL') {
        select.value = prevVal;
    } else {
        select.value = 'ALL';
        currentDashboardDateFilter = 'ALL';
    }
}

/**
 * Render statistik di Dashboard (termasuk Statistik Kehadiran & Filter Tanggal)
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
    const dashboardCandidates = currentDashboardDateFilter === 'ALL'
        ? currentCandidates
        : currentCandidates.filter(c => c.pelaksanaan === currentDashboardDateFilter);

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

    // Render statistik kehadiran per kelompok jabatan (mengikuti filter dashboard yang aktif)
    const kelJabatanContainer = document.getElementById('dashKelJabatanContainer');
    if (kelJabatanContainer) {
        if (total === 0) {
            kelJabatanContainer.innerHTML = `<p class="text-sm text-slate-500 py-6 text-center">Belum ada data peserta untuk ujian/tanggal ini.</p>`;
        } else {
            const byKel = {};
            dashboardCandidates.forEach(c => {
                const rawK = String(c.kelJabatan || '').trim();
                const isNullKel = !rawK || rawK === '-' || rawK === 'NULL';
                const k = isNullKel ? 'Belum Terdata / Kosong' : rawK;
                if (!byKel[k]) byKel[k] = { isNull: isNullKel, hadir: 0, tidakHadir: 0, belum: 0, total: 0 };
                if (c.kehadiran === 'HADIR') byKel[k].hadir++;
                else if (c.kehadiran === 'TIDAK_HADIR') byKel[k].tidakHadir++;
                else byKel[k].belum++;
                byKel[k].total++;
            });

            const sortedKels = Object.keys(byKel).sort((a, b) => byKel[b].total - byKel[a].total);

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
                        ${sortedKels.map(k => {
                            const item = byKel[k];
                            const pct = item.total > 0 ? Math.round((item.hadir / item.total) * 100) : 0;
                            const isWarning = item.isNull;
                            return `
                                <tr class="${isWarning ? 'bg-amber-50/70 text-amber-950 font-medium' : 'hover:bg-slate-50'} transition">
                                    <td class="p-2.5 font-bold ${isWarning ? 'text-amber-800 flex items-center gap-1.5' : 'text-slate-800'}">
                                        ${isWarning ? '<i data-lucide="alert-triangle" class="w-3.5 h-3.5 text-amber-600 inline-block"></i>' : ''}
                                        <span>${k}</span>
                                        ${isWarning ? '<span class="text-[10px] bg-amber-200 text-amber-900 px-1.5 py-0.2 rounded font-bold uppercase">Perlu Dilengkapi</span>' : ''}
                                    </td>
                                    <td class="p-2.5 text-center text-emerald-700 font-bold text-sm">${item.hadir}</td>
                                    <td class="p-2.5 text-center text-rose-700 font-bold text-sm">${item.tidakHadir}</td>
                                    <td class="p-2.5 text-center text-amber-700 font-semibold">${item.belum}</td>
                                    <td class="p-2.5 text-center font-bold text-slate-900">${item.total}</td>
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
                        }).join('')}
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

function populatePelaksanaanFilterDropdown() {
    const select = document.getElementById('selectFilterPelaksanaan');
    if (!select) return;

    const uniqueDates = getSortedExamDates();
    
    let html = `<option value="ALL">-- Semua Tanggal Pelaksanaan (${uniqueDates.length} Tanggal) --</option>`;
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
        ? `-- Semua Sesi (${sortedSessions.length > 0 ? `01 s.d. ${formatCumulativeSessionNumber(sortedSessions[sortedSessions.length - 1].cumNum)}` : '0 Sesi'}) --`
        : `-- Semua Sesi di Tanggal Ini (${sortedSessions.length} Sesi) --`;

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

    let html = `<option value="ALL">-- Semua Kel. Jabatan (${currentCandidates.length}) --</option>`;

    if (countEmpty > 0) {
        html += `<option value="EMPTY">⚠️ [Kosong / Perlu Dilengkapi] (${countEmpty} Peserta)</option>`;
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
window.toggleAttendance = async (candidateId, action) => {
    const cand = currentCandidates.find(c => c.id === candidateId);
    if (!cand) return;

    if (action === 'RESET') {
        cand.kehadiran = null;
    } else {
        cand.kehadiran = action; // 'HADIR' atau 'TIDAK_HADIR'
    }

    try {
        await db.updateCandidate(cand);

        // Sinkronkan ke Firebase Cloud secara Realtime
        if (isCloudActive() && currentExam) {
            updateAttendanceInCloud(currentExam.id, cand.id, cand.kehadiran);
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
}

function renderCandidateListTable() {
    const tbody = document.getElementById('tbodyCandidateList');
    const countBadge = document.getElementById('countTableVisible');
    const paginationInfo = document.getElementById('tablePaginationInfo');

    if (!tbody) return;

    if (countBadge) countBadge.textContent = `${filteredCandidates.length} Data`;
    if (paginationInfo) paginationInfo.textContent = `Menampilkan ${filteredCandidates.length} dari ${currentCandidates.length} total peserta`;

    if (!currentExam) {
        tbody.innerHTML = `
            <tr>
                <td colspan="11" class="p-8 text-center text-slate-400">
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
                <td colspan="11" class="p-8 text-center text-slate-400">
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
                <td colspan="11" class="p-8 text-center text-slate-400">
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
            ? 'bg-amber-50/90 border-l-4 border-l-amber-500 hover:bg-amber-100/80' 
            : (isFriSession2 ? 'bg-amber-50/60' : 'hover:bg-slate-50');

        return `
            <tr class="${rowBgClass} transition">
                <td class="p-3 text-center text-slate-500 font-medium">${idx + 1}</td>
                <td class="p-2.5 text-center whitespace-nowrap">
                    ${c.kehadiran === 'HADIR' ? `
                        <button onclick="toggleAttendance(${c.id}, 'RESET')" class="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-bold bg-emerald-100 text-emerald-800 border border-emerald-300 hover:bg-emerald-200 transition shadow-2xs cursor-pointer" title="Status: Hadir. Klik untuk ubah/batal">
                            <i data-lucide="check" class="w-3.5 h-3.5 stroke-[3]"></i>
                            <span>Hadir</span>
                        </button>
                    ` : c.kehadiran === 'TIDAK_HADIR' ? `
                        <button onclick="toggleAttendance(${c.id}, 'RESET')" class="inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-xs font-bold bg-rose-100 text-rose-800 border border-rose-300 hover:bg-rose-200 transition shadow-2xs cursor-pointer" title="Status: Tidak Hadir. Klik untuk ubah/batal">
                            <i data-lucide="x" class="w-3.5 h-3.5 stroke-[3]"></i>
                            <span>Tidak Hadir</span>
                        </button>
                    ` : `
                        <div class="inline-flex items-center justify-center gap-1.5">
                            <button onclick="toggleAttendance(${c.id}, 'HADIR')" class="p-1.5 rounded-lg bg-emerald-50 hover:bg-emerald-600 hover:text-white text-emerald-600 border border-emerald-300 transition shadow-2xs cursor-pointer" title="Tandai Hadir">
                                <i data-lucide="check" class="w-4 h-4 stroke-[2.5]"></i>
                            </button>
                            <button onclick="toggleAttendance(${c.id}, 'TIDAK_HADIR')" class="p-1.5 rounded-lg bg-rose-50 hover:bg-rose-600 hover:text-white text-rose-600 border border-rose-300 transition shadow-2xs cursor-pointer" title="Tandai Tidak Hadir">
                                <i data-lucide="x" class="w-4 h-4 stroke-[2.5]"></i>
                            </button>
                        </div>
                    `}
                </td>
                <td class="p-3 font-mono font-medium text-slate-900">${c.nip}</td>
                <td class="p-3 font-bold text-slate-900">${c.nama}</td>
                <td class="p-3">
                    ${isKelEmpty ? `
                        <span class="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-bold bg-amber-200 text-amber-900 border border-amber-300 shadow-2xs" title="Kelompok Jabatan belum diisi / kosong! Segera lengkapi">
                            <i data-lucide="alert-triangle" class="w-3 h-3 text-amber-700"></i>
                            <span>Perlu Dilengkapi</span>
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
                <td class="p-3 whitespace-nowrap ${isFriSession2 ? 'font-bold text-amber-800' : 'text-slate-700 font-medium'}">
                    ${(!c.waktu || c.waktu === 'NULL' || c.waktu === '-') ? '<span class="px-2 py-0.5 rounded text-[10px] font-bold bg-slate-100 text-slate-500 border border-slate-200">NULL</span>' : c.waktu}
                </td>
                <td class="p-3 text-center whitespace-nowrap">
                    <button onclick="editCandidate(${c.id})" class="p-1 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded mr-1" title="Edit Data">
                        <i data-lucide="edit-2" class="w-3.5 h-3.5"></i>
                    </button>
                    <button onclick="deleteSingleCandidate(${c.id}, '${c.nama}')" class="p-1 text-rose-600 hover:text-rose-800 hover:bg-rose-50 rounded" title="Hapus Peserta">
                        <i data-lucide="trash-2" class="w-3.5 h-3.5"></i>
                    </button>
                </td>
            </tr>
        `;
    }).join('');

    if (window.lucide) window.lucide.createIcons();
}

window.deleteSingleCandidate = async (candidateId, name) => {
    if (confirm(`Hapus peserta "${name}" dari jadwal ujian?`)) {
        try {
            await db.deleteCandidate(candidateId);
            currentCandidates = currentCandidates.filter(c => c.id !== candidateId);
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

    if (inputDate) inputDate.addEventListener('input', autoCalculateTime);
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
            const pelaksanaan = document.getElementById('inputManualPelaksanaan').value.trim();
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
                    row.id = Number(id);
                    await db.updateCandidate(row);
                    showToast(`Data peserta "${nama}" berhasil diperbarui.`, "success");
                } else {
                    row.no = currentCandidates.length + 1;
                    const newId = await db.addCandidate(row);
                    row.id = newId;
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

    if (currentExam && currentExam.startDate) {
        document.getElementById('inputManualPelaksanaan').value = formatDateDisplay(currentExam.startDate, 'short');
        document.getElementById('inputManualWaktu').value = getSessionTime(1, currentExam.startDate);
    }

    modal.classList.remove('hidden');
    modal.classList.add('flex');
};

window.editCandidate = (candidateId) => {
    const cand = currentCandidates.find(c => c.id === candidateId);
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
    document.getElementById('inputManualPelaksanaan').value = cand.pelaksanaan || '';
    document.getElementById('selectManualSesi').value = cand.sesi || 1;
    document.getElementById('inputManualWaktu').value = cand.waktu || '';

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
 * Cetak Lembar Resmi Jadwal & Daftar Hadir
 */
window.printOfficialSchedule = () => {
    if (!currentExam || filteredCandidates.length === 0) {
        showToast("Tidak ada data peserta untuk dicetak!", "warning");
        return;
    }

    const printContainer = document.getElementById('officialPrintArea');
    if (!printContainer) return;

    const sortedDates = getSortedExamDates();

    // Kelompokkan peserta berdasarkan sesi kumulatif
    const sessionsMap = new Map();
    filteredCandidates.forEach(c => {
        const cum = getCumulativeSessionNumber(c, sortedDates);
        if (!sessionsMap.has(cum)) {
            sessionsMap.set(cum, {
                cumNum: cum,
                cumFormatted: formatCumulativeSessionNumber(cum),
                dailySession: c.sesi,
                date: c.pelaksanaan,
                waktu: c.waktu,
                candidates: []
            });
        }
        sessionsMap.get(cum).candidates.push(c);
    });

    const sortedSessionGroups = Array.from(sessionsMap.values()).sort((a, b) => a.cumNum - b.cumNum);
    const todayStr = formatDateDisplay(new Date(), 'long');

    let fullHtml = '';

    sortedSessionGroups.forEach((group) => {
        // Urutkan nama peserta A-Z dalam setiap sesi
        const sortedList = [...group.candidates].sort((a, b) => 
            String(a.nama || '').localeCompare(String(b.nama || ''), 'id', { sensitivity: 'base' })
        );

        const dayName = group.date ? getDayNameID(group.date) : '';

        fullHtml += `
            <div class="print-session-page">
                <!-- KOP RESMI BKN -->
                <div style="border-bottom: 2px solid #000; padding-bottom: 8px; margin-bottom: 12px; text-align: center;">
                    <div style="font-size: 11pt; font-weight: bold; letter-spacing: 0.5px;">BADAN KEPEGAWAIAN NEGARA</div>
                    <div style="font-size: 10pt; font-weight: bold;">KANTOR REGIONAL XIV MANOKWARI</div>
                    <div style="font-size: 12pt; font-weight: 800; margin-top: 4px; text-decoration: underline;">DAFTAR HADIR & JADWAL PESERTA UJIAN PROFILING ASN</div>
                </div>

                <!-- HEADER KETERANGAN SESI YANG DICETAK -->
                <table style="width: 100%; font-size: 8.5pt; margin-bottom: 8px; border: none;">
                    <tr>
                        <td style="width: 18%; font-weight: bold; padding: 2px 0;">Instansi</td>
                        <td style="width: 2%; padding: 2px 0;">:</td>
                        <td style="width: 45%; font-weight: bold; padding: 2px 0;">${currentExam.instansi}</td>
                        <td style="width: 15%; font-weight: bold; padding: 2px 0;">Sesi Ujian</td>
                        <td style="width: 2%; padding: 2px 0;">:</td>
                        <td style="width: 18%; font-weight: bold; color: #1e3a8a; padding: 2px 0;">Sesi ${group.dailySession} (${group.cumFormatted})</td>
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
                        <td style="padding: 2px 0;">${dayName ? `${dayName}, ` : ''}${group.date || '-'}</td>
                        <td style="font-weight: bold; padding: 2px 0;">Jumlah Peserta</td>
                        <td style="padding: 2px 0;">:</td>
                        <td style="padding: 2px 0; font-weight: bold;">${sortedList.length} Orang</td>
                    </tr>
                </table>

                <!-- TABEL PESERTA SESI INI (Kolom Waktu diganti Kolom Sesi & Kumulatif) -->
                <table class="print-table">
                    <thead>
                        <tr>
                            <th style="width: 28px;">No</th>
                            <th style="width: 125px;">NIP</th>
                            <th>Nama Peserta</th>
                            <th>Unit Kerja / Jabatan</th>
                            <th style="width: 100px;">Sesi</th>
                            <th style="width: 115px;">Tanda Tangan</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${sortedList.map((c, idx) => {
                            const statusBadge = c.kehadiran === 'HADIR'
                                ? '<span style="color: #047857; font-weight: bold;">[ HADIR ]</span>'
                                : (c.kehadiran === 'TIDAK_HADIR'
                                    ? '<span style="color: #b91c1c; font-weight: bold;">[ TDK HADIR ]</span>'
                                    : '');
                            return `
                                <tr>
                                    <td style="text-align: center;">${idx + 1}</td>
                                    <td style="font-family: monospace; text-align: center;">${c.nip}</td>
                                    <td style="font-weight: bold;">${c.nama}</td>
                                    <td>${c.unitKerja || '-'}${c.jabatan ? `<br><span style="font-size: 7.5pt; color: #444;">${c.jabatan}</span>` : ''}</td>
                                    <td style="text-align: center; font-weight: bold;">Sesi ${c.sesi} (${group.cumFormatted})</td>
                                    <td style="height: 24px; vertical-align: middle;">
                                        ${statusBadge || `<span style="color: #888;">${idx + 1}. .........</span>`}
                                    </td>
                                </tr>
                            `;
                        }).join('')}
                    </tbody>
                </table>

                <!-- TANDA TANGAN PENANGGUNG JAWAB -->
                <div style="margin-top: 18px; display: flex; justify-content: flex-end; page-break-inside: avoid;">
                    <div style="width: 240px; text-align: center; font-size: 8.5pt;">
                        <div>Manokwari, ${todayStr}</div>
                        <div style="margin-top: 4px; font-weight: bold;">Koordinator Tim Pelaksana CAT BKN,</div>
                        <div style="height: 48px;"></div>
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
