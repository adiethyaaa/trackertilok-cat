import * as db from './js/db.js';

export function toTitleCase(str) {
    if (!str) return '';
    return str.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
}

export const DEFAULT_MASTER_INSTANSI = [
    { name: "Prov. Papua Barat", wilker: "Papua Barat", pin: "papuabarat" },
    { name: "Kab. Manokwari", wilker: "Papua Barat", pin: "manokwari" },
    { name: "Kab. Manokwari Selatan", wilker: "Papua Barat", pin: "mansel" },
    { name: "Kab. Pegunungan Arfak", wilker: "Papua Barat", pin: "arfak" },
    { name: "Kab. Teluk Bintuni", wilker: "Papua Barat", pin: "bintuni" },
    { name: "Kab. Teluk Wondama", wilker: "Papua Barat", pin: "wondama" },
    { name: "Kab. Kaimana", wilker: "Papua Barat", pin: "kaimana" },
    { name: "Kab. Fak-Fak", wilker: "Papua Barat", pin: "fakfak" },

    { name: "Prov. Papua Barat Daya", wilker: "Papua Barat Daya", pin: "pbd" },
    { name: "Kota Sorong", wilker: "Papua Barat Daya", pin: "kota" },
    { name: "Kab. Sorong", wilker: "Papua Barat Daya", pin: "kabsor" },
    { name: "Kab. Sorong Selatan", wilker: "Papua Barat Daya", pin: "sorsel" },
    { name: "Kab. Raja Ampat", wilker: "Papua Barat Daya", pin: "raja4" },
    { name: "Kab. Tambrauw", wilker: "Papua Barat Daya", pin: "tambrauw" },
    { name: "Kab. Maybrat", wilker: "Papua Barat Daya", pin: "maybrat" },

    { name: "Mahkamah Agung", wilker: "Instansi Vertikal", pin: "1414" },
    { name: "Kejaksaan Agung", wilker: "Instansi Vertikal", pin: "1414" },
    { name: "Kementerian Hukum Dan HAM", wilker: "Instansi Vertikal", pin: "1414" },
    { name: "Kementerian Agama", wilker: "Instansi Vertikal", pin: "1414" },
    { name: "Kementerian Keuangan", wilker: "Instansi Vertikal", pin: "1414" },
    { name: "Kementerian Kesehatan", wilker: "Instansi Vertikal", pin: "1414" },
    { name: "Kementerian Agraria Dan Tata Ruang/BPN", wilker: "Instansi Vertikal", pin: "1414" },
    { name: "Badan Pertanahan Nasional", wilker: "Instansi Vertikal", pin: "1414" },
    { name: "Badan Pusat Statistik", wilker: "Instansi Vertikal", pin: "1414" },
    { name: "Badan Meteorologi, Klimatologi, Dan Geofisika", wilker: "Instansi Vertikal", pin: "1414" },
    { name: "Kepolisian Negara Republik Indonesia", wilker: "Instansi Vertikal", pin: "1414" }
];

export let masterInstansiData = (() => {
    try {
        const stored = localStorage.getItem('master_instansi_pi');
        if (stored) {
            const parsed = JSON.parse(stored);
            if (Array.isArray(parsed) && parsed.length > 0) {
                return parsed;
            }
        }
    } catch (e) {
        console.warn("Gagal parse master_instansi_pi:", e);
    }
    return [...DEFAULT_MASTER_INSTANSI];
})();

/**
 * Sinkronisasi data master instansi dari Database
 */
export async function syncMasterInstansiFromDatabase() {
    try {
        const dbList = await db.getMasterInstansiFromDb();
        if (dbList && Array.isArray(dbList) && dbList.length > 0) {
            masterInstansiData.length = 0;
            dbList.forEach(item => masterInstansiData.push(item));
            localStorage.setItem('master_instansi_pi', JSON.stringify(masterInstansiData));
        }
    } catch (e) {
        console.warn("Gagal sync master instansi dari database:", e);
    }
    return masterInstansiData;
}

/**
 * Helper untuk mendapatkan PIN resmi dari nama instansi (Database Driven)
 */
export function getInstansiPin(instansiName) {
    if (!instansiName) return '';
    const clean = String(instansiName).trim().toLowerCase();

    // 1. Cek dari masterInstansiData di memori
    const found = masterInstansiData.find(i => String(i.name || '').trim().toLowerCase() === clean);
    if (found && found.pin) return String(found.pin).trim();

    return '';
}

export function renderMasterInstansiTable() {
    const tbody = document.getElementById('tbodyMasterInstansi');
    if (!tbody) return;

    if (masterInstansiData.length === 0) {
        tbody.innerHTML = `<tr><td colspan="3" style="text-align: center; padding: 12px; color: #94a3b8;">Belum ada data master instansi.</td></tr>`;
        return;
    }

    const vertikalList = masterInstansiData.filter(item => item.wilker === 'Instansi Vertikal');
    const nonVertikalList = masterInstansiData.filter(item => item.wilker !== 'Instansi Vertikal');

    vertikalList.sort((a, b) => a.name.localeCompare(b.name, 'id', { sensitivity: 'base' }));

    const sortedData = [...nonVertikalList, ...vertikalList];

    tbody.innerHTML = sortedData.map((item) => {
        const originalIndex = masterInstansiData.findIndex(orig => orig.name === item.name && orig.wilker === item.wilker);

        return `
            <tr>
                <td style="padding: 7px 10px; border-bottom: 1px solid #e2e8f0; font-weight: 600; color: #1e293b;">${item.name}</td>
                <td style="padding: 7px 10px; border-bottom: 1px solid #e2e8f0;">
                    <span style="font-size: 10px; font-weight: 700; padding: 2px 8px; border-radius: 4px; background: #e0f2fe; color: #0369a1;">${item.wilker}</span>
                </td>
                <td style="padding: 7px 10px; border-bottom: 1px solid #e2e8f0; text-align: center;">
                    <button type="button" onclick="deleteMasterInstansi(${originalIndex})" style="background: #fee2e2; border: 1px solid #fca5a5; color: #dc2626; border-radius: 4px; padding: 2px 8px; font-size: 11px; font-weight: bold; cursor: pointer;" title="Hapus Instansi">Hapus</button>
                </td>
            </tr>
        `;
    }).join('');
}

export function openModalMasterInstansi() {
    const modalWrapper = document.getElementById('modalMasterInstansi');
    if (modalWrapper) {
        modalWrapper.style.display = 'flex';
        
        const modalContent = modalWrapper.querySelector('.modal-content');
        if (modalContent) {
            modalContent.style.display = 'flex';
        }

        document.body.style.overflow = 'hidden';
        renderMasterInstansiTable();
    }
}

export function closeModalMasterInstansi() {
    const modalWrapper = document.getElementById('modalMasterInstansi');
    if (modalWrapper) {
        modalWrapper.style.display = 'none';
        
        const modalContent = modalWrapper.querySelector('.modal-content');
        if (modalContent) {
            modalContent.style.display = 'none';
        }

        document.body.style.overflow = 'auto';
    }
}

export async function deleteMasterInstansi(index) {
    if (confirm(`Hapus "${masterInstansiData[index].name}" dari master data?`)) {
        masterInstansiData.splice(index, 1);
        localStorage.setItem('master_instansi_pi', JSON.stringify(masterInstansiData));
        await db.saveMasterInstansiToDb(masterInstansiData);
        renderMasterInstansiTable();
        if (typeof window.renderInstansiDatalist === 'function') {
            window.renderInstansiDatalist();
        }
    }
}

export function setupMasterInstansiForm() {
    const modalMaster = document.getElementById('modalMasterInstansi');
    if (modalMaster) {
        modalMaster.addEventListener('click', function(event) {
            if (event.target === this) {
                closeModalMasterInstansi();
            }
        });
    }

    const formAddMaster = document.getElementById('formAddMasterInstansi');
    if (formAddMaster) {
        formAddMaster.addEventListener('submit', async function(e) {
            e.preventDefault();
            const inputName = document.getElementById('newMasterInstansiName');
            const selectWilker = document.getElementById('newMasterInstansiWilker');
            const inputPin = document.getElementById('newMasterInstansiPin');

            if (!inputName || !selectWilker) return;

            const nameValue = toTitleCase(inputName.value.trim());
            const wilkerValue = selectWilker.value;
            const pinValue = inputPin ? inputPin.value.trim() : '';

            if (!nameValue) {
                alert("Harap masukkan nama instansi!");
                return;
            }

            const exists = masterInstansiData.some(item => item.name.toUpperCase() === nameValue.toUpperCase());
            if (exists) {
                alert(`Instansi "${nameValue}" sudah ada di dalam master data!`);
                return;
            }

            masterInstansiData.push({ 
                name: nameValue, 
                wilker: wilkerValue,
                pin: pinValue || nameValue.toLowerCase().replace(/[^a-z0-9]/g, '')
            });
            localStorage.setItem('master_instansi_pi', JSON.stringify(masterInstansiData));
            await db.saveMasterInstansiToDb(masterInstansiData);

            inputName.value = '';
            if (inputPin) inputPin.value = '';
            renderMasterInstansiTable();
            if (typeof window.renderInstansiDatalist === 'function') {
                window.renderInstansiDatalist();
            }
            alert(`✅ Instansi "${nameValue}" berhasil ditambahkan!`);
        });
    }
}

export function renderInstansiDatalist() {
    const datalist = document.getElementById('listInstansiSuggest');
    if (datalist && typeof masterInstansiData !== 'undefined') {
        datalist.innerHTML = masterInstansiData.map(item => `<option value="${item.name}"></option>`).join('');
    }
    const datalistPGA = document.getElementById('listInstansiSuggestPGA');
    if (datalistPGA && typeof masterInstansiData !== 'undefined') {
        const daerahOnly = masterInstansiData.filter(i => i.wilker !== 'Instansi Vertikal');
        datalistPGA.innerHTML = daerahOnly.map(item => `<option value="${item.name}"></option>`).join('');
    }
}

// Pasang ke window agar onclick HTML bisa mengakses
if (typeof window !== 'undefined') {
    window.masterInstansiData = masterInstansiData;
    window.renderMasterInstansiTable = renderMasterInstansiTable;
    window.openModalMasterInstansi = openModalMasterInstansi;
    window.closeModalMasterInstansi = closeModalMasterInstansi;
    window.deleteMasterInstansi = deleteMasterInstansi;
    window.renderInstansiDatalist = renderInstansiDatalist;
}

