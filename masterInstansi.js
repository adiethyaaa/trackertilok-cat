export function toTitleCase(str) {
    if (!str) return '';
    return str.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase());
}

export let masterInstansiData = JSON.parse(localStorage.getItem('master_instansi_pi')) || [
    { name: "Prov. Papua Barat", wilker: "Papua Barat" },
    { name: "Kab. Manokwari", wilker: "Papua Barat" },
    { name: "Kab. Manokwari Selatan", wilker: "Papua Barat" },
    { name: "Kab. Pegunungan Arfak", wilker: "Papua Barat" },
    { name: "Kab. Teluk Bintuni", wilker: "Papua Barat" },
    { name: "Kab. Teluk Wondama", wilker: "Papua Barat" },
    { name: "Kab. Kaimana", wilker: "Papua Barat" },
    { name: "Kab. Fak-fak", wilker: "Papua Barat" },

    { name: "Prov. Papua Barat Daya", wilker: "Papua Barat Daya" },
    { name: "Kota Sorong", wilker: "Papua Barat Daya" },
    { name: "Kab. Sorong", wilker: "Papua Barat Daya" },
    { name: "Kab. Sorong Selatan", wilker: "Papua Barat Daya" },
    { name: "Kab. Raja Ampat", wilker: "Papua Barat Daya" },
    { name: "Kab. Tambrauw", wilker: "Papua Barat Daya" },
    { name: "Kab. Maybrat", wilker: "Papua Barat Daya" },

    { name: "Mahkamah Agung", wilker: "Instansi Vertikal" },
    { name: "Kejaksaan Agung", wilker: "Instansi Vertikal" },
    { name: "Kementerian Hukum Dan HAM", wilker: "Instansi Vertikal" },
    { name: "Kementerian Agama", wilker: "Instansi Vertikal" },
    { name: "Kementerian Keuangan", wilker: "Instansi Vertikal" },
    { name: "Kementerian Kesehatan", wilker: "Instansi Vertikal" },
    { name: "Kementerian Agraria Dan Tata Ruang/BPN", wilker: "Instansi Vertikal" },
    { name: "Badan Pertanahan Nasional", wilker: "Instansi Vertikal" },
    { name: "Badan Pusat Statistik", wilker: "Instansi Vertikal" },
    { name: "Badan Meteorologi, Klimatologi, Dan Geofisika", wilker: "Instansi Vertikal" },
    { name: "Kepolisian Negara Republik Indonesia", wilker: "Instansi Vertikal" }
];

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

export function deleteMasterInstansi(index) {
    if (confirm(`Hapus "${masterInstansiData[index].name}" dari master data?`)) {
        masterInstansiData.splice(index, 1);
        localStorage.setItem('master_instansi_pi', JSON.stringify(masterInstansiData));
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
        formAddMaster.addEventListener('submit', function(e) {
            e.preventDefault();
            const inputName = document.getElementById('newMasterInstansiName');
            const selectWilker = document.getElementById('newMasterInstansiWilker');

            if (!inputName || !selectWilker) return;

            const nameValue = toTitleCase(inputName.value.trim());
            const wilkerValue = selectWilker.value;

            if (!nameValue) {
                alert("Harap masukkan nama instansi!");
                return;
            }

            const exists = masterInstansiData.some(item => item.name.toUpperCase() === nameValue.toUpperCase());
            if (exists) {
                alert(`Instansi "${nameValue}" sudah ada di dalam master data!`);
                return;
            }

            masterInstansiData.push({ name: nameValue, wilker: wilkerValue });
            localStorage.setItem('master_instansi_pi', JSON.stringify(masterInstansiData));

            inputName.value = '';
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

