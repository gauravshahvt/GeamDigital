// Geam Digital - Multi-Ward Panchayat Voter List to Excel Client Application Logic (with Status Column)

document.addEventListener('DOMContentLoaded', () => {
    // Application State
    const state = {
        sessionId: null,
        fileName: 'Geam_Digital_Panchayat_Voter_List.xlsx',
        summaryGrid: [],     // Sheet 1: वार्ड_सारांश
        masterGrid: [],      // Sheet 2: समस्त_मतदाता_सूची (All Wards Sequence-wise with 13 columns)
        wards: [],           // [{ part_no, filename, active_count, deleted_count, grid }]
        allTables: [],       // Array of all tables for openpyxl
        allTableNames: [],   // Array of sheet names
        currentView: 'master', // 'master' or ward index (0, 1...)
        statusFilter: 'all', // 'all', 'Active', 'Deleted'
        activeGrid: [],      // Currently selected grid in view
        filteredGrid: [],    // Filtered rows for live search & status
        metadata: {}
    };

    // DOM Elements
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const selectedFilesBox = document.getElementById('selectedFilesBox');
    const filesCountText = document.getElementById('filesCountText');
    const filesListBadges = document.getElementById('filesListBadges');
    const quickSampleVoterBtn = document.getElementById('quickSampleVoterBtn');
    const uploadForm = document.getElementById('uploadForm');
    const convertBtn = document.getElementById('convertBtn');
    const convertBtnText = document.getElementById('convertBtnText');
    const autoDownloadCheck = document.getElementById('autoDownloadCheck');
    const loadingIndicator = document.getElementById('loadingIndicator');
    const loadingStatusText = document.getElementById('loadingStatusText');
    const resultsWorkspace = document.getElementById('resultsWorkspace');

    // Live Streaming Progress & Preview Elements
    const liveProgressContainer = document.getElementById('liveProgressContainer');
    const liveWardBadge = document.getElementById('liveWardBadge');
    const liveStatusText = document.getElementById('liveStatusText');
    const liveFileSubtitle = document.getElementById('liveFileSubtitle');
    const liveVotersCount = document.getElementById('liveVotersCount');
    const liveProgressPercentText = document.getElementById('liveProgressPercentText');
    const liveProgressBar = document.getElementById('liveProgressBar');
    const liveSpreadsheetBody = document.getElementById('liveSpreadsheetBody');

    // History System Elements
    const historyModalBtn = document.getElementById('historyModalBtn');
    const historyCountBadge = document.getElementById('historyCountBadge');
    const historyDrawerBackdrop = document.getElementById('historyDrawerBackdrop');
    const historyDrawerPanel = document.getElementById('historyDrawerPanel');
    const closeHistoryDrawerBtn = document.getElementById('closeHistoryDrawerBtn');
    const clearHistoryBtn = document.getElementById('clearHistoryBtn');
    const historyListContainer = document.getElementById('historyListContainer');
    const historyFilterBtns = document.querySelectorAll('.history-filter-btn');

    // Hero Download Card Elements
    const quickStatsBadge = document.getElementById('quickStatsBadge');
    const quickDownloadTitle = document.getElementById('quickDownloadTitle');
    const heroDownloadBtn = document.getElementById('heroDownloadBtn');

    // Summary Section
    const summaryGridTable = document.getElementById('summaryGridTable');
    const summaryWardsCountBadge = document.getElementById('summaryWardsCountBadge');

    // Master 11-Column Spreadsheet Studio Elements
    const wardViewSelect = document.getElementById('wardViewSelect');
    const searchInput = document.getElementById('searchInput');
    const tableInfoBadge = document.getElementById('tableInfoBadge');
    const editableGrid = document.getElementById('editableGrid');
    const copyTsvBtn = document.getElementById('copyTsvBtn');

    // Status Filter Buttons
    const filterStatusAll = document.getElementById('filterStatusAll');
    const filterStatusActive = document.getElementById('filterStatusActive');
    const filterStatusDeleted = document.getElementById('filterStatusDeleted');

    // Export Controls
    const themeSelect = document.getElementById('themeSelect');
    const exportExcelBtn = document.getElementById('exportExcelBtn');
    const exportCsvBtn = document.getElementById('exportCsvBtn');

    // Tab Switcher Elements
    const tabFeature1Btn = document.getElementById('tabFeature1Btn');
    const tabFeature2Btn = document.getElementById('tabFeature2Btn');
    const tabFeature3Btn = document.getElementById('tabFeature3Btn');
    const feature1Container = document.getElementById('feature1Container');
    const feature2Container = document.getElementById('feature2Container');
    const feature3Container = document.getElementById('feature3Container');
    const feature4Container = document.getElementById('feature4Container');
    const tabFeature4Btn = document.getElementById('tabFeature4Btn');
    const tabFeature5Btn = document.getElementById('tabFeature5Btn');
    const tabFeature6Btn = document.getElementById('tabFeature6Btn');
    const featureActiveDesc = document.getElementById('featureActiveDesc');
    const goToParchiHeroBtn = document.getElementById('goToParchiHeroBtn');
    const goToColorParchiHeroBtn = document.getElementById('goToColorParchiHeroBtn');
    const footerParchiBtn = document.getElementById('footerParchiBtn');

    function switchTab(tabId) {
        const containers = [
            { id: 'feature1', el: feature1Container, btn: tabFeature1Btn, activeClass: 'text-emerald-800' },
            { id: 'feature2', el: feature2Container, btn: tabFeature2Btn, activeClass: 'text-indigo-800' },
            { id: 'feature3', el: feature3Container, btn: tabFeature3Btn, activeClass: 'text-rose-800' },
            { id: 'feature4', el: feature4Container, btn: tabFeature4Btn, activeClass: 'text-amber-900' },
            { id: 'feature5', el: document.getElementById('feature5Container'), btn: tabFeature5Btn, activeClass: 'text-purple-900' },
            { id: 'feature6', el: document.getElementById('feature6Container'), btn: tabFeature6Btn, activeClass: 'text-teal-900' }
        ];

        containers.forEach(c => {
            if (c.el) {
                if (c.id === tabId) {
                    c.el.classList.remove('hidden');
                } else {
                    c.el.classList.add('hidden');
                }
            }
            if (c.btn) {
                if (c.id === tabId) {
                    c.btn.className = `flex items-center gap-2 px-5 py-2 rounded-lg text-xs font-extrabold transition shadow-xs bg-white ${c.activeClass} border border-slate-200 cursor-pointer`;
                } else {
                    c.btn.className = 'flex items-center gap-2 px-5 py-2 rounded-lg text-xs font-extrabold transition text-slate-600 hover:text-slate-900 hover:bg-white/70 cursor-pointer';
                }
            }
        });

        if (featureActiveDesc) {
            if (tabId === 'feature1') featureActiveDesc.textContent = 'PDF से 13-कॉलम एक्सेल रूपांतरण';
            else if (tabId === 'feature2') featureActiveDesc.textContent = 'A4 शीट पर 12 Plain मतदाता पर्ची';
            else if (tabId === 'feature3') featureActiveDesc.textContent = 'A4 शीट पर कलर मतदाता पर्ची (प्रत्याशी पोस्टर सहित)';
            else if (tabId === 'feature4') featureActiveDesc.textContent = 'A, B, C, D अनुसार वर्णमाला (Alphabetical) वोटर लिस्ट';
            else if (tabId === 'feature5') featureActiveDesc.textContent = 'समान मकान नंबर वार बड़े परिवारों की लिस्ट (अवरोही क्रम)';
            else if (tabId === 'feature6') featureActiveDesc.textContent = 'आयु अनुसार युवा (18-26 बढ़ते क्रम) एवं बुजुर्ग (120-70 घटते क्रम) लिस्ट';
        }

        if (window.lucide) lucide.createIcons();
    }

    if (tabFeature1Btn) tabFeature1Btn.addEventListener('click', () => switchTab('feature1'));
    if (tabFeature2Btn) tabFeature2Btn.addEventListener('click', () => switchTab('feature2'));
    if (tabFeature3Btn) tabFeature3Btn.addEventListener('click', () => switchTab('feature3'));
    if (tabFeature4Btn) tabFeature4Btn.addEventListener('click', () => switchTab('feature4'));
    if (tabFeature5Btn) tabFeature5Btn.addEventListener('click', () => switchTab('feature5'));
    if (tabFeature6Btn) tabFeature6Btn.addEventListener('click', () => switchTab('feature6'));

    if (goToParchiHeroBtn) {
        goToParchiHeroBtn.addEventListener('click', () => {
            switchTab('feature2');
            loadParchiFromFeature1Session();
        });
    }
    if (goToColorParchiHeroBtn) {
        goToColorParchiHeroBtn.addEventListener('click', () => {
            switchTab('feature3');
            loadColorParchiFromFeature1Session();
        });
    }
    if (footerParchiBtn) {
        footerParchiBtn.addEventListener('click', () => {
            switchTab('feature2');
            loadParchiFromFeature1Session();
        });
    }

    // 1. Drag & Drop & Multiple File Selection
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('border-emerald-500', 'bg-emerald-50/50');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('border-emerald-500', 'bg-emerald-50/50');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('border-emerald-500', 'bg-emerald-50/50');
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            fileInput.files = e.dataTransfer.files;
            handleFilesChosen(fileInput.files);
        }
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files && fileInput.files.length > 0) {
            handleFilesChosen(fileInput.files);
        }
    });

    function handleFilesChosen(files) {
        const pdfFiles = Array.from(files).filter(f => f.name.toLowerCase().endsWith('.pdf'));
        if (pdfFiles.length === 0) {
            alert('कृपया केवल वैध .pdf मतदाता सूची फ़ाइलें चुनें');
            return;
        }

        filesCountText.textContent = `${pdfFiles.length} फ़ाइलें चयनित (${pdfFiles.length} Wards Selected)`;
        filesListBadges.innerHTML = '';

        pdfFiles.forEach(file => {
            const badge = document.createElement('span');
            badge.className = 'text-[11px] font-semibold bg-white border border-emerald-300 text-slate-800 px-2.5 py-0.5 rounded-md shadow-2xs truncate max-w-xs';
            badge.textContent = `${file.name} (${(file.size / 1024).toFixed(0)} KB)`;
            filesListBadges.appendChild(badge);
        });

        selectedFilesBox.classList.remove('hidden');
        selectedFilesBox.classList.add('flex');

        if (pdfFiles.length === 1) {
            convertBtnText.textContent = `अपलोड करें और "${pdfFiles[0].name}" की 13-कॉलम एक्सेल पाएं`;
        } else {
            convertBtnText.textContent = `सभी ${pdfFiles.length} वार्ड प्रोसेस करें और संपूर्ण पंचायत एक्सेल पाएं`;
        }

        if (window.lucide) lucide.createIcons();
    }

    // 2. Quick Demo Button Click (Loads 3 Wards: Ward 1, Ward 2, Ward 7)
    if (quickSampleVoterBtn) {
        quickSampleVoterBtn.addEventListener('click', async () => {
            setLoading(true, 'डेमो पंचायत (वार्ड 1, वार्ड 2, वार्ड 7) लोड की जा रही है...');
            try {
                const response = await fetch('/api/load_sample', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ sample_type: 'multi_ward' })
                });
                const resData = await response.json();
                if (!response.ok) throw new Error(resData.error || 'Failed to load sample');

                handleBatchResponse(resData);
            } catch (err) {
                alert(`Error: ${err.message}`);
            } finally {
                setLoading(false);
            }
        });
    }

    // 3. Form Submit - Upload & Process Multi-Ward PDFs with Live Preview & Real-Time Row Streaming
    let activeUploadPollTimer = null;

    uploadForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        if (!fileInput.files || fileInput.files.length === 0) {
            alert('कृपया पहले एक या अधिक वोटर लिस्ट पीडीएफ फ़ाइलें चुनें!');
            return;
        }

        const formData = new FormData();
        const fileNames = [];
        for (let i = 0; i < fileInput.files.length; i++) {
            formData.append('pdf_files', fileInput.files[i]);
            fileNames.push(fileInput.files[i].name);
        }

        const count = fileInput.files.length;

        // Hide previous results and old spinner
        if (resultsWorkspace) resultsWorkspace.classList.add('hidden');
        if (loadingIndicator) loadingIndicator.classList.add('hidden');

        // Reveal and initialize live progress container
        if (liveProgressContainer) {
            liveProgressContainer.classList.remove('hidden');
            if (liveSpreadsheetBody) liveSpreadsheetBody.innerHTML = '';
            if (liveProgressBar) liveProgressBar.style.width = '0%';
            if (liveProgressPercentText) liveProgressPercentText.textContent = '0%';
            if (liveVotersCount) liveVotersCount.textContent = '0';
            if (liveWardBadge) liveWardBadge.textContent = `वार्ड 1/${count}`;
            if (liveStatusText) liveStatusText.textContent = `सभी ${count} वार्डों की पीडीएफ अपलोड एवं विश्लेषण शुरू...`;
            if (liveFileSubtitle) liveFileSubtitle.textContent = fileNames.join(', ');
            liveProgressContainer.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        }

        convertBtn.disabled = true;
        if (quickSampleVoterBtn) quickSampleVoterBtn.disabled = true;

        try {
            const startResp = await fetch('/api/upload_start', {
                method: 'POST',
                body: formData
            });
            const startData = await startResp.json();
            if (!startResp.ok) throw new Error(startData.error || 'अपलोड प्रक्रिया शुरू नहीं हो सकी');

            const taskId = startData.task_id;
            const seenVoterKeys = new Set();

            if (activeUploadPollTimer) clearInterval(activeUploadPollTimer);

            activeUploadPollTimer = setInterval(async () => {
                try {
                    const progResp = await fetch(`/api/upload_progress/${taskId}`);
                    if (!progResp.ok) return;
                    const progData = await progResp.json();

                    // Update UI stats
                    if (liveWardBadge) liveWardBadge.textContent = `वार्ड ${progData.current_ward || 1}/${progData.total_wards || count}`;
                    if (liveStatusText) liveStatusText.textContent = progData.message || 'मतदाता डेटा स्कैन एवं एक्सट्रैक्शन जारी...';
                    if (liveFileSubtitle && progData.current_file) liveFileSubtitle.textContent = `फ़ाइल: ${progData.current_file}`;
                    if (liveProgressBar) liveProgressBar.style.width = `${progData.percent || 0}%`;
                    if (liveProgressPercentText) liveProgressPercentText.textContent = `${progData.percent || 0}%`;
                    if (liveVotersCount) liveVotersCount.textContent = progData.total_voters || 0;

                    // Stream recent extracted voters live into table
                    if (liveSpreadsheetBody && Array.isArray(progData.recent_voters)) {
                        progData.recent_voters.forEach(v => {
                            const key = `${v['वार्ड नं.']}_${v['क्रम संख्या']}_${v['नाम']}`;
                            if (!seenVoterKeys.has(key)) {
                                seenVoterKeys.add(key);
                                const tr = document.createElement('tr');
                                tr.className = 'border-b border-slate-100 hover:bg-emerald-50/50 bg-emerald-50/20 transition-all duration-300';

                                const isDel = String(v['Status'] || '').toLowerCase() === 'deleted';
                                tr.innerHTML = `
                                    <td class="py-1.5 px-3 border-r border-slate-200 text-center font-bold text-sky-800">${v['जि. प.'] || '-'}</td>
                                    <td class="py-1.5 px-3 border-r border-slate-200 text-center font-bold text-sky-800">${v['पं. स.'] || '-'}</td>
                                    <td class="py-1.5 px-3 border-r border-slate-200 font-bold text-slate-700">${v['वार्ड नं.'] || '-'}</td>
                                    <td class="py-1.5 px-3 border-r border-slate-200 font-mono text-slate-800">${v['क्रम संख्या'] || '-'}</td>
                                    <td class="py-1.5 px-3 border-r border-slate-200 font-bold text-slate-900">${v['नाम'] || '-'}</td>
                                    <td class="py-1.5 px-3 border-r border-slate-200 text-slate-700">${v['पिता/पति का नाम'] || '-'}</td>
                                    <td class="py-1.5 px-3 border-r border-slate-200 text-center text-slate-700">${v['आयु'] || '-'}</td>
                                    <td class="py-1.5 px-3 border-r border-slate-200 text-center text-slate-700">${v['लिंग'] || '-'}</td>
                                    <td class="py-1.5 px-3 border-r border-slate-200 text-slate-700">${v['हाउस नंबर'] || '-'}</td>
                                    <td class="py-1.5 px-3 text-center">
                                        <span class="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-extrabold ${isDel ? 'bg-rose-100 text-rose-800 border border-rose-200' : 'bg-emerald-100 text-emerald-800 border border-emerald-200'}">
                                            <span class="w-1.5 h-1.5 rounded-full ${isDel ? 'bg-rose-600' : 'bg-emerald-600'}"></span>
                                            ${v['Status'] || 'Active'}
                                        </span>
                                    </td>
                                `;
                                liveSpreadsheetBody.appendChild(tr);

                                // Auto-scroll to bottom of live stream table
                                const scrollContainer = liveSpreadsheetBody.closest('.overflow-y-auto');
                                if (scrollContainer) scrollContainer.scrollTop = scrollContainer.scrollHeight;
                            }
                        });
                    }

                    // Check completion
                    if (progData.status === 'completed') {
                        clearInterval(activeUploadPollTimer);
                        activeUploadPollTimer = null;
                        if (liveProgressBar) liveProgressBar.style.width = '100%';
                        if (liveProgressPercentText) liveProgressPercentText.textContent = '100%';
                        if (liveStatusText) liveStatusText.textContent = '✅ सम्पूर्ण 13-कॉलम एक्सेल तैयार हो गई!';

                        setTimeout(() => {
                            if (liveProgressContainer) liveProgressContainer.classList.add('hidden');
                            handleBatchResponse(progData);
                            refreshHistoryPills();
                            convertBtn.disabled = false;
                            if (quickSampleVoterBtn) quickSampleVoterBtn.disabled = false;
                        }, 700);
                    } else if (progData.status === 'error') {
                        clearInterval(activeUploadPollTimer);
                        activeUploadPollTimer = null;
                        if (liveProgressContainer) liveProgressContainer.classList.add('hidden');
                        alert(`प्रक्रिया में त्रुटि: ${progData.error || 'अज्ञात त्रुटि'}`);
                        convertBtn.disabled = false;
                        if (quickSampleVoterBtn) quickSampleVoterBtn.disabled = false;
                    }
                } catch (err) {
                    console.warn('Poll error:', err);
                }
            }, 350);

        } catch (err) {
            if (activeUploadPollTimer) clearInterval(activeUploadPollTimer);
            if (liveProgressContainer) liveProgressContainer.classList.add('hidden');
            alert(`अपलोड त्रुटि: ${err.message}`);
            convertBtn.disabled = false;
            if (quickSampleVoterBtn) quickSampleVoterBtn.disabled = false;
        }
    });

    // 4. Handle Server Response for Multi-Ward Batch
    function handleBatchResponse(resData) {
        state.sessionId = resData.session_id;
        state.fileName = resData.file_name || (resData.data && (resData.data.suggested_filename || resData.data.file_name)) || 'Bagmali Ward 1 to 7.xlsx';
        const data = resData.data || {};

        state.summaryGrid = data.summary_grid || [];
        state.masterGrid = data.master_grid || [];
        state.wards = data.wards || [];
        state.allTables = data.tables || [state.summaryGrid, state.masterGrid];
        state.allTableNames = data.table_names || ['वार्ड_सारांश', 'समस्त_मतदाता_सूची'];
        state.metadata = data.invoice_metadata || {};

        // Default view: Master Sequence Sheet with 13 columns
        state.currentView = 'master';
        state.statusFilter = 'all';
        updateStatusFilterButtons();
        state.activeGrid = state.masterGrid;
        state.filteredGrid = [...state.activeGrid];

        // Update UI
        resultsWorkspace.classList.remove('hidden');
        renderSummaryTable();
        populateWardSelector();
        renderActiveTable();

        // Update badges & hero text
        const totalWards = data.total_wards || state.wards.length || 1;
        const totalActive = data.total_active_voters || 0;
        const totalDeleted = data.total_deleted_voters || 0;
        const totalScanned = data.total_scanned_voters || (totalActive + totalDeleted);

        quickStatsBadge.textContent = `${totalWards} वार्ड्स • ${totalScanned} कुल मतदाता (${totalActive} Active, ${totalDeleted} Deleted)`;
        quickDownloadTitle.textContent = `${state.fileName} तैयार है! (${totalWards} वार्ड्स, ${totalScanned} रिकॉर्ड्स)`;
        summaryWardsCountBadge.textContent = `${totalWards} वार्ड्स सम्मिलित (कुल ${totalScanned}: ${totalActive} Active, ${totalDeleted} Deleted)`;

        // Auto Download if checked
        if (autoDownloadCheck && autoDownloadCheck.checked) {
            setTimeout(() => {
                triggerExcelDownload();
            }, 600);
        }

        // Scroll to results
        resultsWorkspace.scrollIntoView({ behavior: 'smooth' });
        refreshHistoryPills();
        if (window.lucide) lucide.createIcons();
    }

    // 5. Render Sheet 1: वार्ड_सारांश Table
    function renderSummaryTable() {
        summaryGridTable.innerHTML = '';
        if (!state.summaryGrid || state.summaryGrid.length === 0) return;

        const headers = state.summaryGrid[0];
        const rows = state.summaryGrid.slice(1);

        // Header
        const thead = document.createElement('thead');
        thead.className = 'bg-slate-900 text-white';
        const trHead = document.createElement('tr');
        headers.forEach(h => {
            const th = document.createElement('th');
            th.className = 'p-2.5 text-center border border-slate-700 text-xs font-bold whitespace-nowrap bg-slate-900';
            th.textContent = h;
            trHead.appendChild(th);
        });
        thead.appendChild(trHead);
        summaryGridTable.appendChild(thead);

        // Body
        const tbody = document.createElement('tbody');
        tbody.className = 'divide-y divide-slate-200';

        rows.forEach((row, rowIdx) => {
            const tr = document.createElement('tr');
            const isTotalRow = (rowIdx === rows.length - 1) && String(row[0] || '').includes('कुल');

            if (isTotalRow) {
                tr.className = 'bg-sky-50 font-bold border-t-2 border-slate-400 text-slate-900';
            } else {
                tr.className = rowIdx % 2 === 0 ? 'bg-white hover:bg-emerald-50/40' : 'bg-slate-50/70 hover:bg-emerald-50/40';
            }

            row.forEach((cellVal, colIdx) => {
                const td = document.createElement('td');
                td.className = 'p-2 border border-slate-200 text-xs text-slate-800 whitespace-nowrap';

                const hName = headers[colIdx] || '';
                if (['क्र.सं.', 'वार्ड नं.', 'वार्ड संख्या', 'वार्ड / भाग संख्या', 'कुल मतदाता', 'सक्रिय (Active)', 'विलोपित (Deleted)'].includes(hName)) {
                    td.classList.add('text-center', 'font-semibold');
                }
                if (hName === 'सक्रिय (Active)') {
                    td.classList.add('text-emerald-700', 'font-bold');
                }
                if (hName === 'विलोपित (Deleted)') {
                    td.classList.add('text-rose-600', 'font-bold');
                }
                if (isTotalRow) {
                    td.classList.add('text-slate-900', 'font-bold');
                }

                td.textContent = cellVal !== null && cellVal !== undefined ? cellVal : '';
                tr.appendChild(td);
            });

            tbody.appendChild(tr);
        });

        summaryGridTable.appendChild(tbody);
    }

    // 6. Populate Ward Selector Dropdown
    function populateWardSelector() {
        wardViewSelect.innerHTML = '';

        // Master Option
        const totalVoters = Math.max(0, state.masterGrid.length - 1);
        const optMaster = document.createElement('option');
        optMaster.value = 'master';
        optMaster.textContent = `📌 समस्त वार्ड (संयुक्त अनुक्रम - Master Sheet: ${totalVoters} मतदाता)`;
        wardViewSelect.appendChild(optMaster);

        // Individual Wards
        state.wards.forEach((w, idx) => {
            const opt = document.createElement('option');
            opt.value = String(idx);
            opt.textContent = `वार्ड ${w.part_no} (${w.total_scanned || w.active_count} मतदाता - ${w.filename})`;
            wardViewSelect.appendChild(opt);
        });

        wardViewSelect.value = 'master';
    }

    // Ward Selector Change Handler
    wardViewSelect.addEventListener('change', () => {
        const val = wardViewSelect.value;
        state.currentView = val;
        if (val === 'master') {
            state.activeGrid = state.masterGrid;
        } else {
            const idx = parseInt(val, 10);
            if (state.wards[idx]) {
                state.activeGrid = state.wards[idx].grid;
            }
        }

        applySearch();
    });

    // 7. Status Filter Button Toggles
    function updateStatusFilterButtons() {
        [filterStatusAll, filterStatusActive, filterStatusDeleted].forEach(b => {
            if (b) {
                b.classList.remove('bg-white', 'shadow-2xs');
            }
        });
        if (state.statusFilter === 'all' && filterStatusAll) {
            filterStatusAll.classList.add('bg-white', 'shadow-2xs');
        } else if (state.statusFilter === 'Active' && filterStatusActive) {
            filterStatusActive.classList.add('bg-white', 'shadow-2xs');
        } else if (state.statusFilter === 'Deleted' && filterStatusDeleted) {
            filterStatusDeleted.classList.add('bg-white', 'shadow-2xs');
        }
    }

    if (filterStatusAll) {
        filterStatusAll.addEventListener('click', () => {
            state.statusFilter = 'all';
            updateStatusFilterButtons();
            applySearch();
        });
    }
    if (filterStatusActive) {
        filterStatusActive.addEventListener('click', () => {
            state.statusFilter = 'Active';
            updateStatusFilterButtons();
            applySearch();
        });
    }
    if (filterStatusDeleted) {
        filterStatusDeleted.addEventListener('click', () => {
            state.statusFilter = 'Deleted';
            updateStatusFilterButtons();
            applySearch();
        });
    }

    // 8. Render Active 11-Column Spreadsheet Grid
    function renderActiveTable() {
        editableGrid.innerHTML = '';

        if (!state.filteredGrid || state.filteredGrid.length === 0) {
            editableGrid.innerHTML = '<tbody><tr><td class="p-6 text-center text-slate-400 font-medium">कोई रिकॉर्ड नहीं मिला</td></tr></tbody>';
            tableInfoBadge.textContent = '0 Records';
            return;
        }

        const headers = state.filteredGrid[0] || [];
        const rows = state.filteredGrid.slice(1);

        // Table Header
        const thead = document.createElement('thead');
        thead.className = 'bg-slate-900 text-white sticky top-0 z-10';
        const trHead = document.createElement('tr');

        // Row Index Column
        const thIdx = document.createElement('th');
        thIdx.className = 'p-2 text-center border border-slate-700 w-12 text-[11px] font-extrabold text-slate-400 bg-slate-950';
        thIdx.textContent = '#';
        trHead.appendChild(thIdx);

        headers.forEach((h, colIdx) => {
            const th = document.createElement('th');
            th.className = 'p-2.5 text-left border border-slate-700 text-xs font-bold tracking-tight whitespace-nowrap bg-slate-900';

            if (['वार्ड नं.', 'जि. प.', 'पं. स.', 'वार्ड संख्या', 'भाग संख्या', 'क्रम संख्या', 'आयु', 'मोबाइल नो', 'Status'].includes(h)) {
                th.classList.add('text-center');
            }
            if (h === 'मोबाइल नो') {
                th.innerHTML = `${h} <span class="text-[10px] text-amber-300 font-normal">(खाली)</span>`;
            } else if (h === 'Status') {
                th.className = 'p-2.5 text-center border border-slate-700 text-xs font-extrabold tracking-tight whitespace-nowrap bg-slate-950 text-rose-300';
                th.innerHTML = `Status <span class="text-[9px] text-slate-300 font-normal">(Col 13)</span>`;
            } else {
                th.textContent = h;
            }
            trHead.appendChild(th);
        });

        thead.appendChild(trHead);
        editableGrid.appendChild(thead);

        // Table Body
        const tbody = document.createElement('tbody');
        tbody.className = 'divide-y divide-slate-200';

        const statusColIdx = headers.indexOf('Status');

        rows.forEach((row, rowIdx) => {
            const tr = document.createElement('tr');
            const statusVal = statusColIdx >= 0 ? String(row[statusColIdx] || '').trim() : '';
            const isRowDeleted = statusVal.toLowerCase() === 'deleted';

            if (isRowDeleted) {
                tr.className = 'bg-rose-50/60 hover:bg-rose-100/70 border-l-4 border-rose-500';
            } else {
                tr.className = rowIdx % 2 === 0 ? 'bg-white hover:bg-emerald-50/40' : 'bg-slate-50/70 hover:bg-emerald-50/40';
            }

            // Row Index
            const tdIdx = document.createElement('td');
            tdIdx.className = 'p-1.5 text-center text-slate-400 text-[11px] font-semibold select-none border border-slate-200';
            tdIdx.textContent = rowIdx + 1;
            tr.appendChild(tdIdx);

            row.forEach((cellVal, colIdx) => {
                const td = document.createElement('td');
                td.className = 'p-2 border border-slate-200 text-slate-800 focus:bg-amber-50 focus:outline-emerald-500 whitespace-nowrap';
                td.contentEditable = 'true';

                const hName = headers[colIdx] || '';
                if (['वार्ड नं.', 'जि. प.', 'पं. स.', 'वार्ड संख्या', 'भाग संख्या', 'क्रम संख्या', 'आयु', 'मोबाइल नो', 'वोटर ID', 'हाउस नंबर', 'Status'].includes(hName)) {
                    td.classList.add('text-center');
                }
                if (hName === 'जि. प.' || hName === 'पं. स.') {
                    td.classList.add('font-bold', 'text-sky-800');
                }
                if (hName === 'वोटर ID') {
                    td.classList.add('font-mono', 'font-semibold', 'text-sky-700');
                }
                if (hName === 'नाम') {
                    td.classList.add('font-bold', 'text-slate-900');
                }
                if (hName === 'वार्ड नं.' || hName === 'वार्ड संख्या' || hName === 'भाग संख्या') {
                    td.classList.add('font-bold', 'text-slate-700');
                }

                // Render Column "Status" with dedicated Badge
                if (hName === 'Status') {
                    const sStr = String(cellVal || '').trim();
                    if (sStr.toLowerCase() === 'deleted') {
                        td.innerHTML = '<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-extrabold bg-rose-100 text-rose-800 border border-rose-300 shadow-2xs"><span class="w-1.5 h-1.5 rounded-full bg-rose-600 animate-pulse"></span>Deleted</span>';
                    } else {
                        td.innerHTML = '<span class="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300"><span class="w-1.5 h-1.5 rounded-full bg-emerald-600"></span>Active</span>';
                    }
                } else {
                    td.textContent = cellVal !== null && cellVal !== undefined ? cellVal : '';
                }

                td.addEventListener('blur', () => {
                    const textVal = td.textContent.trim();
                    updateCellValue(rowIdx + 1, colIdx, textVal);
                });

                tr.appendChild(td);
            });

            tbody.appendChild(tr);
        });

        editableGrid.appendChild(tbody);

        // Update badge
        const totalRowsInActive = Math.max(0, state.activeGrid.length - 1);
        const filteredCount = rows.length;
        if (totalRowsInActive === filteredCount) {
            tableInfoBadge.textContent = `${totalRowsInActive} मतदाता`;
        } else {
            tableInfoBadge.textContent = `${filteredCount} / ${totalRowsInActive} प्रदर्शित (${state.statusFilter})`;
        }

        if (window.lucide) lucide.createIcons();
    }

    // 9. Live Search & Status Filtering
    if (searchInput) {
        searchInput.addEventListener('input', applySearch);
    }

    function applySearch() {
        const query = searchInput ? searchInput.value.trim().toLowerCase() : '';
        if (!state.activeGrid || state.activeGrid.length === 0) return;

        const headers = state.activeGrid[0];
        const statusIdx = headers.indexOf('Status');

        let rows = state.activeGrid.slice(1);

        // 1. Status Filter (Active / Deleted)
        if (state.statusFilter !== 'all' && statusIdx !== -1) {
            rows = rows.filter(row => {
                const s = String(row[statusIdx] || '').trim().toLowerCase();
                return s === state.statusFilter.toLowerCase();
            });
        }

        // 2. Text Query Filter
        if (query) {
            rows = rows.filter(row => {
                return row.some(cell => String(cell || '').toLowerCase().includes(query));
            });
        }

        state.filteredGrid = [headers, ...rows];
        renderActiveTable();
    }

    // 10. Cell Value Update in State & Activity History Tracking
    let editHistoryTimer = null;
    function logCellEditActivity() {
        if (editHistoryTimer) clearTimeout(editHistoryTimer);
        editHistoryTimer = setTimeout(() => {
            const fileNames = state.wards.map(w => w.filename).filter(Boolean);
            if (fileNames.length === 0 && state.fileName) fileNames.push(state.fileName);
            fetch('/api/history/log', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    feature_id: 'feature1',
                    feature_name: 'फ़ीचर 1: Excel डेटाबेस',
                    file_names: fileNames.length > 0 ? fileNames : ['मतदाता_सूची.xlsx'],
                    action: 'सेल डेटा संपादन (Inline Edit)',
                    details: 'मतदाता विवरण तालिका में प्रविष्टि संशोधित की गई',
                    status: 'success'
                })
            }).then(() => refreshHistoryPills()).catch(e => console.warn('Edit history log err:', e));
        }, 1200);
    }

    function updateCellValue(filteredRowIdx, colIdx, newVal) {
        if (!state.filteredGrid[filteredRowIdx]) return;
        const prevVal = state.filteredGrid[filteredRowIdx][colIdx];
        if (prevVal === newVal) return;

        state.filteredGrid[filteredRowIdx][colIdx] = newVal;

        const row = state.filteredGrid[filteredRowIdx];
        const origIdx = state.activeGrid.indexOf(row);
        if (origIdx !== -1) {
            state.activeGrid[origIdx][colIdx] = newVal;
        }

        // Record edit in history log
        logCellEditActivity();
    }

    // 11. Excel Download Action (Generates Multi-Sheet Workbook with Status)
    async function triggerExcelDownload() {
        if (!state.allTables || state.allTables.length === 0) {
            alert('डाउनलोड करने के लिए कोई डेटा उपलब्ध नहीं है');
            return;
        }

        const downloadBtns = [heroDownloadBtn, exportExcelBtn];
        downloadBtns.forEach(b => { if (b) b.disabled = true; });

        try {
            // Ensure state.allTables uses the latest summaryGrid and masterGrid
            if (state.allTables && state.allTables.length >= 2) {
                if (state.summaryGrid && state.summaryGrid.length > 0) {
                    state.allTables[0] = state.summaryGrid;
                }
                if (state.masterGrid && state.masterGrid.length > 0) {
                    state.allTables[1] = state.masterGrid;
                }
            }

            const payload = {
                tables: state.allTables,
                table_names: state.allTableNames,
                theme: themeSelect ? themeSelect.value : 'geam_digital',
                layout_mode: 'multi_sheet',
                metadata: state.metadata,
                file_name: state.fileName
            };

            const response = await fetch('/api/export_excel', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || 'Excel export failed');
            }

            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = state.fileName.endsWith('.xlsx') ? state.fileName : `${state.fileName}.xlsx`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
            document.body.removeChild(a);
        } catch (err) {
            alert(`एक्सेल डाउनलोड विफल: ${err.message}`);
        } finally {
            downloadBtns.forEach(b => { if (b) b.disabled = false; });
            refreshHistoryPills();
        }
    }

    if (heroDownloadBtn) heroDownloadBtn.addEventListener('click', triggerExcelDownload);
    if (exportExcelBtn) exportExcelBtn.addEventListener('click', triggerExcelDownload);

    // 12. CSV Download Action
    if (exportCsvBtn) {
        exportCsvBtn.addEventListener('click', async () => {
            if (!state.masterGrid || state.masterGrid.length === 0) return;
            try {
                const payload = {
                    table: state.masterGrid,
                    file_name: state.fileName.replace(/\.xlsx$/, '') + '.csv'
                };
                const response = await fetch('/api/export_csv', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                if (!response.ok) throw new Error('CSV export failed');

                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = payload.file_name;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            } catch (err) {
                alert(`CSV निर्यात विफल: ${err.message}`);
            }
        });
    }

    // 13. Copy TSV Action
    if (copyTsvBtn) {
        copyTsvBtn.addEventListener('click', async () => {
            if (!state.filteredGrid || state.filteredGrid.length === 0) return;
            const tsvContent = state.filteredGrid.map(row => row.join('\t')).join('\n');
            try {
                await navigator.clipboard.writeText(tsvContent);
                const originalHtml = copyTsvBtn.innerHTML;
                copyTsvBtn.innerHTML = '<i data-lucide="check" class="w-3.5 h-3.5 text-emerald-600"></i> कॉपी हो गया!';
                if (window.lucide) lucide.createIcons();
                setTimeout(() => {
                    copyTsvBtn.innerHTML = originalHtml;
                    if (window.lucide) lucide.createIcons();
                }, 2000);
            } catch (err) {
                alert('डेटा कॉपी करने में विफल: ' + err.message);
            }
        });
    }

    // 14. Helper: Set Loading State
    function setLoading(isLoading, statusText = '') {
        if (isLoading) {
            loadingIndicator.classList.remove('hidden');
            if (statusText) loadingStatusText.textContent = statusText;
            convertBtn.disabled = true;
            if (quickSampleVoterBtn) quickSampleVoterBtn.disabled = true;
        } else {
            loadingIndicator.classList.add('hidden');
            convertBtn.disabled = false;
            if (quickSampleVoterBtn) quickSampleVoterBtn.disabled = false;
        }
    }

    // ==============================================================
    // FEATURE 2: PLAIN MATDATA PARCHI (12 PER A4 SHEET) LOGIC
    // ==============================================================

    const parchiState = {
        sessionId: null,
        sheetNames: [],
        activeSheet: '',
        totalVoters: 0,
        activeCount: 0,
        deletedCount: 0,
        currentPage: 1,
        totalPages: 1,
        isZoomed: false,
        layoutType: '12',
        showPageNumber: true,
        jilaParishad: '',
        panchayatSamiti: ''
    };

    let pendingParchiFile = null;
    let pendingFromSession = false;

    // Feature 2 DOM Elements
    const parchiFileInput = document.getElementById('parchiFileInput');
    const parchiDropZone = document.getElementById('parchiDropZone');
    const parchiFileStatusText = document.getElementById('parchiFileStatusText');
    const parchiFileSelectedBox = document.getElementById('parchiFileSelectedBox');
    const parchiSelectedFileName = document.getElementById('parchiSelectedFileName');
    const parchiSelectedFileSize = document.getElementById('parchiSelectedFileSize');
    const parchiStartProcessBtn = document.getElementById('parchiStartProcessBtn');
    const parchiStartProcessBtnText = document.getElementById('parchiStartProcessBtnText');
    const parchiLoadSessionBtn = document.getElementById('parchiLoadSessionBtn');
    const parchiConfigBox = document.getElementById('parchiConfigBox');
    const parchiLayoutTypeSelect = document.getElementById('parchiLayoutTypeSelect');
    const parchiShowPageNumCheck = document.getElementById('parchiShowPageNumCheck');
    const parchiPanchayatInput = document.getElementById('parchiPanchayatInput');
    const parchiVotingTimeInput = document.getElementById('parchiVotingTimeInput');
    const parchiJilaParishadInput = document.getElementById('parchiJilaParishadInput');
    const parchiPanchayatSamitiInput = document.getElementById('parchiPanchayatSamitiInput');
    const parchiSheetSelect = document.getElementById('parchiSheetSelect');
    const parchiActiveOnlyCheck = document.getElementById('parchiActiveOnlyCheck');
    const parchiRefreshBtn = document.getElementById('parchiRefreshBtn');
    const parchiLoading = document.getElementById('parchiLoading');
    const parchiWorkspace = document.getElementById('parchiWorkspace');
    const parchiTotalBadge = document.getElementById('parchiTotalBadge');
    const parchiPagesBadge = document.getElementById('parchiPagesBadge');
    const parchiSlipTypeBadge = document.getElementById('parchiSlipTypeBadge');
    const parchiPrintDirectBtn = document.getElementById('parchiPrintDirectBtn');
    const parchiDownloadPdfBtn = document.getElementById('parchiDownloadPdfBtn');
    const parchiDownloadPdfBtnText = document.getElementById('parchiDownloadPdfBtnText');
    const parchiPrevPageBtn = document.getElementById('parchiPrevPageBtn');
    const parchiNextPageBtn = document.getElementById('parchiNextPageBtn');
    const parchiCurrentPageText = document.getElementById('parchiCurrentPageText');
    const parchiTotalPagesText = document.getElementById('parchiTotalPagesText');
    const parchiVotersRangeText = document.getElementById('parchiVotersRangeText');
    const parchiZoomToggleBtn = document.getElementById('parchiZoomToggleBtn');
    const parchiZoomBtnText = document.getElementById('parchiZoomBtnText');
    const parchiA4Paper = document.getElementById('parchiA4Paper');

    // 1. File Selection Feedback Handler
    function selectExcelFile(file) {
        if (!file || (!file.name.toLowerCase().endsWith('.xlsx') && !file.name.toLowerCase().endsWith('.xls'))) {
            alert('कृपया केवल .xlsx या .xls प्रारूप में एक्सेल फ़ाइल चुनें');
            return;
        }

        pendingParchiFile = file;
        pendingFromSession = false;

        if (parchiSelectedFileName) parchiSelectedFileName.textContent = file.name;
        if (parchiSelectedFileSize) parchiSelectedFileSize.textContent = `${(file.size / 1024).toFixed(0)} KB`;
        if (parchiFileSelectedBox) {
            parchiFileSelectedBox.classList.remove('hidden');
            parchiFileSelectedBox.classList.add('flex');
        }
        if (parchiFileStatusText) {
            parchiFileStatusText.textContent = `चयनित फ़ाइल: ${file.name} - कृपया स्टार्ट बटन दबाएं`;
        }
        if (window.lucide) lucide.createIcons();
    }

    // Drag & Drop for Feature 2 Excel
    if (parchiDropZone && parchiFileInput) {
        parchiDropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            parchiDropZone.classList.add('border-indigo-500', 'bg-indigo-50/40');
        });

        parchiDropZone.addEventListener('dragleave', () => {
            parchiDropZone.classList.remove('border-indigo-500', 'bg-indigo-50/40');
        });

        parchiDropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            parchiDropZone.classList.remove('border-indigo-500', 'bg-indigo-50/40');
            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                parchiFileInput.files = e.dataTransfer.files;
                selectExcelFile(parchiFileInput.files[0]);
            }
        });

        parchiFileInput.addEventListener('change', () => {
            if (parchiFileInput.files && parchiFileInput.files.length > 0) {
                selectExcelFile(parchiFileInput.files[0]);
            }
        });
    }

    // 2. Load from Feature 1 Session Button (Prepares pending session data)
    if (parchiLoadSessionBtn) {
        parchiLoadSessionBtn.addEventListener('click', () => {
            if (!state.sessionId && (!state.allTables || state.allTables.length === 0)) {
                alert('फ़ीचर 1 से कोई डेटा उपलब्ध नहीं है। कृपया पहले पीडीएफ कन्वर्ट करें या एक्सेल फ़ाइल चुनें।');
                return;
            }
            pendingParchiFile = null;
            pendingFromSession = true;

            const sessName = state.fileName ? `${state.fileName.replace(/\.pdf$/i, '')}.xlsx` : 'फ़ीचर_1_मतदाता_सूची.xlsx';
            if (parchiSelectedFileName) parchiSelectedFileName.textContent = sessName;
            if (parchiSelectedFileSize) parchiSelectedFileSize.textContent = 'वर्तमान सत्र डेटा';
            if (parchiFileSelectedBox) {
                parchiFileSelectedBox.classList.remove('hidden');
                parchiFileSelectedBox.classList.add('flex');
            }
            if (parchiFileStatusText) {
                parchiFileStatusText.textContent = 'सत्र डेटा चुना गया - कृपया स्टार्ट बटन दबाएं';
            }
            if (window.lucide) lucide.createIcons();
        });
    }

    // 3. Start Process Button (Starts processing only upon click)
    if (parchiStartProcessBtn) {
        parchiStartProcessBtn.addEventListener('click', async () => {
            if (pendingFromSession) {
                await executeLoadParchiFromSession();
            } else if (pendingParchiFile) {
                await executeParchiExcelUpload(pendingParchiFile);
            } else if (parchiFileInput && parchiFileInput.files && parchiFileInput.files.length > 0) {
                await executeParchiExcelUpload(parchiFileInput.files[0]);
            } else {
                alert('कृपया पहले कोई एक्सेल फ़ाइल चुनें या सत्र लोड करें');
            }
        });
    }

    // Upload & Parse Excel File
    async function executeParchiExcelUpload(file) {
        if (!file) return;

        setParchiLoading(true, 'एक्सेल फ़ाइल का विश्लेषण किया जा रहा है...');
        if (parchiStartProcessBtn) parchiStartProcessBtn.disabled = true;

        const formData = new FormData();
        formData.append('excel_file', file);

        try {
            const response = await fetch('/api/parchi/upload_excel', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || 'एक्सेल अपलोड विफल');
            }

            const data = await response.json();
            applyParchiData(data);
        } catch (err) {
            alert('त्रुटि: ' + err.message);
        } finally {
            setParchiLoading(false);
            if (parchiStartProcessBtn) parchiStartProcessBtn.disabled = false;
        }
    }

    // Load from Session Execution
    async function executeLoadParchiFromSession() {
        setParchiLoading(true, 'वर्तमान सत्र से मतदाता डेटा लोड किया जा रहा है...');
        if (parchiStartProcessBtn) parchiStartProcessBtn.disabled = true;

        try {
            const response = await fetch('/api/parchi/from_session', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ session_id: state.sessionId })
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || 'सत्र डेटा प्राप्त नहीं हो सका');
            }

            const data = await response.json();
            applyParchiData(data);
        } catch (err) {
            alert('सूचना: ' + err.message);
        } finally {
            setParchiLoading(false);
            if (parchiStartProcessBtn) parchiStartProcessBtn.disabled = false;
        }
    }

    // 4. Update Dynamic Labels based on Layout Type
    function updateLayoutLabels() {
        const type = parchiLayoutTypeSelect ? parchiLayoutTypeSelect.value : '12';
        const label = type === '16' ? '16 पर्ची / शीट' : (type === '8' ? '8 पर्ची / शीट' : '12 पर्ची / शीट');
        if (parchiSlipTypeBadge) parchiSlipTypeBadge.textContent = label;
        if (parchiDownloadPdfBtnText) parchiDownloadPdfBtnText.textContent = `${type}-पर्ची PDF डाउनलोड करें`;
    }

    // Layout Type Selector Change Event
    if (parchiLayoutTypeSelect) {
        parchiLayoutTypeSelect.addEventListener('change', () => {
            parchiState.layoutType = parchiLayoutTypeSelect.value;
            updateLayoutLabels();
            parchiState.currentPage = 1;
            fetchParchiPreview();
        });
    }

    // Page Number Toggle Checkbox Change Event
    if (parchiShowPageNumCheck) {
        parchiShowPageNumCheck.addEventListener('change', () => {
            parchiState.showPageNumber = parchiShowPageNumCheck.checked;
            fetchParchiPreview();
        });
    }

    // Active Only Filter Checkbox Change Event
    if (parchiActiveOnlyCheck) {
        parchiActiveOnlyCheck.addEventListener('change', () => {
            parchiState.currentPage = 1;
            fetchParchiPreview();
        });
    }

    // 5. Apply Parsed Data to State and UI
    function applyParchiData(data) {
        parchiState.sessionId = data.session_id;
        parchiState.sheetNames = data.sheet_names || [];
        parchiState.activeSheet = data.active_sheet || '';
        parchiState.totalVoters = data.total_voters || 0;
        parchiState.activeCount = data.active_count || 0;
        parchiState.deletedCount = data.deleted_count || 0;
        parchiState.currentPage = 1;
        parchiState.jilaParishad = data.default_jila_parishad || '';
        parchiState.panchayatSamiti = data.default_panchayat_samiti || '';

        if (parchiPanchayatInput) {
            parchiPanchayatInput.value = data.default_panchayat_name || 'ग्राम पंचायत';
        }
        if (parchiJilaParishadInput) {
            parchiJilaParishadInput.value = parchiState.jilaParishad;
        }
        if (parchiPanchayatSamitiInput) {
            parchiPanchayatSamitiInput.value = parchiState.panchayatSamiti;
        }

        // Populate Sheet Select
        if (parchiSheetSelect) {
            parchiSheetSelect.innerHTML = '';
            parchiState.sheetNames.forEach(s => {
                const opt = document.createElement('option');
                opt.value = s;
                opt.textContent = s;
                if (s === parchiState.activeSheet) opt.selected = true;
                parchiSheetSelect.appendChild(opt);
            });
        }

        updateLayoutLabels();

        // Fetch Live A4 Preview
        fetchParchiPreview();
        refreshHistoryPills();
    }

    // 6. Sheet Selection Change
    if (parchiSheetSelect) {
        parchiSheetSelect.addEventListener('change', async () => {
            const sheetName = parchiSheetSelect.value;
            if (!sheetName || !parchiState.sessionId) return;

            setParchiLoading(true, `शीट '${sheetName}' लोड हो रही है...`);
            try {
                const response = await fetch('/api/parchi/select_sheet', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        session_id: parchiState.sessionId,
                        sheet_name: sheetName
                    })
                });

                if (!response.ok) throw new Error('शीट लोड करने में विफल');
                const data = await response.json();
                applyParchiData(data);
            } catch (err) {
                alert(err.message);
            } finally {
                setParchiLoading(false);
            }
        });
    }

    // 7. Fetch and Render Live A4 Preview HTML
    async function fetchParchiPreview() {
        if (!parchiState.sessionId) return;

        setParchiLoading(true, 'पूर्वावलोकन तैयार हो रहा है...');
        try {
            const layoutType = parchiLayoutTypeSelect ? parchiLayoutTypeSelect.value : '12';
            const showPageNum = parchiShowPageNumCheck ? parchiShowPageNumCheck.checked : true;
            const jp = parchiJilaParishadInput ? parchiJilaParishadInput.value.trim() : (parchiState.jilaParishad || '');
            const ps = parchiPanchayatSamitiInput ? parchiPanchayatSamitiInput.value.trim() : (parchiState.panchayatSamiti || '');

            const payload = {
                session_id: parchiState.sessionId,
                panchayat_name: parchiPanchayatInput ? parchiPanchayatInput.value.trim() || 'ग्राम पंचायत' : 'ग्राम पंचायत',
                voting_time: parchiVotingTimeInput ? parchiVotingTimeInput.value.trim() || 'प्रातः 7 से सायं 6 बजे तक' : 'प्रातः 7 से सायं 6 बजे तक',
                filter_active_only: parchiActiveOnlyCheck ? parchiActiveOnlyCheck.checked : true,
                layout_type: layoutType,
                show_page_number: showPageNum,
                jila_parishad: jp,
                panchayat_samiti: ps,
                page: parchiState.currentPage
            };

            const response = await fetch('/api/parchi/preview_html', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || 'पूर्वावलोकन प्राप्त नहीं हो सका');
            }

            const res = await response.json();
            parchiState.totalPages = res.total_pages || 1;
            parchiState.currentPage = res.current_page || 1;

            // Render HTML into A4 Simulation Paper
            if (parchiA4Paper) {
                const styleMatch = res.html.match(/<style[^>]*>([\s\S]*?)<\/style>/i);
                const bodyMatch = res.html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
                const styles = styleMatch ? `<style>${styleMatch[1]}</style>` : '';
                const bodyContent = bodyMatch ? bodyMatch[1] : res.html;
                parchiA4Paper.innerHTML = styles + bodyContent;
            }

            // Update Badges and Pagination Text
            if (parchiWorkspace) parchiWorkspace.classList.remove('hidden');
            if (parchiTotalBadge) parchiTotalBadge.textContent = `${res.total_voters} मतदाता सम्मिलित`;
            if (parchiPagesBadge) parchiPagesBadge.textContent = `${res.total_pages} A4 पेजेस`;
            if (parchiCurrentPageText) parchiCurrentPageText.textContent = parchiState.currentPage;
            if (parchiTotalPagesText) parchiTotalPagesText.textContent = parchiState.totalPages;

            const cardsPerPage = layoutType === '16' ? 16 : (layoutType === '8' ? 8 : 12);
            const startIdx = (parchiState.currentPage - 1) * cardsPerPage + 1;
            const endIdx = Math.min(parchiState.currentPage * cardsPerPage, res.total_voters);
            if (parchiVotersRangeText) {
                parchiVotersRangeText.textContent = res.total_voters > 0 ? `(पर्ची ${startIdx} से ${endIdx})` : '';
            }

            // Update Pagination buttons state
            if (parchiPrevPageBtn) parchiPrevPageBtn.disabled = (parchiState.currentPage <= 1);
            if (parchiNextPageBtn) parchiNextPageBtn.disabled = (parchiState.currentPage >= parchiState.totalPages);

            if (window.lucide) lucide.createIcons();
        } catch (err) {
            alert('पूर्वावलोकन त्रुटि: ' + err.message);
        } finally {
            setParchiLoading(false);
        }
    }

    if (parchiRefreshBtn) {
        parchiRefreshBtn.addEventListener('click', () => {
            parchiState.currentPage = 1;
            fetchParchiPreview();
        });
    }

    // 8. Pagination Buttons
    if (parchiPrevPageBtn) {
        parchiPrevPageBtn.addEventListener('click', () => {
            if (parchiState.currentPage > 1) {
                parchiState.currentPage--;
                fetchParchiPreview();
            }
        });
    }

    if (parchiNextPageBtn) {
        parchiNextPageBtn.addEventListener('click', () => {
            if (parchiState.currentPage < parchiState.totalPages) {
                parchiState.currentPage++;
                fetchParchiPreview();
            }
        });
    }

    // 9. Zoom Toggle
    if (parchiZoomToggleBtn && parchiA4Paper) {
        parchiZoomToggleBtn.addEventListener('click', () => {
            parchiState.isZoomed = !parchiState.isZoomed;
            if (parchiState.isZoomed) {
                parchiA4Paper.classList.remove('a4-scaled');
                parchiZoomBtnText.textContent = 'फ़िट टू स्क्रीन';
            } else {
                parchiA4Paper.classList.add('a4-scaled');
                parchiZoomBtnText.textContent = '100% ज़ूम';
            }
        });
    }

    // 10. Direct A4 Browser Print Action
    if (parchiPrintDirectBtn) {
        parchiPrintDirectBtn.addEventListener('click', () => {
            if (!parchiState.sessionId) {
                alert('कृपया पहले कोई एक्सेल फ़ाइल अपलोड करें या सत्र लोड करें');
                return;
            }

            const pName = encodeURIComponent((parchiPanchayatInput ? parchiPanchayatInput.value.trim() : '') || 'ग्राम पंचायत');
            const vTime = encodeURIComponent((parchiVotingTimeInput ? parchiVotingTimeInput.value.trim() : '') || 'प्रातः 7 से सायं 6 बजे तक');
            const actOnly = parchiActiveOnlyCheck ? parchiActiveOnlyCheck.checked : true;
            const layoutType = parchiLayoutTypeSelect ? parchiLayoutTypeSelect.value : '12';
            const showPageNum = parchiShowPageNumCheck ? parchiShowPageNumCheck.checked : true;
            const jp = encodeURIComponent((parchiJilaParishadInput ? parchiJilaParishadInput.value.trim() : parchiState.jilaParishad) || '');
            const ps = encodeURIComponent((parchiPanchayatSamitiInput ? parchiPanchayatSamitiInput.value.trim() : parchiState.panchayatSamiti) || '');

            const printUrl = `/parchi/print?session_id=${parchiState.sessionId}&panchayat_name=${pName}&voting_time=${vTime}&filter_active_only=${actOnly}&layout_type=${layoutType}&show_page_number=${showPageNum}&jila_parishad=${jp}&panchayat_samiti=${ps}`;
            
            const printWin = window.open(printUrl, '_blank');
            if (!printWin) {
                alert('कृपया अपने ब्राउज़र में पॉप-अप (Pop-ups) की अनुमति दें ताकि प्रिंट डायलॉग खुल सके।');
            }
        });
    }

    // 11. Vector PDF Download Action
    if (parchiDownloadPdfBtn) {
        parchiDownloadPdfBtn.addEventListener('click', async () => {
            if (!parchiState.sessionId) {
                alert('कृपया पहले कोई एक्सेल फ़ाइल अपलोड करें या सत्र लोड करें');
                return;
            }

            parchiDownloadPdfBtn.disabled = true;
            const originalHtml = parchiDownloadPdfBtn.innerHTML;
            parchiDownloadPdfBtn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i> <span>PDF जनरेट हो रही है...</span>';
            if (window.lucide) lucide.createIcons();

            try {
                const layoutType = parchiLayoutTypeSelect ? parchiLayoutTypeSelect.value : '12';
                const showPageNum = parchiShowPageNumCheck ? parchiShowPageNumCheck.checked : true;
                const jp = parchiJilaParishadInput ? parchiJilaParishadInput.value.trim() : (parchiState.jilaParishad || '');
                const ps = parchiPanchayatSamitiInput ? parchiPanchayatSamitiInput.value.trim() : (parchiState.panchayatSamiti || '');

                const payload = {
                    session_id: parchiState.sessionId,
                    panchayat_name: parchiPanchayatInput ? parchiPanchayatInput.value.trim() || 'ग्राम पंचायत' : 'ग्राम पंचायत',
                    voting_time: parchiVotingTimeInput ? parchiVotingTimeInput.value.trim() || 'प्रातः 7 से सायं 6 बजे तक' : 'प्रातः 7 से सायं 6 बजे तक',
                    filter_active_only: parchiActiveOnlyCheck ? parchiActiveOnlyCheck.checked : true,
                    layout_type: layoutType,
                    show_page_number: showPageNum,
                    jila_parishad: jp,
                    panchayat_samiti: ps
                };

                const response = await fetch('/api/parchi/export_pdf', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.error || 'PDF डाउनलोड विफल');
                }

                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                const safeName = ((parchiPanchayatInput ? parchiPanchayatInput.value.trim() : '') || 'Matdata_Parchi').replace(/[^\w\s-]/g, '').trim() || 'Parchi';
                a.download = `Geam_Digital_${safeName}_A4_${layoutType}_Parchi.pdf`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            } catch (err) {
                alert('PDF निर्यात त्रुटि: ' + err.message);
            } finally {
                parchiDownloadPdfBtn.disabled = false;
                parchiDownloadPdfBtn.innerHTML = originalHtml;
                if (window.lucide) lucide.createIcons();
            }
        });
    }

    // Helper: Set Parchi Loading
    function setParchiLoading(isLoading, text = '') {
        if (!parchiLoading) return;
        if (isLoading) {
            parchiLoading.classList.remove('hidden');
            if (text && parchiLoading.querySelector('p')) {
                parchiLoading.querySelector('p').textContent = text;
            }
            if (parchiRefreshBtn) parchiRefreshBtn.disabled = true;
        } else {
            parchiLoading.classList.add('hidden');
            if (parchiRefreshBtn) parchiRefreshBtn.disabled = false;
        }
    }

    // ==============================================================
    // FEATURE 3: COLOR VOTER SLIP (CANDIDATE POSTER) CLIENT LOGIC
    // ==============================================================

    const colorState = {
        sessionId: null,
        sheetNames: [],
        activeSheet: '',
        totalVoters: 0,
        activeCount: 0,
        deletedCount: 0,
        currentPage: 1,
        totalPages: 1,
        isZoomed: false,
        selectedFile: null,
        uploadedPosterDataUrl: null,
        uploadedPhotoDataUrl: null,
        uploadedSymbolDataUrl: null,
        jilaParishad: '',
        panchayatSamiti: ''
    };

    // DOM Elements - Feature 3
    const colorDropZone = document.getElementById('colorDropZone');
    const colorFileInput = document.getElementById('colorFileInput');
    const colorLoadSessionBtn = document.getElementById('colorLoadSessionBtn');
    const colorFileSelectedBox = document.getElementById('colorFileSelectedBox');
    const colorSelectedFileName = document.getElementById('colorSelectedFileName');
    const colorSelectedFileSize = document.getElementById('colorSelectedFileSize');
    const colorStartProcessBtn = document.getElementById('colorStartProcessBtn');
    const colorStartProcessBtnText = document.getElementById('colorStartProcessBtnText');
    const colorConfigBox = document.getElementById('colorConfigBox');

    // Poster Controls & Sizing
    const colorPosterFileInput = document.getElementById('colorPosterFileInput');
    const colorPosterEmptyState = document.getElementById('colorPosterEmptyState');
    const colorPosterUploadedState = document.getElementById('colorPosterUploadedState');
    const colorPosterThumbImg = document.getElementById('colorPosterThumbImg');
    const colorPosterFileName = document.getElementById('colorPosterFileName');
    const colorPosterActiveBadge = document.getElementById('colorPosterActiveBadge');
    const colorChangePosterBtn = document.getElementById('colorChangePosterBtn');
    const colorDeletePosterBtn = document.getElementById('colorDeletePosterBtn');

    // Customization Accordion
    const colorCustomizationToggleBtn = document.getElementById('colorCustomizationToggleBtn');
    const colorCustomizationStatusBadge = document.getElementById('colorCustomizationStatusBadge');
    const colorCustomizationChevron = document.getElementById('colorCustomizationChevron');
    const colorCustomizationBody = document.getElementById('colorCustomizationBody');

    // Live Mini-Poster Preview Elements
    const colorMiniPosterPreview = document.getElementById('colorMiniPosterPreview');
    const miniPosterRibbon = document.getElementById('miniPosterRibbon');
    const miniPosterPrefix = document.getElementById('miniPosterPrefix');
    const miniPosterWardBadge = document.getElementById('miniPosterWardBadge');
    const miniPosterSuffix = document.getElementById('miniPosterSuffix');
    const miniPosterSymbolImg = document.getElementById('miniPosterSymbolImg');
    const miniPosterCandName = document.getElementById('miniPosterCandName');
    const miniPosterBtnBox = document.getElementById('miniPosterBtnBox');
    const miniPosterBtnNum = document.getElementById('miniPosterBtnNum');
    const miniPosterSlogan = document.getElementById('miniPosterSlogan');
    const miniPosterPhotoImg = document.getElementById('miniPosterPhotoImg');

    // Candidate Text Inputs
    const colorHeaderPrefixInput = document.getElementById('colorHeaderPrefixInput');
    const colorWardBadgeInput = document.getElementById('colorWardBadgeInput');
    const colorHeaderSuffixInput = document.getElementById('colorHeaderSuffixInput');
    const colorCandidateNameInput = document.getElementById('colorCandidateNameInput');
    const colorButtonNoInput = document.getElementById('colorButtonNoInput');
    const colorSloganInput = document.getElementById('colorSloganInput');
    const colorJilaParishadInput = document.getElementById('colorJilaParishadInput');
    const colorPanchayatSamitiInput = document.getElementById('colorPanchayatSamitiInput');

    // Candidate Photo Controls
    const colorPhotoFileInput = document.getElementById('colorPhotoFileInput');
    const colorPhotoUploadBtn = document.getElementById('colorPhotoUploadBtn');
    const colorPhotoEmptyState = document.getElementById('colorPhotoEmptyState');
    const colorPhotoUploadedState = document.getElementById('colorPhotoUploadedState');
    const colorPhotoThumbImg = document.getElementById('colorPhotoThumbImg');
    const colorChangePhotoBtn = document.getElementById('colorChangePhotoBtn');
    const colorDeletePhotoBtn = document.getElementById('colorDeletePhotoBtn');

    // Symbol & Party Logo Controls
    const colorSymbolSelect = document.getElementById('colorSymbolSelect');
    const colorSymbolFileInput = document.getElementById('colorSymbolFileInput');
    const colorSymbolUploadBtn = document.getElementById('colorSymbolUploadBtn');
    const colorSymbolEmptyState = document.getElementById('colorSymbolEmptyState');
    const colorSymbolUploadedState = document.getElementById('colorSymbolUploadedState');
    const colorSymbolThumbImg = document.getElementById('colorSymbolThumbImg');
    const colorChangeSymbolBtn = document.getElementById('colorChangeSymbolBtn');
    const colorDeleteSymbolBtn = document.getElementById('colorDeleteSymbolBtn');

    // Color Pickers
    const colorPosterBgColorInput = document.getElementById('colorPosterBgColorInput');
    const colorNameColorInput = document.getElementById('colorNameColorInput');
    const colorBtnColorInput = document.getElementById('colorBtnColorInput');
    const colorTimeColorInput = document.getElementById('colorTimeColorInput');

    // Layout Controls
    const colorLayoutTypeSelect = document.getElementById('colorLayoutTypeSelect');
    const colorSheetSelect = document.getElementById('colorSheetSelect');
    const colorVotingTimeInput = document.getElementById('colorVotingTimeInput');
    const colorActiveOnlyCheck = document.getElementById('colorActiveOnlyCheck');
    const colorShowPageNumCheck = document.getElementById('colorShowPageNumCheck');
    const colorRefreshBtn = document.getElementById('colorRefreshBtn');
    const colorLoading = document.getElementById('colorLoading');

    // Output & Workspace
    const colorWorkspace = document.getElementById('colorWorkspace');
    const colorTotalBadge = document.getElementById('colorTotalBadge');
    const colorPagesBadge = document.getElementById('colorPagesBadge');
    const colorSlipTypeBadge = document.getElementById('colorSlipTypeBadge');
    const colorStatusHeading = document.getElementById('colorStatusHeading');
    const colorPrintDirectBtn = document.getElementById('colorPrintDirectBtn');
    const colorDownloadPdfBtn = document.getElementById('colorDownloadPdfBtn');
    const colorDownloadPdfBtnText = document.getElementById('colorDownloadPdfBtnText');
    const colorPrevPageBtn = document.getElementById('colorPrevPageBtn');
    const colorNextPageBtn = document.getElementById('colorNextPageBtn');
    const colorCurrentPageText = document.getElementById('colorCurrentPageText');
    const colorTotalPagesText = document.getElementById('colorTotalPagesText');
    const colorVotersRangeText = document.getElementById('colorVotersRangeText');
    const colorZoomToggleBtn = document.getElementById('colorZoomToggleBtn');
    const colorZoomBtnText = document.getElementById('colorZoomBtnText');
    const colorParchiA4Paper = document.getElementById('colorParchiA4Paper');

    // Default SVGs Cache for Instant Live Preview
    const colorDefaultSvgs = {
        hand: '',
        lotus: '',
        cycle: '',
        candidate_photo: ''
    };

    async function loadColorParchiDefaults() {
        try {
            const res = await fetch('/api/color_parchi/defaults');
            if (res.ok) {
                const data = await res.json();
                colorDefaultSvgs.hand = data.hand;
                colorDefaultSvgs.lotus = data.lotus;
                colorDefaultSvgs.cycle = data.cycle;
                colorDefaultSvgs.candidate_photo = data.candidate_photo;
                updateMiniPosterLivePreview();
            }
        } catch (e) {
            console.warn('Could not load SVG defaults', e);
        }
    }
    loadColorParchiDefaults();

    // Live Mini-Poster Real-Time Updater (Zero Latency)
    function updateMiniPosterLivePreview() {
        if (miniPosterPrefix) miniPosterPrefix.textContent = colorHeaderPrefixInput?.value || 'पार्षद पद हेतु वार्ड नं';
        if (miniPosterWardBadge) miniPosterWardBadge.textContent = colorWardBadgeInput?.value || '03';
        if (miniPosterSuffix) miniPosterSuffix.textContent = colorHeaderSuffixInput?.value || 'से लोकप्रिय प्रत्याशी';
        
        if (miniPosterCandName) {
            miniPosterCandName.textContent = colorCandidateNameInput?.value || 'संजय कुमार मेवाड़ा';
            miniPosterCandName.style.color = colorNameColorInput?.value || '#b91c1c';
        }
        
        if (miniPosterBtnBox) {
            miniPosterBtnBox.style.background = colorBtnColorInput?.value || '#0284c7';
        }
        
        if (miniPosterBtnNum) {
            miniPosterBtnNum.textContent = colorButtonNoInput?.value || '2';
            miniPosterBtnNum.style.color = colorBtnColorInput?.value || '#0284c7';
        }
        
        if (miniPosterSlogan) {
            miniPosterSlogan.textContent = colorSloganInput?.value || 'को हाथ के निशान पर बटन दबा कर विजयी बनावें।';
        }
        
        if (colorMiniPosterPreview) {
            colorMiniPosterPreview.style.backgroundColor = colorPosterBgColorInput?.value || '#fff033';
        }

        // Live Photo render
        if (miniPosterPhotoImg) {
            if (colorState.uploadedPhotoDataUrl) {
                miniPosterPhotoImg.src = colorState.uploadedPhotoDataUrl;
            } else if (colorDefaultSvgs.candidate_photo) {
                miniPosterPhotoImg.src = colorDefaultSvgs.candidate_photo;
            }
        }

        // Live Symbol render
        if (miniPosterSymbolImg) {
            if (colorState.uploadedSymbolDataUrl) {
                miniPosterSymbolImg.src = colorState.uploadedSymbolDataUrl;
            } else {
                const preset = colorSymbolSelect ? colorSymbolSelect.value : 'hand';
                if (colorDefaultSvgs[preset]) {
                    miniPosterSymbolImg.src = colorDefaultSvgs[preset];
                }
            }
        }
    }

    // Attach input listeners to all candidate fields for live instant feedback
    [
        colorHeaderPrefixInput,
        colorWardBadgeInput,
        colorHeaderSuffixInput,
        colorCandidateNameInput,
        colorButtonNoInput,
        colorSloganInput,
        colorPosterBgColorInput,
        colorNameColorInput,
        colorBtnColorInput
    ].forEach(el => {
        if (el) {
            el.addEventListener('input', updateMiniPosterLivePreview);
            el.addEventListener('change', () => {
                updateMiniPosterLivePreview();
                fetchColorParchiPreview();
            });
        }
    });

    // Preset Swatches Click Handlers
    document.querySelectorAll('.color-preset-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            const targetId = btn.getAttribute('data-target');
            const colorVal = btn.getAttribute('data-color');
            const targetInput = document.getElementById(targetId);
            if (targetInput && colorVal) {
                targetInput.value = colorVal;
                updateMiniPosterLivePreview();
                fetchColorParchiPreview();
            }
        });
    });

    // Customization Accordion Toggle
    if (colorCustomizationToggleBtn && colorCustomizationBody) {
        colorCustomizationToggleBtn.addEventListener('click', () => {
            colorCustomizationBody.classList.toggle('hidden');
            if (colorCustomizationChevron) {
                colorCustomizationChevron.classList.toggle('rotate-180');
            }
        });
    }

    // 1. Ready-Made Full Poster Upload, Change & Delete Logic
    if (colorPosterEmptyState && colorPosterFileInput) {
        colorPosterEmptyState.addEventListener('click', () => {
            colorPosterFileInput.click();
        });
    }

    if (colorChangePosterBtn && colorPosterFileInput) {
        colorChangePosterBtn.addEventListener('click', () => {
            colorPosterFileInput.click();
        });
    }

    async function handleReadyMadePosterFile(file) {
        if (!file) return;
        try {
            setColorLoading(true, 'तैयार पोस्टर लोड हो रहा है...');
            const dataUrl = await uploadImageAsset(file);
            colorState.uploadedPosterDataUrl = dataUrl;

            if (colorPosterThumbImg) colorPosterThumbImg.src = dataUrl;
            if (colorPosterFileName) colorPosterFileName.textContent = `${file.name} (${(file.size / 1024).toFixed(0)} KB)`;

            if (colorPosterEmptyState) colorPosterEmptyState.classList.add('hidden');
            if (colorPosterUploadedState) {
                colorPosterUploadedState.classList.remove('hidden');
                colorPosterUploadedState.classList.add('flex');
            }
            if (colorPosterActiveBadge) {
                colorPosterActiveBadge.classList.remove('hidden');
                colorPosterActiveBadge.classList.add('flex');
            }

            // Auto-collapse customization section as ready-made poster is now active
            if (colorCustomizationBody) colorCustomizationBody.classList.add('hidden');
            if (colorCustomizationChevron) colorCustomizationChevron.classList.remove('rotate-180');
            if (colorCustomizationStatusBadge) {
                colorCustomizationStatusBadge.textContent = 'रेडीमेड पोस्टर सक्रिय (कस्टमाइज़ेशन बंद)';
                colorCustomizationStatusBadge.className = 'text-[10px] bg-emerald-100 text-emerald-800 font-extrabold px-2 py-0.5 rounded-full lowercase tracking-normal';
            }

            fetchColorParchiPreview();
        } catch (err) {
            alert('पोस्टर अपलोड त्रुटि: ' + err.message);
        } finally {
            setColorLoading(false);
        }
    }

    if (colorPosterEmptyState) {
        colorPosterEmptyState.addEventListener('dragover', (e) => {
            e.preventDefault();
            colorPosterEmptyState.classList.add('border-amber-500', 'bg-amber-100/60');
        });
        colorPosterEmptyState.addEventListener('dragleave', () => {
            colorPosterEmptyState.classList.remove('border-amber-500', 'bg-amber-100/60');
        });
        colorPosterEmptyState.addEventListener('drop', (e) => {
            e.preventDefault();
            colorPosterEmptyState.classList.remove('border-amber-500', 'bg-amber-100/60');
            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                handleReadyMadePosterFile(e.dataTransfer.files[0]);
            }
        });
    }

    if (colorPosterFileInput) {
        colorPosterFileInput.addEventListener('change', async () => {
            if (colorPosterFileInput.files && colorPosterFileInput.files.length > 0) {
                await handleReadyMadePosterFile(colorPosterFileInput.files[0]);
            }
        });
    }

    if (colorDeletePosterBtn) {
        colorDeletePosterBtn.addEventListener('click', () => {
            colorState.uploadedPosterDataUrl = null;
            if (colorPosterFileInput) colorPosterFileInput.value = '';

            if (colorPosterUploadedState) {
                colorPosterUploadedState.classList.add('hidden');
                colorPosterUploadedState.classList.remove('flex');
            }
            if (colorPosterEmptyState) colorPosterEmptyState.classList.remove('hidden');
            if (colorPosterActiveBadge) {
                colorPosterActiveBadge.classList.add('hidden');
                colorPosterActiveBadge.classList.remove('flex');
            }

            // Auto-expand customization section
            if (colorCustomizationBody) colorCustomizationBody.classList.remove('hidden');
            if (colorCustomizationChevron) colorCustomizationChevron.classList.add('rotate-180');
            if (colorCustomizationStatusBadge) {
                colorCustomizationStatusBadge.textContent = 'मैन्युअल कस्टमाइज़ेशन सक्रिय';
                colorCustomizationStatusBadge.className = 'text-[10px] bg-slate-200 text-slate-700 font-extrabold px-2 py-0.5 rounded-full lowercase tracking-normal';
            }

            fetchColorParchiPreview();
        });
    }

    // 2. Candidate Photo Upload, Change & Delete Logic
    if (colorPhotoUploadBtn && colorPhotoFileInput) {
        colorPhotoUploadBtn.addEventListener('click', () => {
            colorPhotoFileInput.click();
        });
    }

    if (colorChangePhotoBtn && colorPhotoFileInput) {
        colorChangePhotoBtn.addEventListener('click', () => {
            colorPhotoFileInput.click();
        });
    }

    if (colorPhotoFileInput) {
        colorPhotoFileInput.addEventListener('change', async () => {
            if (colorPhotoFileInput.files && colorPhotoFileInput.files.length > 0) {
                const file = colorPhotoFileInput.files[0];
                try {
                    setColorLoading(true, 'उम्मीदवार फोटो अपलोड हो रही है...');
                    const dataUrl = await uploadImageAsset(file);
                    colorState.uploadedPhotoDataUrl = dataUrl;

                    if (colorPhotoThumbImg) colorPhotoThumbImg.src = dataUrl;
                    if (colorPhotoEmptyState) colorPhotoEmptyState.classList.add('hidden');
                    if (colorPhotoUploadedState) {
                        colorPhotoUploadedState.classList.remove('hidden');
                        colorPhotoUploadedState.classList.add('flex');
                    }

                    updateMiniPosterLivePreview();
                    fetchColorParchiPreview();
                } catch (err) {
                    alert('फोटो अपलोड त्रुटि: ' + err.message);
                } finally {
                    setColorLoading(false);
                }
            }
        });
    }

    if (colorDeletePhotoBtn) {
        colorDeletePhotoBtn.addEventListener('click', () => {
            colorState.uploadedPhotoDataUrl = null;
            if (colorPhotoFileInput) colorPhotoFileInput.value = '';

            if (colorPhotoUploadedState) {
                colorPhotoUploadedState.classList.add('hidden');
                colorPhotoUploadedState.classList.remove('flex');
            }
            if (colorPhotoEmptyState) colorPhotoEmptyState.classList.remove('hidden');

            updateMiniPosterLivePreview();
            fetchColorParchiPreview();
        });
    }

    // 3. Symbol / Party Logo Upload, Change & Delete Logic
    if (colorSymbolUploadBtn && colorSymbolFileInput) {
        colorSymbolUploadBtn.addEventListener('click', () => {
            colorSymbolFileInput.click();
        });
    }

    if (colorChangeSymbolBtn && colorSymbolFileInput) {
        colorChangeSymbolBtn.addEventListener('click', () => {
            colorSymbolFileInput.click();
        });
    }

    if (colorSymbolFileInput) {
        colorSymbolFileInput.addEventListener('change', async () => {
            if (colorSymbolFileInput.files && colorSymbolFileInput.files.length > 0) {
                const file = colorSymbolFileInput.files[0];
                try {
                    setColorLoading(true, 'पार्टी सिंबल अपलोड हो रहा है...');
                    const dataUrl = await uploadImageAsset(file);
                    colorState.uploadedSymbolDataUrl = dataUrl;

                    if (colorSymbolThumbImg) colorSymbolThumbImg.src = dataUrl;
                    if (colorSymbolEmptyState) colorSymbolEmptyState.classList.add('hidden');
                    if (colorSymbolUploadedState) {
                        colorSymbolUploadedState.classList.remove('hidden');
                        colorSymbolUploadedState.classList.add('flex');
                    }
                    if (colorSymbolSelect) colorSymbolSelect.value = 'custom';

                    updateMiniPosterLivePreview();
                    fetchColorParchiPreview();
                } catch (err) {
                    alert('सिंबल अपलोड त्रुटि: ' + err.message);
                } finally {
                    setColorLoading(false);
                }
            }
        });
    }

    if (colorDeleteSymbolBtn) {
        colorDeleteSymbolBtn.addEventListener('click', () => {
            colorState.uploadedSymbolDataUrl = null;
            if (colorSymbolFileInput) colorSymbolFileInput.value = '';

            if (colorSymbolUploadedState) {
                colorSymbolUploadedState.classList.add('hidden');
                colorSymbolUploadedState.classList.remove('flex');
            }
            if (colorSymbolEmptyState) colorSymbolEmptyState.classList.remove('hidden');
            if (colorSymbolSelect) colorSymbolSelect.value = 'hand';

            updateMiniPosterLivePreview();
            fetchColorParchiPreview();
        });
    }

    if (colorSymbolSelect) {
        colorSymbolSelect.addEventListener('change', () => {
            if (colorSymbolSelect.value === 'custom' && !colorState.uploadedSymbolDataUrl) {
                colorSymbolFileInput?.click();
            } else {
                updateMiniPosterLivePreview();
                fetchColorParchiPreview();
            }
        });
    }

    // Helper: Gather Candidate & Poster Info with Custom Colors
    function gatherCandidateInfo() {
        const jp = colorJilaParishadInput ? colorJilaParishadInput.value.trim() : (colorState.jilaParishad || '');
        const ps = colorPanchayatSamitiInput ? colorPanchayatSamitiInput.value.trim() : (colorState.panchayatSamiti || '');
        return {
            header_prefix: colorHeaderPrefixInput?.value.trim() || 'पार्षद पद हेतु वार्ड नं',
            ward_badge: colorWardBadgeInput?.value.trim() || '03',
            header_suffix: colorHeaderSuffixInput?.value.trim() || 'से लोकप्रिय प्रत्याशी',
            candidate_name: colorCandidateNameInput?.value.trim() || 'संजय कुमार मेवाड़ा',
            button_no: colorButtonNoInput?.value.trim() || '2',
            slogan: colorSloganInput?.value.trim() || 'को हाथ के निशान पर बटन दबा कर विजयी बनावें।',
            symbol_preset: colorSymbolSelect?.value || 'hand',
            custom_symbol_data: colorState.uploadedSymbolDataUrl,
            candidate_photo_data: colorState.uploadedPhotoDataUrl,
            poster_img_data: colorState.uploadedPosterDataUrl,
            poster_bg_color: colorPosterBgColorInput?.value || '#fff033',
            candidate_name_color: colorNameColorInput?.value || '#b91c1c',
            btn_box_color: colorBtnColorInput?.value || '#0284c7',
            voting_time: colorVotingTimeInput?.value.trim() || 'प्रातः 7 से सायं 6 बजे तक',
            time_color: colorTimeColorInput?.value || '#e11d48',
            jila_parishad: jp,
            panchayat_samiti: ps
        };
    }

    // Drag & Drop for Feature 3
    if (colorDropZone) {
        colorDropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            colorDropZone.classList.add('border-rose-500', 'bg-rose-50/40');
        });

        colorDropZone.addEventListener('dragleave', () => {
            colorDropZone.classList.remove('border-rose-500', 'bg-rose-50/40');
        });

        colorDropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            colorDropZone.classList.remove('border-rose-500', 'bg-rose-50/40');
            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                handleColorFileChosen(e.dataTransfer.files[0]);
            }
        });

        colorFileInput?.addEventListener('change', () => {
            if (colorFileInput.files && colorFileInput.files.length > 0) {
                handleColorFileChosen(colorFileInput.files[0]);
            }
        });
    }

    function handleColorFileChosen(file) {
        const name = file.name.toLowerCase();
        if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
            alert('कृपया केवल वैध .xlsx या .xls एक्सेल फ़ाइल चुनें');
            return;
        }

        colorState.selectedFile = file;

        if (colorSelectedFileName) colorSelectedFileName.textContent = file.name;
        if (colorSelectedFileSize) colorSelectedFileSize.textContent = `${(file.size / 1024).toFixed(1)} KB`;

        if (colorFileSelectedBox) {
            colorFileSelectedBox.classList.remove('hidden');
            colorFileSelectedBox.classList.add('flex');
        }

        if (colorStartProcessBtnText) {
            colorStartProcessBtnText.textContent = `🚀 "${file.name}" से कलर पर्ची प्रोसेस करें`;
        }
    }

    // Start Process Button Handler
    if (colorStartProcessBtn) {
        colorStartProcessBtn.addEventListener('click', () => {
            if (colorState.selectedFile) {
                executeUploadColorExcel(colorState.selectedFile);
            } else if (colorState.sessionId) {
                fetchColorParchiPreview();
            } else {
                alert('कृपया पहले कोई एक्सेल फ़ाइल चुनें या वर्तमान सत्र से डेटा लोड करें');
            }
        });
    }

    // Upload Excel Execution
    async function executeUploadColorExcel(file) {
        setColorLoading(true, `एक्सेल फ़ाइल '${file.name}' पार्स की जा रही है...`);
        if (colorStartProcessBtn) colorStartProcessBtn.disabled = true;

        try {
            const formData = new FormData();
            formData.append('excel_file', file);

            const response = await fetch('/api/color_parchi/upload_excel', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || 'एक्सेल फ़ाइल लोड करने में विफल');
            }

            const data = await response.json();
            applyColorParchiData(data);
        } catch (err) {
            alert('त्रुटि: ' + err.message);
        } finally {
            setColorLoading(false);
            if (colorStartProcessBtn) colorStartProcessBtn.disabled = false;
        }
    }

    // Load from Session Button Handler
    if (colorLoadSessionBtn) {
        colorLoadSessionBtn.addEventListener('click', () => {
            executeLoadColorParchiFromSession();
        });
    }

    async function executeLoadColorParchiFromSession() {
        setColorLoading(true, 'वर्तमान सत्र से मतदाता डेटा लोड किया जा रहा है...');
        if (colorStartProcessBtn) colorStartProcessBtn.disabled = true;

        try {
            const response = await fetch('/api/color_parchi/from_session', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ session_id: state.sessionId })
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || 'सत्र डेटा प्राप्त नहीं हो सका');
            }

            const data = await response.json();
            applyColorParchiData(data);
        } catch (err) {
            alert('सूचना: ' + err.message);
        } finally {
            setColorLoading(false);
            if (colorStartProcessBtn) colorStartProcessBtn.disabled = false;
        }
    }

    function loadColorParchiFromFeature1Session() {
        switchTab('feature3');
        executeLoadColorParchiFromSession();
    }

    // Apply Parsed Data to Color State & UI
    function applyColorParchiData(data) {
        colorState.sessionId = data.session_id;
        colorState.sheetNames = data.sheet_names || [];
        colorState.activeSheet = data.active_sheet || '';
        colorState.totalVoters = data.total_voters || 0;
        colorState.activeCount = data.active_count || 0;
        colorState.deletedCount = data.deleted_count || 0;
        colorState.currentPage = 1;
        colorState.jilaParishad = data.default_jila_parishad || '';
        colorState.panchayatSamiti = data.default_panchayat_samiti || '';

        if (data.default_panchayat_name && colorWardBadgeInput) {
            const match = data.default_panchayat_name.match(/\d+/);
            if (match) colorWardBadgeInput.value = match[0].padStart(2, '0');
        }
        if (colorJilaParishadInput) {
            colorJilaParishadInput.value = colorState.jilaParishad;
        }
        if (colorPanchayatSamitiInput) {
            colorPanchayatSamitiInput.value = colorState.panchayatSamiti;
        }

        // Populate Sheet Select
        if (colorSheetSelect) {
            colorSheetSelect.innerHTML = '';
            colorState.sheetNames.forEach(s => {
                const opt = document.createElement('option');
                opt.value = s;
                opt.textContent = s;
                if (s === colorState.activeSheet) opt.selected = true;
                colorSheetSelect.appendChild(opt);
            });
        }

        updateColorLayoutLabels();
        fetchColorParchiPreview();
        refreshHistoryPills();
    }

    // Sheet Selection Change
    if (colorSheetSelect) {
        colorSheetSelect.addEventListener('change', async () => {
            const sheetName = colorSheetSelect.value;
            if (!sheetName || !colorState.sessionId) return;

            setColorLoading(true, `शीट '${sheetName}' लोड हो रही है...`);
            try {
                const response = await fetch('/api/color_parchi/select_sheet', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        session_id: colorState.sessionId,
                        sheet_name: sheetName
                    })
                });

                if (!response.ok) throw new Error('शीट लोड करने में विफल');
                const data = await response.json();
                applyColorParchiData(data);
            } catch (err) {
                alert(err.message);
            } finally {
                setColorLoading(false);
            }
        });
    }

    // Helper: Upload Image Asset (Photo, Symbol, or Full Poster)
    async function uploadImageAsset(file) {
        if (!file) throw new Error('कृपया एक वैध इमेज फ़ाइल चुनें');

        // 1. Fast client-side read (instant 0ms conversion, no network limits, works offline)
        try {
            const dataUrl = await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = (e) => {
                    if (e.target && e.target.result) {
                        resolve(e.target.result);
                    } else {
                        reject(new Error('इमेज डेटा खाली है'));
                    }
                };
                reader.onerror = () => reject(new Error('फ़ाइल पढ़ने में विफल'));
                reader.readAsDataURL(file);
            });
            if (dataUrl) return dataUrl;
        } catch (clientErr) {
            console.warn('Client FileReader fallback to server upload...', clientErr);
        }

        // 2. Server-side fallback via /api/color_parchi/upload_asset
        const formData = new FormData();
        formData.append('image_file', file);

        const response = await fetch('/api/color_parchi/upload_asset', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            const err = await response.json().catch(() => ({}));
            throw new Error(err.error || 'इमेज अपलोड विफल');
        }

        const data = await response.json();
        return data.data_url;
    }

    // Fetch Live Color Parchi Preview HTML
    async function fetchColorParchiPreview() {
        if (!colorState.sessionId) return;

        setColorLoading(true, 'कलर पर्ची पूर्वावलोकन तैयार हो रहा है...');
        try {
            const layoutType = colorLayoutTypeSelect ? colorLayoutTypeSelect.value : '8';
            const showPageNum = colorShowPageNumCheck ? colorShowPageNumCheck.checked : true;
            const filterActive = colorActiveOnlyCheck ? colorActiveOnlyCheck.checked : true;

            const payload = {
                session_id: colorState.sessionId,
                candidate_info: gatherCandidateInfo(),
                layout_type: layoutType,
                show_page_number: showPageNum,
                filter_active_only: filterActive,
                page: colorState.currentPage
            };

            const response = await fetch('/api/color_parchi/preview_html', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || 'पूर्वावलोकन प्राप्त नहीं हो सका');
            }

            const res = await response.json();
            colorState.totalPages = res.total_pages || 1;
            colorState.currentPage = res.current_page || 1;

            if (colorParchiA4Paper) {
                const styleMatch = res.html.match(/<style[^>]*>([\s\S]*?)<\/style>/i);
                const bodyMatch = res.html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
                const styles = styleMatch ? `<style>${styleMatch[1]}</style>` : '';
                const bodyContent = bodyMatch ? bodyMatch[1] : res.html;
                colorParchiA4Paper.innerHTML = styles + bodyContent;
            }

            if (colorWorkspace) colorWorkspace.classList.remove('hidden');
            if (colorTotalBadge) colorTotalBadge.textContent = `${res.total_voters} मतदाता सम्मिलित`;
            if (colorPagesBadge) colorPagesBadge.textContent = `${res.total_pages} A4 पेजेस`;
            if (colorCurrentPageText) colorCurrentPageText.textContent = colorState.currentPage;
            if (colorTotalPagesText) colorTotalPagesText.textContent = colorState.totalPages;

            const cardsPerPage = layoutType === '12' ? 12 : 8;
            const startIdx = (colorState.currentPage - 1) * cardsPerPage + 1;
            const endIdx = Math.min(colorState.currentPage * cardsPerPage, res.total_voters);
            if (colorVotersRangeText) {
                colorVotersRangeText.textContent = res.total_voters > 0 ? `(पर्ची ${startIdx} से ${endIdx})` : '';
            }

            if (colorPrevPageBtn) colorPrevPageBtn.disabled = (colorState.currentPage <= 1);
            if (colorNextPageBtn) colorNextPageBtn.disabled = (colorState.currentPage >= colorState.totalPages);

            if (window.lucide) lucide.createIcons();
        } catch (err) {
            alert('कलर पर्ची त्रुटि: ' + err.message);
        } finally {
            setColorLoading(false);
        }
    }

    if (colorRefreshBtn) {
        colorRefreshBtn.addEventListener('click', () => {
            fetchColorParchiPreview();
        });
    }

    // Pagination Buttons
    if (colorPrevPageBtn) {
        colorPrevPageBtn.addEventListener('click', () => {
            if (colorState.currentPage > 1) {
                colorState.currentPage--;
                fetchColorParchiPreview();
            }
        });
    }

    if (colorNextPageBtn) {
        colorNextPageBtn.addEventListener('click', () => {
            if (colorState.currentPage < colorState.totalPages) {
                colorState.currentPage++;
                fetchColorParchiPreview();
            }
        });
    }

    // Zoom Toggle
    if (colorZoomToggleBtn && colorParchiA4Paper) {
        colorZoomToggleBtn.addEventListener('click', () => {
            colorState.isZoomed = !colorState.isZoomed;
            if (colorState.isZoomed) {
                colorParchiA4Paper.classList.remove('a4-scaled');
                colorZoomBtnText.textContent = 'फ़िट टू स्क्रीन';
            } else {
                colorParchiA4Paper.classList.add('a4-scaled');
                colorZoomBtnText.textContent = '100% ज़ूम';
            }
        });
    }

    // Direct Browser Print Action
    if (colorPrintDirectBtn) {
        colorPrintDirectBtn.addEventListener('click', async () => {
            if (!colorState.sessionId) {
                alert('कृपया पहले कोई एक्सेल फ़ाइल अपलोड करें या सत्र लोड करें');
                return;
            }

            // Sync latest candidate settings with backend cache
            try {
                const layoutType = colorLayoutTypeSelect ? colorLayoutTypeSelect.value : '8';
                const showPageNum = colorShowPageNumCheck ? colorShowPageNumCheck.checked : true;
                const filterActive = colorActiveOnlyCheck ? colorActiveOnlyCheck.checked : true;

                await fetch('/api/color_parchi/preview_html', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        session_id: colorState.sessionId,
                        candidate_info: gatherCandidateInfo(),
                        layout_type: layoutType,
                        show_page_number: showPageNum,
                        filter_active_only: filterActive,
                        page: 1
                    })
                });

                const cand = gatherCandidateInfo();
                const jp = encodeURIComponent(cand.jila_parishad || '');
                const ps = encodeURIComponent(cand.panchayat_samiti || '');
                const printUrl = `/color_parchi/print?session_id=${colorState.sessionId}&layout_type=${layoutType}&show_page_number=${showPageNum}&filter_active_only=${filterActive}&jila_parishad=${jp}&panchayat_samiti=${ps}`;
                const printWin = window.open(printUrl, '_blank');
                if (!printWin) {
                    alert('कृपया अपने ब्राउज़र में पॉप-अप (Pop-ups) की अनुमति दें ताकि प्रिंट डायलॉग खुल सके।');
                }
            } catch (err) {
                alert('प्रिंट डायलॉग खोलने में विफल: ' + err.message);
            }
        });
    }

    // Vector PDF Download Action
    if (colorDownloadPdfBtn) {
        colorDownloadPdfBtn.addEventListener('click', async () => {
            if (!colorState.sessionId) {
                alert('कृपया पहले कोई एक्सेल फ़ाइल अपलोड करें या सत्र लोड करें');
                return;
            }

            colorDownloadPdfBtn.disabled = true;
            const originalHtml = colorDownloadPdfBtn.innerHTML;
            colorDownloadPdfBtn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i> <span>कलर PDF जनरेट हो रही है...</span>';
            if (window.lucide) lucide.createIcons();

            try {
                const layoutType = colorLayoutTypeSelect ? colorLayoutTypeSelect.value : '8';
                const showPageNum = colorShowPageNumCheck ? colorShowPageNumCheck.checked : true;
                const filterActive = colorActiveOnlyCheck ? colorActiveOnlyCheck.checked : true;

                const payload = {
                    session_id: colorState.sessionId,
                    candidate_info: gatherCandidateInfo(),
                    layout_type: layoutType,
                    show_page_number: showPageNum,
                    filter_active_only: filterActive
                };

                const response = await fetch('/api/color_parchi/export_pdf', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.error || 'PDF डाउनलोड विफल');
                }

                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                const safeName = (colorCandidateNameInput?.value.trim() || 'Candidate').replace(/[^\w\s-]/g, '').trim() || 'Color_Parchi';
                a.download = `Geam_Digital_${safeName}_A4_${layoutType}_Color_Parchi.pdf`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            } catch (err) {
                alert('PDF निर्यात त्रुटि: ' + err.message);
            } finally {
                colorDownloadPdfBtn.disabled = false;
                colorDownloadPdfBtn.innerHTML = originalHtml;
                if (window.lucide) lucide.createIcons();
            }
        });
    }

    // Helper: Set Color Loading
    function setColorLoading(isLoading, text = '') {
        if (!colorLoading) return;
        if (isLoading) {
            colorLoading.classList.remove('hidden');
            if (text && colorLoading.querySelector('p')) {
                colorLoading.querySelector('p').textContent = text;
            }
            if (colorRefreshBtn) colorRefreshBtn.disabled = true;
        } else {
            colorLoading.classList.add('hidden');
            if (colorRefreshBtn) colorRefreshBtn.disabled = false;
        }
    }

    // ==============================================================
    // FEATURE 4: ALPHABETICAL (ABCD) VOTER LIST CLIENT LOGIC
    // ==============================================================

    const alphaState = {
        sessionId: null,
        sheetNames: [],
        activeSheet: '',
        wards: [],
        selectedWard: 'all',
        totalVoters: 0,
        activeCount: 0,
        deletedCount: 0,
        currentPage: 1,
        totalPages: 1,
        isZoomed: false,
        selectedFile: null,
        jilaParishad: '',
        panchayatSamiti: ''
    };

    // DOM Elements - Feature 4
    const alphaDropZone = document.getElementById('alphaDropZone');
    const alphaFileInput = document.getElementById('alphaFileInput');
    const alphaLoadSessionBtn = document.getElementById('alphaLoadSessionBtn');
    const alphaFileSelectedBox = document.getElementById('alphaFileSelectedBox');
    const alphaSelectedFileName = document.getElementById('alphaSelectedFileName');
    const alphaSelectedFileSize = document.getElementById('alphaSelectedFileSize');
    const alphaStartProcessBtn = document.getElementById('alphaStartProcessBtn');
    const alphaStartProcessBtnText = document.getElementById('alphaStartProcessBtnText');
    const alphaConfigBox = document.getElementById('alphaConfigBox');

    // Controls
    const alphaPanchayatInput = document.getElementById('alphaPanchayatInput');
    const alphaListTitleInput = document.getElementById('alphaListTitleInput');
    const alphaWardSelect = document.getElementById('alphaWardSelect');
    const alphaBoothInput = document.getElementById('alphaBoothInput');
    const alphaSheetSelect = document.getElementById('alphaSheetSelect');
    const alphaActiveOnlyCheck = document.getElementById('alphaActiveOnlyCheck');
    const alphaRefreshBtn = document.getElementById('alphaRefreshBtn');
    const alphaJilaParishadInput = document.getElementById('alphaJilaParishadInput');
    const alphaPanchayatSamitiInput = document.getElementById('alphaPanchayatSamitiInput');
    const alphaLoading = document.getElementById('alphaLoading');

    // Workspace & Outputs
    const alphaWorkspace = document.getElementById('alphaWorkspace');
    const alphaTotalBadge = document.getElementById('alphaTotalBadge');
    const alphaPagesBadge = document.getElementById('alphaPagesBadge');
    const alphaActiveWardBadge = document.getElementById('alphaActiveWardBadge');
    const alphaPrintDirectBtn = document.getElementById('alphaPrintDirectBtn');
    const alphaDownloadPdfBtn = document.getElementById('alphaDownloadPdfBtn');
    const alphaDownloadPdfBtnText = document.getElementById('alphaDownloadPdfBtnText');
    const alphaDownloadExcelBtn = document.getElementById('alphaDownloadExcelBtn');
    const alphaPrevPageBtn = document.getElementById('alphaPrevPageBtn');
    const alphaNextPageBtn = document.getElementById('alphaNextPageBtn');
    const alphaCurrentPageText = document.getElementById('alphaCurrentPageText');
    const alphaTotalPagesText = document.getElementById('alphaTotalPagesText');
    const alphaWardInfoBadge = document.getElementById('alphaWardInfoBadge');
    const alphaVotersRangeText = document.getElementById('alphaVotersRangeText');
    const alphaZoomToggleBtn = document.getElementById('alphaZoomToggleBtn');
    const alphaZoomBtnText = document.getElementById('alphaZoomBtnText');
    const alphaA4Paper = document.getElementById('alphaA4Paper');

    // Helper: Set Alpha Loading
    function setAlphaLoading(isLoading, text = '') {
        if (!alphaLoading) return;
        if (isLoading) {
            alphaLoading.classList.remove('hidden');
            alphaLoading.classList.add('flex');
            if (text && alphaLoading.querySelector('p')) {
                alphaLoading.querySelector('p').textContent = text;
            }
            if (alphaRefreshBtn) alphaRefreshBtn.disabled = true;
        } else {
            alphaLoading.classList.add('hidden');
            alphaLoading.classList.remove('flex');
            if (alphaRefreshBtn) alphaRefreshBtn.disabled = false;
        }
    }

    // Drag & Drop for Feature 4
    if (alphaDropZone) {
        alphaDropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            alphaDropZone.classList.add('border-amber-500', 'bg-amber-50/40');
        });

        alphaDropZone.addEventListener('dragleave', () => {
            alphaDropZone.classList.remove('border-amber-500', 'bg-amber-50/40');
        });

        alphaDropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            alphaDropZone.classList.remove('border-amber-500', 'bg-amber-50/40');
            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                handleAlphaFileChosen(e.dataTransfer.files[0]);
            }
        });

        alphaFileInput?.addEventListener('change', () => {
            if (alphaFileInput.files && alphaFileInput.files.length > 0) {
                handleAlphaFileChosen(alphaFileInput.files[0]);
            }
        });
    }

    function handleAlphaFileChosen(file) {
        const name = file.name.toLowerCase();
        if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
            alert('कृपया केवल वैध .xlsx या .xls एक्सेल फ़ाइल चुनें');
            return;
        }

        alphaState.selectedFile = file;

        if (alphaSelectedFileName) alphaSelectedFileName.textContent = file.name;
        if (alphaSelectedFileSize) alphaSelectedFileSize.textContent = `${(file.size / 1024).toFixed(1)} KB`;

        if (alphaFileSelectedBox) {
            alphaFileSelectedBox.classList.remove('hidden');
            alphaFileSelectedBox.classList.add('flex');
        }

        if (alphaStartProcessBtnText) {
            alphaStartProcessBtnText.textContent = `🚀 "${file.name}" से अल्फाबेटिक लिस्ट प्रोसेस करें`;
        }
    }

    // Start Process Button Handler
    if (alphaStartProcessBtn) {
        alphaStartProcessBtn.addEventListener('click', () => {
            if (alphaState.selectedFile) {
                executeUploadAlphaExcel(alphaState.selectedFile);
            } else if (alphaState.sessionId) {
                fetchAlphaPreview();
            } else {
                alert('कृपया पहले कोई एक्सेल फ़ाइल चुनें या वर्तमान सत्र से डेटा लोड करें');
            }
        });
    }

    // Upload Excel Execution
    async function executeUploadAlphaExcel(file) {
        setAlphaLoading(true, `एक्सेल फ़ाइल '${file.name}' पार्स की जा रही है...`);
        if (alphaStartProcessBtn) alphaStartProcessBtn.disabled = true;

        try {
            const formData = new FormData();
            formData.append('excel_file', file);

            const response = await fetch('/api/alphabetical/upload_excel', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || 'एक्सेल फ़ाइल लोड करने में विफल');
            }

            const data = await response.json();
            applyAlphaData(data);
        } catch (err) {
            alert('त्रुटि: ' + err.message);
        } finally {
            setAlphaLoading(false);
            if (alphaStartProcessBtn) alphaStartProcessBtn.disabled = false;
        }
    }

    // Load from Session Button Handler
    if (alphaLoadSessionBtn) {
        alphaLoadSessionBtn.addEventListener('click', () => {
            executeLoadAlphaFromSession();
        });
    }

    async function executeLoadAlphaFromSession() {
        setAlphaLoading(true, 'वर्तमान सत्र से मतदाता डेटा लोड किया जा रहा है...');
        if (alphaStartProcessBtn) alphaStartProcessBtn.disabled = true;

        try {
            const response = await fetch('/api/alphabetical/from_session', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ session_id: state.sessionId })
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || 'सत्र डेटा प्राप्त नहीं हो सका');
            }

            const data = await response.json();
            applyAlphaData(data);
        } catch (err) {
            alert('सूचना: ' + err.message);
        } finally {
            setAlphaLoading(false);
            if (alphaStartProcessBtn) alphaStartProcessBtn.disabled = false;
        }
    }

    // Apply Parsed Data to Alpha State & UI
    function applyAlphaData(data) {
        alphaState.sessionId = data.session_id;
        alphaState.sheetNames = data.sheet_names || [];
        alphaState.activeSheet = data.active_sheet || '';
        alphaState.wards = data.wards || [];
        alphaState.totalVoters = data.total_voters || 0;
        alphaState.activeCount = data.active_count || 0;
        alphaState.deletedCount = data.deleted_count || 0;
        alphaState.currentPage = 1;
        alphaState.jilaParishad = data.default_jila_parishad || '';
        alphaState.panchayatSamiti = data.default_panchayat_samiti || '';

        if (data.default_panchayat_name && alphaPanchayatInput) {
            alphaPanchayatInput.value = data.default_panchayat_name;
        }
        if (data.default_booth_address && alphaBoothInput) {
            alphaBoothInput.value = data.default_booth_address;
        }
        if (alphaJilaParishadInput) {
            alphaJilaParishadInput.value = alphaState.jilaParishad;
        }
        if (alphaPanchayatSamitiInput) {
            alphaPanchayatSamitiInput.value = alphaState.panchayatSamiti;
        }

        // Populate Sheet Select
        if (alphaSheetSelect) {
            alphaSheetSelect.innerHTML = '';
            alphaState.sheetNames.forEach(s => {
                const opt = document.createElement('option');
                opt.value = s;
                opt.textContent = s;
                if (s === alphaState.activeSheet) opt.selected = true;
                alphaSheetSelect.appendChild(opt);
            });
        }

        // Populate Ward Select
        if (alphaWardSelect) {
            alphaWardSelect.innerHTML = '';
            const allOpt = document.createElement('option');
            allOpt.value = 'all';
            allOpt.textContent = 'समस्त वार्ड (वार्ड अनुसार क्रमिक 1, 2, 3...) - Ward-Wise';
            allOpt.selected = true;
            alphaWardSelect.appendChild(allOpt);

            alphaState.wards.forEach(w => {
                const opt = document.createElement('option');
                opt.value = w;
                opt.textContent = `वार्ड नं. ${w} (भाग संख्या ${w})`;
                alphaWardSelect.appendChild(opt);
            });
        }

        fetchAlphaPreview();
        refreshHistoryPills();
    }

    // Sheet Selection Change
    if (alphaSheetSelect) {
        alphaSheetSelect.addEventListener('change', async () => {
            const sheetName = alphaSheetSelect.value;
            if (!sheetName || !alphaState.sessionId) return;

            setAlphaLoading(true, `शीट '${sheetName}' लोड हो रही है...`);
            try {
                const response = await fetch('/api/alphabetical/select_sheet', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        session_id: alphaState.sessionId,
                        sheet_name: sheetName
                    })
                });

                if (!response.ok) throw new Error('शीट लोड करने में विफल');
                const data = await response.json();
                applyAlphaData(data);
            } catch (err) {
                alert(err.message);
            } finally {
                setAlphaLoading(false);
            }
        });
    }

    // Ward Selection Change
    if (alphaWardSelect) {
        alphaWardSelect.addEventListener('change', () => {
            const sel = alphaWardSelect.value;
            alphaState.selectedWard = sel;
            alphaState.currentPage = 1;
            fetchAlphaPreview();
        });
    }

    // Fetch Live Alpha Preview HTML
    async function fetchAlphaPreview() {
        if (!alphaState.sessionId) return;

        setAlphaLoading(true, 'अल्फाबेटिक वोटर लिस्ट पूर्वावलोकन तैयार हो रहा है...');
        try {
            const wardTitle = alphaPanchayatInput ? alphaPanchayatInput.value.trim() : 'ग्राम पंचायत';
            const listTitle = alphaListTitleInput ? alphaListTitleInput.value.trim() : 'अल्फाबेटिक ABCD से वोटर लिस्ट';
            const selWard = alphaWardSelect ? alphaWardSelect.value : 'all';
            const partNo = selWard !== 'all' ? selWard : '';
            const boothAddr = alphaBoothInput ? alphaBoothInput.value.trim() : '';
            const filterActive = alphaActiveOnlyCheck ? alphaActiveOnlyCheck.checked : true;
            const jp = alphaJilaParishadInput ? alphaJilaParishadInput.value.trim() : (alphaState.jilaParishad || '');
            const ps = alphaPanchayatSamitiInput ? alphaPanchayatSamitiInput.value.trim() : (alphaState.panchayatSamiti || '');

            const payload = {
                session_id: alphaState.sessionId,
                ward_title: wardTitle,
                list_title: listTitle,
                part_no: partNo,
                booth_address: boothAddr,
                selected_ward: selWard,
                filter_active_only: filterActive,
                jila_parishad: jp,
                panchayat_samiti: ps,
                page: alphaState.currentPage
            };

            const response = await fetch('/api/alphabetical/preview_html', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || 'पूर्वावलोकन प्राप्त नहीं हो सका');
            }

            const res = await response.json();
            alphaState.totalPages = res.total_pages || 1;
            alphaState.currentPage = res.current_page || 1;

            if (alphaA4Paper) {
                const styleMatch = res.html.match(/<style[^>]*>([\s\S]*?)<\/style>/i);
                const bodyMatch = res.html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
                const styles = styleMatch ? `<style>${styleMatch[1]}</style>` : '';
                const bodyContent = bodyMatch ? bodyMatch[1] : res.html;
                alphaA4Paper.innerHTML = styles + bodyContent;
            }

            if (alphaWorkspace) alphaWorkspace.classList.remove('hidden');
            if (alphaTotalBadge) alphaTotalBadge.textContent = `${res.total_voters} मतदाता`;
            if (alphaPagesBadge) alphaPagesBadge.textContent = `${res.total_pages} A4 पेजेस`;
            if (alphaActiveWardBadge) {
                if (selWard === 'all') {
                    alphaActiveWardBadge.textContent = `समस्त वार्ड (वार्ड 1 से ${res.total_wards || alphaState.wards.length} क्रमिक)`;
                } else {
                    alphaActiveWardBadge.textContent = `वार्ड नं.- ${res.current_ward || selWard}`;
                }
            }
            if (alphaCurrentPageText) alphaCurrentPageText.textContent = alphaState.currentPage;
            if (alphaTotalPagesText) alphaTotalPagesText.textContent = alphaState.totalPages;

            if (alphaWardInfoBadge) {
                alphaWardInfoBadge.textContent = `वार्ड ${res.current_ward || 1}`;
            }
            if (alphaVotersRangeText) {
                alphaVotersRangeText.innerHTML = `[वार्ड ${res.current_ward || 1}: पेज ${res.page_in_ward || 1} / ${res.total_in_ward || 1} • कुल ${res.ward_voters || 0} मतदाता]`;
            }

            if (alphaPrevPageBtn) alphaPrevPageBtn.disabled = (alphaState.currentPage <= 1);
            if (alphaNextPageBtn) alphaNextPageBtn.disabled = (alphaState.currentPage >= alphaState.totalPages);

            if (window.lucide) lucide.createIcons();
        } catch (err) {
            alert('अल्फाबेटिक लिस्ट त्रुटि: ' + err.message);
        } finally {
            setAlphaLoading(false);
        }
    }

    if (alphaRefreshBtn) {
        alphaRefreshBtn.addEventListener('click', () => {
            fetchAlphaPreview();
        });
    }

    // Pagination Buttons
    if (alphaPrevPageBtn) {
        alphaPrevPageBtn.addEventListener('click', () => {
            if (alphaState.currentPage > 1) {
                alphaState.currentPage--;
                fetchAlphaPreview();
            }
        });
    }

    if (alphaNextPageBtn) {
        alphaNextPageBtn.addEventListener('click', () => {
            if (alphaState.currentPage < alphaState.totalPages) {
                alphaState.currentPage++;
                fetchAlphaPreview();
            }
        });
    }

    // Zoom Toggle
    if (alphaZoomToggleBtn && alphaA4Paper) {
        alphaZoomToggleBtn.addEventListener('click', () => {
            alphaState.isZoomed = !alphaState.isZoomed;
            if (alphaState.isZoomed) {
                alphaA4Paper.classList.remove('a4-scaled');
                alphaZoomBtnText.textContent = 'फ़िट टू स्क्रीन';
            } else {
                alphaA4Paper.classList.add('a4-scaled');
                alphaZoomBtnText.textContent = '100% ज़ूम';
            }
        });
    }

    // Direct Browser Print Action
    if (alphaPrintDirectBtn) {
        alphaPrintDirectBtn.addEventListener('click', async () => {
            if (!alphaState.sessionId) {
                alert('कृपया पहले कोई एक्सेल फ़ाइल अपलोड करें या सत्र लोड करें');
                return;
            }

            const wardTitle = encodeURIComponent(alphaPanchayatInput ? alphaPanchayatInput.value.trim() : 'ग्राम पंचायत');
            const listTitle = encodeURIComponent(alphaListTitleInput ? alphaListTitleInput.value.trim() : 'अल्फाबेटिक ABCD से वोटर लिस्ट');
            const selWard = encodeURIComponent(alphaWardSelect ? alphaWardSelect.value : 'all');
            const partNo = encodeURIComponent(alphaWardSelect && alphaWardSelect.value !== 'all' ? alphaWardSelect.value : '');
            const boothAddr = encodeURIComponent(alphaBoothInput ? alphaBoothInput.value.trim() : '');
            const filterActive = alphaActiveOnlyCheck ? alphaActiveOnlyCheck.checked : true;
            const jp = encodeURIComponent((alphaJilaParishadInput ? alphaJilaParishadInput.value.trim() : alphaState.jilaParishad) || '');
            const ps = encodeURIComponent((alphaPanchayatSamitiInput ? alphaPanchayatSamitiInput.value.trim() : alphaState.panchayatSamiti) || '');

            const printUrl = `/alphabetical/print?session_id=${alphaState.sessionId}&ward_title=${wardTitle}&list_title=${listTitle}&part_no=${partNo}&booth_address=${boothAddr}&selected_ward=${selWard}&filter_active_only=${filterActive}&jila_parishad=${jp}&panchayat_samiti=${ps}`;
            const printWin = window.open(printUrl, '_blank');
            if (!printWin) {
                alert('कृपया अपने ब्राउज़र में पॉप-अप (Pop-ups) की अनुमति दें ताकि प्रिंट डायलॉग खुल सके।');
            }
        });
    }

    // Vector PDF Download Action
    if (alphaDownloadPdfBtn) {
        alphaDownloadPdfBtn.addEventListener('click', async () => {
            if (!alphaState.sessionId) {
                alert('कृपया पहले कोई एक्सेल फ़ाइल अपलोड करें या सत्र लोड करें');
                return;
            }

            alphaDownloadPdfBtn.disabled = true;
            const originalHtml = alphaDownloadPdfBtn.innerHTML;
            alphaDownloadPdfBtn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i> <span>PDF जनरेट हो रही है...</span>';
            if (window.lucide) lucide.createIcons();

            try {
                const wardTitle = alphaPanchayatInput ? alphaPanchayatInput.value.trim() : 'ग्राम पंचायत';
                const listTitle = alphaListTitleInput ? alphaListTitleInput.value.trim() : 'अल्फाबेटिक ABCD से वोटर लिस्ट';
                const selWard = alphaWardSelect ? alphaWardSelect.value : 'all';
                const partNo = selWard !== 'all' ? selWard : '';
                const boothAddr = alphaBoothInput ? alphaBoothInput.value.trim() : '';
                const filterActive = alphaActiveOnlyCheck ? alphaActiveOnlyCheck.checked : true;
                const jp = alphaJilaParishadInput ? alphaJilaParishadInput.value.trim() : (alphaState.jilaParishad || '');
                const ps = alphaPanchayatSamitiInput ? alphaPanchayatSamitiInput.value.trim() : (alphaState.panchayatSamiti || '');

                const payload = {
                    session_id: alphaState.sessionId,
                    ward_title: wardTitle,
                    list_title: listTitle,
                    part_no: partNo,
                    booth_address: boothAddr,
                    selected_ward: selWard,
                    filter_active_only: filterActive,
                    jila_parishad: jp,
                    panchayat_samiti: ps
                };

                const response = await fetch('/api/alphabetical/export_pdf', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.error || 'PDF डाउनलोड विफल');
                }

                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                const safeName = (wardTitle || 'Alphabetical_List').replace(/[^\w\s-]/g, '').trim() || 'Alphabetical_List';
                a.download = `Geam_Digital_${safeName}_ABCD_Voter_List.pdf`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            } catch (err) {
                alert('PDF निर्यात त्रुटि: ' + err.message);
            } finally {
                alphaDownloadPdfBtn.disabled = false;
                alphaDownloadPdfBtn.innerHTML = originalHtml;
                if (window.lucide) lucide.createIcons();
            }
        });
    }

    // Sorted Excel Download Action
    if (alphaDownloadExcelBtn) {
        alphaDownloadExcelBtn.addEventListener('click', async () => {
            if (!alphaState.sessionId) {
                alert('कृपया पहले कोई एक्सेल फ़ाइल अपलोड करें या सत्र लोड करें');
                return;
            }

            alphaDownloadExcelBtn.disabled = true;
            const originalHtml = alphaDownloadExcelBtn.innerHTML;
            alphaDownloadExcelBtn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i> <span>Excel तैयार हो रहा है...</span>';
            if (window.lucide) lucide.createIcons();

            try {
                const wardTitle = alphaPanchayatInput ? alphaPanchayatInput.value.trim() : 'ग्राम पंचायत';
                const listTitle = alphaListTitleInput ? alphaListTitleInput.value.trim() : 'अल्फाबेटिक ABCD से वोटर लिस्ट';
                const selWard = alphaWardSelect ? alphaWardSelect.value : 'all';
                const partNo = selWard !== 'all' ? selWard : '';
                const boothAddr = alphaBoothInput ? alphaBoothInput.value.trim() : '';
                const filterActive = alphaActiveOnlyCheck ? alphaActiveOnlyCheck.checked : true;
                const jp = alphaJilaParishadInput ? alphaJilaParishadInput.value.trim() : (alphaState.jilaParishad || '');
                const ps = alphaPanchayatSamitiInput ? alphaPanchayatSamitiInput.value.trim() : (alphaState.panchayatSamiti || '');

                const payload = {
                    session_id: alphaState.sessionId,
                    ward_title: wardTitle,
                    list_title: listTitle,
                    part_no: partNo,
                    booth_address: boothAddr,
                    selected_ward: selWard,
                    filter_active_only: filterActive,
                    jila_parishad: jp,
                    panchayat_samiti: ps
                };

                const response = await fetch('/api/alphabetical/export_excel', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.error || 'Excel डाउनलोड विफल');
                }

                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                const safeName = (wardTitle || 'Alphabetical_List').replace(/[^\w\s-]/g, '').trim() || 'Alphabetical_List';
                a.download = `Geam_Digital_${safeName}_ABCD_Voter_List.xlsx`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            } catch (err) {
                alert('Excel निर्यात त्रुटि: ' + err.message);
            } finally {
                alphaDownloadExcelBtn.disabled = false;
                alphaDownloadExcelBtn.innerHTML = originalHtml;
                if (window.lucide) lucide.createIcons();
            }
        });
    }

    // ==============================================================
    // FEATURE 5: BADE PARIVAR / BADE GHAR VOTER LIST JAVASCRIPT
    // ==============================================================

    const familyState = {
        sessionId: null,
        currentPage: 1,
        totalPages: 1,
        totalVoters: 0,
        totalFamilies: 0,
        selectedWard: 'all',
        minFamilySize: 2,
        ignoreZeroHouses: true,
        isZoomed: false,
        selectedFile: null,
        jilaParishad: '',
        panchayatSamiti: ''
    };

    // DOM Elements - Feature 5
    const familyDropZone = document.getElementById('familyDropZone');
    const familyFileInput = document.getElementById('familyFileInput');
    const familyLoadSessionBtn = document.getElementById('familyLoadSessionBtn');
    const familyFileSelectedBox = document.getElementById('familyFileSelectedBox');
    const familySelectedFileName = document.getElementById('familySelectedFileName');
    const familySelectedFileSize = document.getElementById('familySelectedFileSize');
    const familyStartProcessBtn = document.getElementById('familyStartProcessBtn');
    const familyStartProcessBtnText = document.getElementById('familyStartProcessBtnText');
    const familyConfigBox = document.getElementById('familyConfigBox');

    // Controls
    const familyPanchayatInput = document.getElementById('familyPanchayatInput');
    const familyListTitleInput = document.getElementById('familyListTitleInput');
    const familyBoothInput = document.getElementById('familyBoothInput');
    const familyMinSizeSelect = document.getElementById('familyMinSizeSelect');
    const familyWardSelect = document.getElementById('familyWardSelect');
    const familySheetSelect = document.getElementById('familySheetSelect');
    const familyJilaParishadInput = document.getElementById('familyJilaParishadInput');
    const familyPanchayatSamitiInput = document.getElementById('familyPanchayatSamitiInput');
    const familyIgnoreZeroCheck = document.getElementById('familyIgnoreZeroCheck');
    const familyActiveOnlyCheck = document.getElementById('familyActiveOnlyCheck');
    const familyRefreshBtn = document.getElementById('familyRefreshBtn');
    const familyLoading = document.getElementById('familyLoading');

    // Workspace & Outputs
    const familyWorkspace = document.getElementById('familyWorkspace');
    const familyFamiliesBadge = document.getElementById('familyFamiliesBadge');
    const familyTotalBadge = document.getElementById('familyTotalBadge');
    const familyPagesBadge = document.getElementById('familyPagesBadge');
    const familyMaxBadge = document.getElementById('familyMaxBadge');
    const familyPrintDirectBtn = document.getElementById('familyPrintDirectBtn');
    const familyDownloadPdfBtn = document.getElementById('familyDownloadPdfBtn');
    const familyDownloadPdfBtnText = document.getElementById('familyDownloadPdfBtnText');
    const familyDownloadExcelBtn = document.getElementById('familyDownloadExcelBtn');
    const familyPrevPageBtn = document.getElementById('familyPrevPageBtn');
    const familyNextPageBtn = document.getElementById('familyNextPageBtn');
    const familyCurrentPageText = document.getElementById('familyCurrentPageText');
    const familyTotalPagesText = document.getElementById('familyTotalPagesText');
    const familyWardInfoBadge = document.getElementById('familyWardInfoBadge');
    const familyWardStatsText = document.getElementById('familyWardStatsText');
    const familyZoomToggleBtn = document.getElementById('familyZoomToggleBtn');
    const familyZoomBtnText = document.getElementById('familyZoomBtnText');
    const familyA4Paper = document.getElementById('familyA4Paper');

    // Helper: Set Family Loading
    function setFamilyLoading(isLoading, text = '') {
        if (!familyLoading) return;
        if (isLoading) {
            familyLoading.classList.remove('hidden');
            familyLoading.classList.add('flex');
            if (text && familyLoading.querySelector('p')) {
                familyLoading.querySelector('p').textContent = text;
            }
            if (familyRefreshBtn) familyRefreshBtn.disabled = true;
        } else {
            familyLoading.classList.add('hidden');
            familyLoading.classList.remove('flex');
            if (familyRefreshBtn) familyRefreshBtn.disabled = false;
        }
    }

    // Drag & Drop for Feature 5
    if (familyDropZone) {
        familyDropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            familyDropZone.classList.add('border-purple-500', 'bg-purple-50/40');
        });

        familyDropZone.addEventListener('dragleave', () => {
            familyDropZone.classList.remove('border-purple-500', 'bg-purple-50/40');
        });

        familyDropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            familyDropZone.classList.remove('border-purple-500', 'bg-purple-50/40');
            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                handleFamilyFileChosen(e.dataTransfer.files[0]);
            }
        });

        familyFileInput?.addEventListener('change', () => {
            if (familyFileInput.files && familyFileInput.files.length > 0) {
                handleFamilyFileChosen(familyFileInput.files[0]);
            }
        });
    }

    function handleFamilyFileChosen(file) {
        const name = file.name.toLowerCase();
        if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
            alert('कृपया केवल वैध .xlsx या .xls एक्सेल फ़ाइल चुनें');
            return;
        }

        familyState.selectedFile = file;

        if (familySelectedFileName) familySelectedFileName.textContent = file.name;
        if (familySelectedFileSize) familySelectedFileSize.textContent = `${(file.size / 1024).toFixed(1)} KB`;

        if (familyFileSelectedBox) {
            familyFileSelectedBox.classList.remove('hidden');
            familyFileSelectedBox.classList.add('flex');
        }

        if (familyStartProcessBtnText) {
            familyStartProcessBtnText.textContent = `🚀 "${file.name}" से बड़े-घर की लिस्ट प्रोसेस करें`;
        }
    }

    // Start Process Button Handler
    if (familyStartProcessBtn) {
        familyStartProcessBtn.addEventListener('click', () => {
            if (familyState.selectedFile) {
                executeUploadFamilyExcel(familyState.selectedFile);
            } else if (familyState.sessionId) {
                fetchFamilyPreview();
            } else {
                alert('कृपया पहले कोई एक्सेल फ़ाइल चुनें या वर्तमान सत्र से डेटा लोड करें');
            }
        });
    }

    // Upload Excel Execution
    async function executeUploadFamilyExcel(file) {
        setFamilyLoading(true, `एक्सेल फ़ाइल '${file.name}' पार्स की जा रही है...`);
        if (familyStartProcessBtn) familyStartProcessBtn.disabled = true;

        try {
            const formData = new FormData();
            formData.append('excel_file', file);

            const response = await fetch('/api/family/upload_excel', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || 'एक्सेल अपलोड विफल रहा');
            }

            const data = await response.json();
            applyFamilyData(data);
        } catch (err) {
            alert('एक्सेल त्रुटि: ' + err.message);
        } finally {
            setFamilyLoading(false);
            if (familyStartProcessBtn) familyStartProcessBtn.disabled = false;
        }
    }

    // Load from Session Button Handler
    if (familyLoadSessionBtn) {
        familyLoadSessionBtn.addEventListener('click', async () => {
            setFamilyLoading(true, 'वर्तमान सत्र से मतदाता डेटा लोड हो रहा है...');
            try {
                const response = await fetch('/api/family/from_session', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ session_id: currentSessionId || '' })
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.error || 'सत्र से डेटा लोड नहीं हो सका');
                }

                const data = await response.json();
                applyFamilyData(data);
            } catch (err) {
                alert('सत्र त्रुटि: ' + err.message);
            } finally {
                setFamilyLoading(false);
            }
        });
    }

    // Apply Family Data
    function applyFamilyData(data) {
        familyState.sessionId = data.session_id;
        familyState.currentPage = 1;
        familyState.jilaParishad = data.default_jila_parishad || '';
        familyState.panchayatSamiti = data.default_panchayat_samiti || '';

        if (familyJilaParishadInput) familyJilaParishadInput.value = familyState.jilaParishad;
        if (familyPanchayatSamitiInput) familyPanchayatSamitiInput.value = familyState.panchayatSamiti;

        if (familyPanchayatInput && data.default_panchayat_name) {
            familyPanchayatInput.value = data.default_panchayat_name;
        }
        if (familyBoothInput && data.default_booth_address) {
            familyBoothInput.value = data.default_booth_address;
        }

        // Sheet Selector setup
        if (familySheetSelect) {
            familySheetSelect.innerHTML = '';
            (data.sheet_names || ['Sheet1']).forEach(s => {
                const opt = document.createElement('option');
                opt.value = s;
                opt.textContent = s;
                if (s === data.active_sheet) opt.selected = true;
                familySheetSelect.appendChild(opt);
            });
        }

        // Ward Selector setup
        if (familyWardSelect) {
            familyWardSelect.innerHTML = '<option value="all" selected>समस्त वार्ड (वार्ड अनुसार क्रमिक 1, 2, 3...) - Ward-Wise</option>';
            const wardsList = (data.wards && data.wards.length > 0) ? data.wards : [];
            if (wardsList.length === 0 && data.voters_sample) {
                const wardsFound = new Set();
                (data.voters_sample || []).forEach(v => {
                    const p = v.part_no || v.ward;
                    if (p) wardsFound.add(String(p).trim());
                });
                wardsFound.forEach(w => wardsList.push(w));
            }
            wardsList.forEach(w => {
                const opt = document.createElement('option');
                opt.value = w;
                opt.textContent = `वार्ड नं.- ${w}`;
                familyWardSelect.appendChild(opt);
            });
        }

        fetchFamilyPreview();
        refreshHistoryPills();
    }

    // Sheet Selection Change
    if (familySheetSelect) {
        familySheetSelect.addEventListener('change', async () => {
            const sheetName = familySheetSelect.value;
            if (!sheetName || !familyState.sessionId) return;

            setFamilyLoading(true, `शीट '${sheetName}' लोड हो रही है...`);
            try {
                const response = await fetch('/api/family/select_sheet', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        session_id: familyState.sessionId,
                        sheet_name: sheetName
                    })
                });

                if (!response.ok) throw new Error('शीट लोड करने में विफल');
                const data = await response.json();
                applyFamilyData(data);
            } catch (err) {
                alert(err.message);
            } finally {
                setFamilyLoading(false);
            }
        });
    }

    // Ward Selection Change
    if (familyWardSelect) {
        familyWardSelect.addEventListener('change', () => {
            const sel = familyWardSelect.value;
            familyState.selectedWard = sel;
            familyState.currentPage = 1;
            fetchFamilyPreview();
        });
    }

    // Min Family Size Change
    if (familyMinSizeSelect) {
        familyMinSizeSelect.addEventListener('change', () => {
            familyState.minFamilySize = parseInt(familyMinSizeSelect.value) || 2;
            familyState.currentPage = 1;
            fetchFamilyPreview();
        });
    }

    // Checkbox Changes
    if (familyIgnoreZeroCheck) {
        familyIgnoreZeroCheck.addEventListener('change', () => {
            familyState.currentPage = 1;
            fetchFamilyPreview();
        });
    }
    if (familyActiveOnlyCheck) {
        familyActiveOnlyCheck.addEventListener('change', () => {
            familyState.currentPage = 1;
            fetchFamilyPreview();
        });
    }

    // Fetch Live Family Preview HTML
    async function fetchFamilyPreview() {
        if (!familyState.sessionId) return;

        setFamilyLoading(true, 'बड़े-घर की सूची पूर्वावलोकन तैयार हो रहा है...');
        try {
            const wardTitle = familyPanchayatInput ? familyPanchayatInput.value.trim() : 'ग्राम पंचायत';
            const listTitle = familyListTitleInput ? familyListTitleInput.value.trim() : 'बड़े-घर की लिस्ट';
            const selWard = familyWardSelect ? familyWardSelect.value : 'all';
            const partNo = selWard !== 'all' ? selWard : '';
            const boothAddr = familyBoothInput ? familyBoothInput.value.trim() : '';
            const minSize = familyMinSizeSelect ? parseInt(familyMinSizeSelect.value) : 2;
            const ignoreZero = familyIgnoreZeroCheck ? familyIgnoreZeroCheck.checked : true;
            const filterActive = familyActiveOnlyCheck ? familyActiveOnlyCheck.checked : true;
            const jp = familyJilaParishadInput ? familyJilaParishadInput.value.trim() : (familyState.jilaParishad || '');
            const ps = familyPanchayatSamitiInput ? familyPanchayatSamitiInput.value.trim() : (familyState.panchayatSamiti || '');

            const payload = {
                session_id: familyState.sessionId,
                ward_title: wardTitle,
                list_title: listTitle,
                part_no: partNo,
                booth_address: boothAddr,
                selected_ward: selWard,
                min_family_size: minSize,
                ignore_zero_houses: ignoreZero,
                filter_active_only: filterActive,
                jila_parishad: jp,
                panchayat_samiti: ps,
                page: familyState.currentPage
            };

            const response = await fetch('/api/family/preview_html', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || 'पूर्वावलोकन प्राप्त नहीं हो सका');
            }

            const res = await response.json();
            familyState.totalPages = res.total_pages || 1;
            familyState.currentPage = res.current_page || 1;
            familyState.totalVoters = res.total_voters || 0;
            familyState.totalFamilies = res.total_families || 0;

            if (familyA4Paper) {
                const styleMatch = res.html.match(/<style[^>]*>([\s\S]*?)<\/style>/i);
                const bodyMatch = res.html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
                const styles = styleMatch ? `<style>${styleMatch[1]}</style>` : '';
                const bodyContent = bodyMatch ? bodyMatch[1] : res.html;
                familyA4Paper.innerHTML = styles + bodyContent;
            }

            if (familyWorkspace) familyWorkspace.classList.remove('hidden');
            if (familyFamiliesBadge) familyFamiliesBadge.textContent = `${res.total_families} बड़े परिवार`;
            if (familyTotalBadge) familyTotalBadge.textContent = `${res.total_voters} कुल मतदाता`;
            if (familyPagesBadge) familyPagesBadge.textContent = `${res.total_pages} A4 पेजेस`;
            if (familyCurrentPageText) familyCurrentPageText.textContent = familyState.currentPage;
            if (familyTotalPagesText) familyTotalPagesText.textContent = familyState.totalPages;

            if (familyWardInfoBadge) {
                familyWardInfoBadge.textContent = `वार्ड ${res.current_ward || 1}`;
            }
            if (familyWardStatsText) {
                familyWardStatsText.innerHTML = `[वार्ड ${res.current_ward || 1}: पेज ${res.page_in_ward || 1} / ${res.total_in_ward || 1} • ${res.ward_families || 0} परिवार • ${res.ward_voters || 0} मतदाता]`;
            }

            if (familyPrevPageBtn) familyPrevPageBtn.disabled = (familyState.currentPage <= 1);
            if (familyNextPageBtn) familyNextPageBtn.disabled = (familyState.currentPage >= familyState.totalPages);

            if (window.lucide) lucide.createIcons();
        } catch (err) {
            alert('बड़े-घर लिस्ट त्रुटि: ' + err.message);
        } finally {
            setFamilyLoading(false);
        }
    }

    if (familyRefreshBtn) {
        familyRefreshBtn.addEventListener('click', () => {
            fetchFamilyPreview();
        });
    }

    // Pagination Buttons
    if (familyPrevPageBtn) {
        familyPrevPageBtn.addEventListener('click', () => {
            if (familyState.currentPage > 1) {
                familyState.currentPage--;
                fetchFamilyPreview();
            }
        });
    }

    if (familyNextPageBtn) {
        familyNextPageBtn.addEventListener('click', () => {
            if (familyState.currentPage < familyState.totalPages) {
                familyState.currentPage++;
                fetchFamilyPreview();
            }
        });
    }

    // Zoom Toggle
    if (familyZoomToggleBtn && familyA4Paper) {
        familyZoomToggleBtn.addEventListener('click', () => {
            familyState.isZoomed = !familyState.isZoomed;
            if (familyState.isZoomed) {
                familyA4Paper.classList.remove('a4-scaled');
                familyZoomBtnText.textContent = 'फ़िट टू स्क्रीन';
            } else {
                familyA4Paper.classList.add('a4-scaled');
                familyZoomBtnText.textContent = '100% ज़ूम';
            }
        });
    }

    // Direct Browser Print Action
    if (familyPrintDirectBtn) {
        familyPrintDirectBtn.addEventListener('click', async () => {
            if (!familyState.sessionId) {
                alert('कृपया पहले कोई एक्सेल फ़ाइल अपलोड करें या सत्र लोड करें');
                return;
            }

            const wardTitle = encodeURIComponent(familyPanchayatInput ? familyPanchayatInput.value.trim() : 'ग्राम पंचायत');
            const listTitle = encodeURIComponent(familyListTitleInput ? familyListTitleInput.value.trim() : 'बड़े-घर की लिस्ट');
            const selWard = encodeURIComponent(familyWardSelect ? familyWardSelect.value : 'all');
            const partNo = encodeURIComponent(familyWardSelect && familyWardSelect.value !== 'all' ? familyWardSelect.value : '');
            const boothAddr = encodeURIComponent(familyBoothInput ? familyBoothInput.value.trim() : '');
            const minSize = familyMinSizeSelect ? parseInt(familyMinSizeSelect.value) : 2;
            const ignoreZero = familyIgnoreZeroCheck ? familyIgnoreZeroCheck.checked : true;
            const filterActive = familyActiveOnlyCheck ? familyActiveOnlyCheck.checked : true;
            const jp = encodeURIComponent(familyJilaParishadInput ? familyJilaParishadInput.value.trim() : (familyState.jilaParishad || ''));
            const ps = encodeURIComponent(familyPanchayatSamitiInput ? familyPanchayatSamitiInput.value.trim() : (familyState.panchayatSamiti || ''));

            const printUrl = `/family/print?session_id=${familyState.sessionId}&ward_title=${wardTitle}&list_title=${listTitle}&part_no=${partNo}&booth_address=${boothAddr}&selected_ward=${selWard}&min_family_size=${minSize}&ignore_zero_houses=${ignoreZero}&filter_active_only=${filterActive}&jila_parishad=${jp}&panchayat_samiti=${ps}`;
            const printWin = window.open(printUrl, '_blank');
            if (!printWin) {
                alert('कृपया अपने ब्राउज़र में पॉप-अप (Pop-ups) की अनुमति दें ताकि प्रिंट डायलॉग खुल सके।');
            }
        });
    }

    // Vector PDF Download Action
    if (familyDownloadPdfBtn) {
        familyDownloadPdfBtn.addEventListener('click', async () => {
            if (!familyState.sessionId) {
                alert('कृपया पहले कोई एक्सेल फ़ाइल अपलोड करें या सत्र लोड करें');
                return;
            }

            familyDownloadPdfBtn.disabled = true;
            const originalHtml = familyDownloadPdfBtn.innerHTML;
            familyDownloadPdfBtn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i> <span>PDF जनरेट हो रही है...</span>';
            if (window.lucide) lucide.createIcons();

            try {
                const wardTitle = familyPanchayatInput ? familyPanchayatInput.value.trim() : 'ग्राम पंचायत';
                const listTitle = familyListTitleInput ? familyListTitleInput.value.trim() : 'बड़े-घर की लिस्ट';
                const selWard = familyWardSelect ? familyWardSelect.value : 'all';
                const partNo = selWard !== 'all' ? selWard : '';
                const boothAddr = familyBoothInput ? familyBoothInput.value.trim() : '';
                const minSize = familyMinSizeSelect ? parseInt(familyMinSizeSelect.value) : 2;
                const ignoreZero = familyIgnoreZeroCheck ? familyIgnoreZeroCheck.checked : true;
                const filterActive = familyActiveOnlyCheck ? familyActiveOnlyCheck.checked : true;
                const jp = familyJilaParishadInput ? familyJilaParishadInput.value.trim() : (familyState.jilaParishad || '');
                const ps = familyPanchayatSamitiInput ? familyPanchayatSamitiInput.value.trim() : (familyState.panchayatSamiti || '');

                const payload = {
                    session_id: familyState.sessionId,
                    ward_title: wardTitle,
                    list_title: listTitle,
                    part_no: partNo,
                    booth_address: boothAddr,
                    selected_ward: selWard,
                    min_family_size: minSize,
                    ignore_zero_houses: ignoreZero,
                    filter_active_only: filterActive,
                    jila_parishad: jp,
                    panchayat_samiti: ps
                };

                const response = await fetch('/api/family/export_pdf', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.error || 'PDF डाउनलोड विफल');
                }

                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                const safeName = (wardTitle || 'Family_List').replace(/[^\w\s-]/g, '').trim() || 'Family_List';
                a.download = `Geam_Digital_${safeName}_Bade_Ghar_List.pdf`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            } catch (err) {
                alert('PDF निर्यात त्रुटि: ' + err.message);
            } finally {
                familyDownloadPdfBtn.disabled = false;
                familyDownloadPdfBtn.innerHTML = originalHtml;
                if (window.lucide) lucide.createIcons();
            }
        });
    }

    // Sorted Excel Download Action
    if (familyDownloadExcelBtn) {
        familyDownloadExcelBtn.addEventListener('click', async () => {
            if (!familyState.sessionId) {
                alert('कृपया पहले कोई एक्सेल फ़ाइल अपलोड करें या सत्र लोड करें');
                return;
            }

            familyDownloadExcelBtn.disabled = true;
            const originalHtml = familyDownloadExcelBtn.innerHTML;
            familyDownloadExcelBtn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i> <span>Excel तैयार हो रहा है...</span>';
            if (window.lucide) lucide.createIcons();

            try {
                const wardTitle = familyPanchayatInput ? familyPanchayatInput.value.trim() : 'ग्राम पंचायत';
                const listTitle = familyListTitleInput ? familyListTitleInput.value.trim() : 'बड़े-घर की लिस्ट';
                const selWard = familyWardSelect ? familyWardSelect.value : 'all';
                const partNo = selWard !== 'all' ? selWard : '';
                const boothAddr = familyBoothInput ? familyBoothInput.value.trim() : '';
                const minSize = familyMinSizeSelect ? parseInt(familyMinSizeSelect.value) : 2;
                const ignoreZero = familyIgnoreZeroCheck ? familyIgnoreZeroCheck.checked : true;
                const filterActive = familyActiveOnlyCheck ? familyActiveOnlyCheck.checked : true;
                const jp = familyJilaParishadInput ? familyJilaParishadInput.value.trim() : (familyState.jilaParishad || '');
                const ps = familyPanchayatSamitiInput ? familyPanchayatSamitiInput.value.trim() : (familyState.panchayatSamiti || '');

                const payload = {
                    session_id: familyState.sessionId,
                    ward_title: wardTitle,
                    list_title: listTitle,
                    part_no: partNo,
                    booth_address: boothAddr,
                    selected_ward: selWard,
                    min_family_size: minSize,
                    ignore_zero_houses: ignoreZero,
                    filter_active_only: filterActive,
                    jila_parishad: jp,
                    panchayat_samiti: ps
                };

                const response = await fetch('/api/family/export_excel', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.error || 'Excel डाउनलोड विफल');
                }

                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                const safeName = (wardTitle || 'Family_List').replace(/[^\w\s-]/g, '').trim() || 'Family_List';
                a.download = `Geam_Digital_${safeName}_Bade_Ghar_List.xlsx`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            } catch (err) {
                alert('Excel निर्यात त्रुटि: ' + err.message);
            } finally {
                familyDownloadExcelBtn.disabled = false;
                familyDownloadExcelBtn.innerHTML = originalHtml;
                refreshHistoryPills();
                if (window.lucide) lucide.createIcons();
            }
        });
    }

    // ==============================================================
    // FEATURE 6: आयु अनुसार (युवा 18-26 एवं बुजुर्ग 120-70) वोटर लिस्ट
    // ==============================================================

    const ageState = {
        sessionId: null,
        selectedFile: null,
        sheetNames: [],
        activeSheet: '',
        wards: [],
        totalVoters: 0,
        currentPage: 1,
        totalPages: 1,
        listType: 'young', // 'young' (18-26 asc) or 'senior' (120-70 desc)
        minAge: 18,
        maxAge: 26,
        sortOrder: 'asc',
        selectedWard: 'all',
        jilaParishad: '',
        panchayatSamiti: '',
        isZoomed: false
    };

    // DOM Elements - Feature 6
    const ageDropZone = document.getElementById('ageDropZone');
    const ageFileInput = document.getElementById('ageFileInput');
    const ageLoadSessionBtn = document.getElementById('ageLoadSessionBtn');
    const ageFileSelectedBox = document.getElementById('ageFileSelectedBox');
    const ageSelectedFileName = document.getElementById('ageSelectedFileName');
    const ageSelectedFileSize = document.getElementById('ageSelectedFileSize');
    const ageStartProcessBtn = document.getElementById('ageStartProcessBtn');
    const ageStartProcessBtnText = document.getElementById('ageStartProcessBtnText');
    const ageConfigBox = document.getElementById('ageConfigBox');

    // Mode Buttons
    const ageModeYoungBtn = document.getElementById('ageModeYoungBtn');
    const ageModeSeniorBtn = document.getElementById('ageModeSeniorBtn');

    // Controls
    const agePanchayatInput = document.getElementById('agePanchayatInput');
    const ageListTitleInput = document.getElementById('ageListTitleInput');
    const ageBoothInput = document.getElementById('ageBoothInput');
    const ageWardSelect = document.getElementById('ageWardSelect');
    const ageSheetSelect = document.getElementById('ageSheetSelect');
    const ageJilaParishadInput = document.getElementById('ageJilaParishadInput');
    const agePanchayatSamitiInput = document.getElementById('agePanchayatSamitiInput');
    const ageActiveOnlyCheck = document.getElementById('ageActiveOnlyCheck');
    const ageActiveModeSummaryText = document.getElementById('ageActiveModeSummaryText');
    const ageRefreshBtn = document.getElementById('ageRefreshBtn');
    const ageLoading = document.getElementById('ageLoading');

    // Workspace & Outputs
    const ageWorkspace = document.getElementById('ageWorkspace');
    const ageTotalBadge = document.getElementById('ageTotalBadge');
    const agePagesBadge = document.getElementById('agePagesBadge');
    const ageModeBadge = document.getElementById('ageModeBadge');
    const ageWardBadge = document.getElementById('ageWardBadge');
    const ageWorkspaceTitle = document.getElementById('ageWorkspaceTitle');
    const agePrintDirectBtn = document.getElementById('agePrintDirectBtn');
    const ageDownloadPdfBtn = document.getElementById('ageDownloadPdfBtn');
    const ageDownloadPdfBtnText = document.getElementById('ageDownloadPdfBtnText');
    const ageDownloadExcelBtn = document.getElementById('ageDownloadExcelBtn');
    const agePrevPageBtn = document.getElementById('agePrevPageBtn');
    const ageNextPageBtn = document.getElementById('ageNextPageBtn');
    const ageCurrentPageText = document.getElementById('ageCurrentPageText');
    const ageTotalPagesText = document.getElementById('ageTotalPagesText');
    const ageWardInfoBadge = document.getElementById('ageWardInfoBadge');
    const ageWardStatsText = document.getElementById('ageWardStatsText');
    const ageZoomToggleBtn = document.getElementById('ageZoomToggleBtn');
    const ageZoomBtnText = document.getElementById('ageZoomBtnText');
    const ageA4Paper = document.getElementById('ageA4Paper');

    // Helper: Set Age Loading
    function setAgeLoading(isLoading, text = '') {
        if (!ageLoading) return;
        if (isLoading) {
            ageLoading.classList.remove('hidden');
            ageLoading.classList.add('flex');
            if (text && ageLoading.querySelector('p')) {
                ageLoading.querySelector('p').textContent = text;
            }
            if (ageRefreshBtn) ageRefreshBtn.disabled = true;
        } else {
            ageLoading.classList.add('hidden');
            ageLoading.classList.remove('flex');
            if (ageRefreshBtn) ageRefreshBtn.disabled = false;
        }
    }

    // Drag & Drop for Feature 6
    if (ageDropZone) {
        ageDropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            ageDropZone.classList.add('border-teal-500', 'bg-teal-50/40');
        });

        ageDropZone.addEventListener('dragleave', () => {
            ageDropZone.classList.remove('border-teal-500', 'bg-teal-50/40');
        });

        ageDropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            ageDropZone.classList.remove('border-teal-500', 'bg-teal-50/40');
            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                handleAgeFileChosen(e.dataTransfer.files[0]);
            }
        });

        ageFileInput?.addEventListener('change', () => {
            if (ageFileInput.files && ageFileInput.files.length > 0) {
                handleAgeFileChosen(ageFileInput.files[0]);
            }
        });
    }

    function handleAgeFileChosen(file) {
        const name = file.name.toLowerCase();
        if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
            alert('कृपया केवल वैध .xlsx या .xls एक्सेल फ़ाइल चुनें');
            return;
        }

        ageState.selectedFile = file;

        if (ageSelectedFileName) ageSelectedFileName.textContent = file.name;
        if (ageSelectedFileSize) ageSelectedFileSize.textContent = `${(file.size / 1024).toFixed(1)} KB`;

        if (ageFileSelectedBox) {
            ageFileSelectedBox.classList.remove('hidden');
            ageFileSelectedBox.classList.add('flex');
        }

        if (ageStartProcessBtnText) {
            ageStartProcessBtnText.textContent = `🚀 "${file.name}" से आयु-वार लिस्ट प्रोसेस करें`;
        }
    }

    // Start Process Button Handler
    if (ageStartProcessBtn) {
        ageStartProcessBtn.addEventListener('click', () => {
            if (ageState.selectedFile) {
                executeUploadAgeExcel(ageState.selectedFile);
            } else if (ageState.sessionId) {
                fetchAgePreview();
            } else {
                alert('कृपया पहले कोई एक्सेल फ़ाइल चुनें या वर्तमान सत्र से डेटा लोड करें');
            }
        });
    }

    // Upload Excel Execution
    async function executeUploadAgeExcel(file) {
        setAgeLoading(true, `एक्सेल फ़ाइल '${file.name}' पार्स की जा रही है...`);
        if (ageStartProcessBtn) ageStartProcessBtn.disabled = true;

        try {
            const formData = new FormData();
            formData.append('excel_file', file);

            const response = await fetch('/api/age_list/upload_excel', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                let errMsg = 'एक्सेल फ़ाइल लोड करने में विफल';
                try {
                    const err = await response.json();
                    errMsg = err.error || errMsg;
                } catch (_) {
                    errMsg = `सर्वर त्रुटि (${response.status}: ${response.statusText || 'त्रुटि'})`;
                }
                throw new Error(errMsg);
            }

            const data = await response.json();
            applyAgeData(data);
        } catch (err) {
            alert('त्रुटि: ' + err.message);
        } finally {
            setAgeLoading(false);
            if (ageStartProcessBtn) ageStartProcessBtn.disabled = false;
        }
    }

    // Load from Session Button Handler
    if (ageLoadSessionBtn) {
        ageLoadSessionBtn.addEventListener('click', () => {
            executeLoadAgeFromSession();
        });
    }

    async function executeLoadAgeFromSession() {
        setAgeLoading(true, 'वर्तमान सत्र से मतदाता डेटा लोड किया जा रहा है...');
        if (ageStartProcessBtn) ageStartProcessBtn.disabled = true;

        try {
            const response = await fetch('/api/age_list/from_session', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ session_id: state.sessionId })
            });

            if (!response.ok) {
                let errMsg = 'सत्र डेटा प्राप्त नहीं हो सका';
                try {
                    const err = await response.json();
                    errMsg = err.error || errMsg;
                } catch (_) {
                    errMsg = `सर्वर त्रुटि (${response.status}: ${response.statusText || 'त्रुटि'})`;
                }
                throw new Error(errMsg);
            }

            const data = await response.json();
            applyAgeData(data);
        } catch (err) {
            alert('सूचना: ' + err.message);
        } finally {
            setAgeLoading(false);
            if (ageStartProcessBtn) ageStartProcessBtn.disabled = false;
        }
    }

    // Apply Parsed Data to Age State & UI
    function applyAgeData(data) {
        ageState.sessionId = data.session_id;
        ageState.sheetNames = data.sheet_names || [];
        ageState.activeSheet = data.active_sheet || '';
        ageState.wards = data.wards || [];
        ageState.totalVoters = data.total_voters || 0;
        ageState.currentPage = 1;
        ageState.jilaParishad = data.default_jila_parishad || '';
        ageState.panchayatSamiti = data.default_panchayat_samiti || '';

        if (ageJilaParishadInput) ageJilaParishadInput.value = ageState.jilaParishad;
        if (agePanchayatSamitiInput) agePanchayatSamitiInput.value = ageState.panchayatSamiti;

        if (agePanchayatInput && data.default_panchayat_name) {
            agePanchayatInput.value = data.default_panchayat_name;
        }
        if (ageBoothInput && data.default_booth_address) {
            ageBoothInput.value = data.default_booth_address;
        }

        // Sheet Selector setup
        if (ageSheetSelect) {
            ageSheetSelect.innerHTML = '';
            (data.sheet_names || ['Sheet1']).forEach(s => {
                const opt = document.createElement('option');
                opt.value = s;
                opt.textContent = s;
                if (s === data.active_sheet) opt.selected = true;
                ageSheetSelect.appendChild(opt);
            });
        }

        // Ward Selector setup
        if (ageWardSelect) {
            ageWardSelect.innerHTML = `
                <option value="all" selected>समस्त वार्ड (पूरी पंचायत - आयु अनुसार)</option>
                <option value="all_wardwise">समस्त वार्ड (वार्ड अनुसार क्रमिक 1, 2, 3...) - Ward-Wise Sequence</option>
            `;
            const wardsList = (data.wards && data.wards.length > 0) ? data.wards : [];
            if (wardsList.length === 0 && data.voters_sample) {
                const wardsFound = new Set();
                (data.voters_sample || []).forEach(v => {
                    const p = v.part_no || v.ward;
                    if (p) wardsFound.add(String(p).trim());
                });
                wardsFound.forEach(w => wardsList.push(w));
            }
            wardsList.forEach(w => {
                const opt = document.createElement('option');
                opt.value = w;
                opt.textContent = `वार्ड नं.- ${w}`;
                ageWardSelect.appendChild(opt);
            });
        }

        fetchAgePreview();
        refreshHistoryPills();
    }

    // Mode Switcher Function
    function setAgeMode(mode) {
        ageState.listType = mode;
        ageState.currentPage = 1;

        if (mode === 'young') {
            ageState.minAge = 18;
            ageState.maxAge = 26;
            ageState.sortOrder = 'asc';

            if (ageModeYoungBtn) {
                ageModeYoungBtn.className = 'flex items-center justify-between p-3.5 rounded-xl border-2 border-teal-600 bg-teal-50/80 text-teal-950 font-bold text-xs shadow-xs transition cursor-pointer hover:bg-teal-100/80 text-left';
                const sp = ageModeYoungBtn.querySelector('span:last-child');
                if (sp) { sp.className = 'bg-teal-600 text-white text-[10px] font-extrabold px-2 py-0.5 rounded-full shrink-0'; sp.textContent = 'सक्रिय'; }
            }
            if (ageModeSeniorBtn) {
                ageModeSeniorBtn.className = 'flex items-center justify-between p-3.5 rounded-xl border-2 border-slate-200 bg-white text-slate-700 font-bold text-xs shadow-2xs transition cursor-pointer hover:bg-slate-100 text-left';
                const sp = ageModeSeniorBtn.querySelector('span:last-child');
                if (sp) { sp.className = 'bg-slate-200 text-slate-600 text-[10px] font-extrabold px-2 py-0.5 rounded-full shrink-0'; sp.textContent = 'चुनें'; }
            }

            if (ageListTitleInput) ageListTitleInput.value = 'युवा मतदाता सूची (आयु 18 से 26 वर्ष)';
            if (ageActiveModeSummaryText) ageActiveModeSummaryText.innerHTML = 'वर्तमान मोड: <strong>युवा (18 - 26 वर्ष • बढ़ते क्रम में 18 &rarr; 26)</strong>';
            if (ageModeBadge) ageModeBadge.textContent = 'युवा (18-26 वर्ष)';
            if (ageWorkspaceTitle) ageWorkspaceTitle.textContent = 'युवा मतदाता सूची (A4 शीट प्रिव्यू)';

        } else if (mode === 'senior') {
            ageState.minAge = 70;
            ageState.maxAge = 120;
            ageState.sortOrder = 'desc';

            if (ageModeSeniorBtn) {
                ageModeSeniorBtn.className = 'flex items-center justify-between p-3.5 rounded-xl border-2 border-amber-600 bg-amber-50/80 text-amber-950 font-bold text-xs shadow-xs transition cursor-pointer hover:bg-amber-100/80 text-left';
                const sp = ageModeSeniorBtn.querySelector('span:last-child');
                if (sp) { sp.className = 'bg-amber-600 text-white text-[10px] font-extrabold px-2 py-0.5 rounded-full shrink-0'; sp.textContent = 'सक्रिय'; }
            }
            if (ageModeYoungBtn) {
                ageModeYoungBtn.className = 'flex items-center justify-between p-3.5 rounded-xl border-2 border-slate-200 bg-white text-slate-700 font-bold text-xs shadow-2xs transition cursor-pointer hover:bg-slate-100 text-left';
                const sp = ageModeYoungBtn.querySelector('span:last-child');
                if (sp) { sp.className = 'bg-slate-200 text-slate-600 text-[10px] font-extrabold px-2 py-0.5 rounded-full shrink-0'; sp.textContent = 'चुनें'; }
            }

            if (ageListTitleInput) ageListTitleInput.value = 'वरिष्ठ / बुजुर्ग मतदाता सूची (आयु 120 से 70 वर्ष)';
            if (ageActiveModeSummaryText) ageActiveModeSummaryText.innerHTML = 'वर्तमान मोड: <strong>बुजुर्ग / वरिष्ठ (120 - 70 वर्ष • घटते क्रम में 120 &rarr; 70)</strong>';
            if (ageModeBadge) ageModeBadge.textContent = 'बुजुर्ग (120-70 वर्ष)';
            if (ageWorkspaceTitle) ageWorkspaceTitle.textContent = 'वरिष्ठ / बुजुर्ग मतदाता सूची (A4 शीट प्रिव्यू)';
        }

        if (window.lucide) lucide.createIcons();
        fetchAgePreview();
    }

    if (ageModeYoungBtn) {
        ageModeYoungBtn.addEventListener('click', () => setAgeMode('young'));
    }
    if (ageModeSeniorBtn) {
        ageModeSeniorBtn.addEventListener('click', () => setAgeMode('senior'));
    }

    // Sheet Selection Change
    if (ageSheetSelect) {
        ageSheetSelect.addEventListener('change', async () => {
            const sheetName = ageSheetSelect.value;
            if (!sheetName || !ageState.sessionId) return;

            setAgeLoading(true, `शीट '${sheetName}' लोड हो रही है...`);
            try {
                const response = await fetch('/api/age_list/select_sheet', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        session_id: ageState.sessionId,
                        sheet_name: sheetName
                    })
                });

                if (!response.ok) throw new Error('शीट लोड करने में विफल');
                const data = await response.json();
                applyAgeData(data);
            } catch (err) {
                alert(err.message);
            } finally {
                setAgeLoading(false);
            }
        });
    }

    // Ward Selection Change
    if (ageWardSelect) {
        ageWardSelect.addEventListener('change', () => {
            const sel = ageWardSelect.value;
            ageState.selectedWard = sel;
            ageState.currentPage = 1;
            fetchAgePreview();
        });
    }

    // Checkbox Changes
    if (ageActiveOnlyCheck) {
        ageActiveOnlyCheck.addEventListener('change', () => {
            ageState.currentPage = 1;
            fetchAgePreview();
        });
    }

    // Fetch Live Age Preview HTML
    async function fetchAgePreview() {
        if (!ageState.sessionId) return;

        setAgeLoading(true, 'आयु-वार मतदाता सूची पूर्वावलोकन तैयार हो रहा है...');
        try {
            const wardTitle = agePanchayatInput ? agePanchayatInput.value.trim() : 'ग्राम पंचायत';
            const listTitle = ageListTitleInput ? ageListTitleInput.value.trim() : '';
            const boothAddr = ageBoothInput ? ageBoothInput.value.trim() : '';
            const selWard = ageWardSelect ? ageWardSelect.value : 'all';
            const filterActive = ageActiveOnlyCheck ? ageActiveOnlyCheck.checked : true;
            const jp = ageJilaParishadInput ? ageJilaParishadInput.value.trim() : (ageState.jilaParishad || '');
            const ps = agePanchayatSamitiInput ? agePanchayatSamitiInput.value.trim() : (ageState.panchayatSamiti || '');

            const payload = {
                session_id: ageState.sessionId,
                ward_title: wardTitle,
                list_title: listTitle,
                booth_address: boothAddr,
                list_type: ageState.listType,
                min_age: ageState.minAge,
                max_age: ageState.maxAge,
                sort_order: ageState.sortOrder,
                selected_ward: selWard,
                filter_active_only: filterActive,
                jila_parishad: jp,
                panchayat_samiti: ps,
                page: ageState.currentPage
            };

            const response = await fetch('/api/age_list/preview_html', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            if (!response.ok) {
                let errMsg = 'पूर्वावलोकन प्राप्त नहीं हो सका';
                try {
                    const err = await response.json();
                    errMsg = err.error || errMsg;
                } catch (_) {
                    errMsg = `सर्वर त्रुटि (${response.status}: ${response.statusText || 'त्रुटि'})`;
                }
                throw new Error(errMsg);
            }

            const res = await response.json();
            ageState.totalPages = res.total_pages || 1;
            ageState.currentPage = res.current_page || 1;

            if (ageA4Paper) {
                const styleMatch = res.html.match(/<style[^>]*>([\s\S]*?)<\/style>/i);
                const bodyMatch = res.html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
                const styles = styleMatch ? `<style>${styleMatch[1]}</style>` : '';
                const bodyContent = bodyMatch ? bodyMatch[1] : res.html;
                ageA4Paper.innerHTML = styles + bodyContent;
            }

            if (ageWorkspace) ageWorkspace.classList.remove('hidden');
            if (ageTotalBadge) ageTotalBadge.textContent = `${res.total_matched || 0} मतदाता`;
            if (agePagesBadge) agePagesBadge.textContent = `${res.total_pages} A4 पेजेस`;
            if (ageCurrentPageText) ageCurrentPageText.textContent = ageState.currentPage;
            if (ageTotalPagesText) ageTotalPagesText.textContent = ageState.totalPages;

            if (ageWardBadge) {
                if (selWard === 'all') ageWardBadge.textContent = 'समस्त वार्ड (आयु अनुसार)';
                else if (selWard === 'all_wardwise') ageWardBadge.textContent = 'समस्त वार्ड (क्रमिक वार्ड-वार)';
                else ageWardBadge.textContent = `वार्ड नं.- ${res.current_ward || selWard}`;
            }

            if (ageWardInfoBadge) {
                ageWardInfoBadge.textContent = (res.current_ward && res.current_ward !== 'समस्त') ? `वार्ड ${res.current_ward}` : 'समस्त वार्ड';
            }
            if (ageWardStatsText) {
                ageWardStatsText.innerHTML = `[पेज ${res.page_in_ward || 1} / ${res.total_in_ward || 1} • कुल ${res.ward_voters || res.total_matched || 0} मतदाता]`;
            }

            if (agePrevPageBtn) agePrevPageBtn.disabled = (ageState.currentPage <= 1);
            if (ageNextPageBtn) ageNextPageBtn.disabled = (ageState.currentPage >= ageState.totalPages);

            if (window.lucide) lucide.createIcons();
        } catch (err) {
            alert('आयु सूची पूर्वावलोकन त्रुटि: ' + err.message);
        } finally {
            setAgeLoading(false);
        }
    }

    if (ageRefreshBtn) {
        ageRefreshBtn.addEventListener('click', () => {
            fetchAgePreview();
        });
    }

    // Pagination Buttons
    if (agePrevPageBtn) {
        agePrevPageBtn.addEventListener('click', () => {
            if (ageState.currentPage > 1) {
                ageState.currentPage--;
                fetchAgePreview();
            }
        });
    }

    if (ageNextPageBtn) {
        ageNextPageBtn.addEventListener('click', () => {
            if (ageState.currentPage < ageState.totalPages) {
                ageState.currentPage++;
                fetchAgePreview();
            }
        });
    }

    // Zoom Toggle
    if (ageZoomToggleBtn && ageA4Paper) {
        ageZoomToggleBtn.addEventListener('click', () => {
            ageState.isZoomed = !ageState.isZoomed;
            if (ageState.isZoomed) {
                ageA4Paper.classList.remove('a4-scaled');
                ageZoomBtnText.textContent = 'फ़िट टू स्क्रीन';
            } else {
                ageA4Paper.classList.add('a4-scaled');
                ageZoomBtnText.textContent = '100% ज़ूम';
            }
        });
    }

    // Direct Browser Print Action
    if (agePrintDirectBtn) {
        agePrintDirectBtn.addEventListener('click', async () => {
            if (!ageState.sessionId) {
                alert('कृपया पहले कोई एक्सेल फ़ाइल अपलोड करें या सत्र लोड करें');
                return;
            }

            const wardTitle = encodeURIComponent(agePanchayatInput ? agePanchayatInput.value.trim() : 'ग्राम पंचायत');
            const listTitle = encodeURIComponent(ageListTitleInput ? ageListTitleInput.value.trim() : '');
            const boothAddr = encodeURIComponent(ageBoothInput ? ageBoothInput.value.trim() : '');
            const selWard = encodeURIComponent(ageWardSelect ? ageWardSelect.value : 'all');
            const filterActive = ageActiveOnlyCheck ? ageActiveOnlyCheck.checked : true;
            const jp = encodeURIComponent((ageJilaParishadInput ? ageJilaParishadInput.value.trim() : ageState.jilaParishad) || '');
            const ps = encodeURIComponent((agePanchayatSamitiInput ? agePanchayatSamitiInput.value.trim() : ageState.panchayatSamiti) || '');

            const printUrl = `/age_list/print?session_id=${ageState.sessionId}&ward_title=${wardTitle}&list_title=${listTitle}&booth_address=${boothAddr}&list_type=${ageState.listType}&min_age=${ageState.minAge}&max_age=${ageState.maxAge}&sort_order=${ageState.sortOrder}&selected_ward=${selWard}&filter_active_only=${filterActive}&jila_parishad=${jp}&panchayat_samiti=${ps}`;
            const printWin = window.open(printUrl, '_blank');
            if (!printWin) {
                alert('कृपया अपने ब्राउज़र में पॉप-अप (Pop-ups) की अनुमति दें ताकि प्रिंट डायलॉग खुल सके।');
            }
        });
    }

    // Vector PDF Download Action
    if (ageDownloadPdfBtn) {
        ageDownloadPdfBtn.addEventListener('click', async () => {
            if (!ageState.sessionId) {
                alert('कृपया पहले कोई एक्सेल फ़ाइल अपलोड करें या सत्र लोड करें');
                return;
            }

            ageDownloadPdfBtn.disabled = true;
            const originalHtml = ageDownloadPdfBtn.innerHTML;
            ageDownloadPdfBtn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i> <span>PDF जनरेट हो रही है...</span>';
            if (window.lucide) lucide.createIcons();

            try {
                const wardTitle = agePanchayatInput ? agePanchayatInput.value.trim() : 'ग्राम पंचायत';
                const listTitle = ageListTitleInput ? ageListTitleInput.value.trim() : '';
                const boothAddr = ageBoothInput ? ageBoothInput.value.trim() : '';
                const selWard = ageWardSelect ? ageWardSelect.value : 'all';
                const filterActive = ageActiveOnlyCheck ? ageActiveOnlyCheck.checked : true;
                const jp = ageJilaParishadInput ? ageJilaParishadInput.value.trim() : (ageState.jilaParishad || '');
                const ps = agePanchayatSamitiInput ? agePanchayatSamitiInput.value.trim() : (ageState.panchayatSamiti || '');

                const payload = {
                    session_id: ageState.sessionId,
                    ward_title: wardTitle,
                    list_title: listTitle,
                    booth_address: boothAddr,
                    list_type: ageState.listType,
                    min_age: ageState.minAge,
                    max_age: ageState.maxAge,
                    sort_order: ageState.sortOrder,
                    selected_ward: selWard,
                    filter_active_only: filterActive,
                    jila_parishad: jp,
                    panchayat_samiti: ps
                };

                const response = await fetch('/api/age_list/export_pdf', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.error || 'PDF डाउनलोड विफल');
                }

                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                const safeName = (wardTitle || 'Age_List').replace(/[^\w\s-]/g, '').trim() || 'Age_List';
                const typeStr = ageState.listType === 'young' ? 'Young_18-26' : 'Senior_120-70';
                a.download = `Geam_Digital_${safeName}_${typeStr}_List.pdf`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            } catch (err) {
                alert('PDF निर्यात त्रुटि: ' + err.message);
            } finally {
                ageDownloadPdfBtn.disabled = false;
                ageDownloadPdfBtn.innerHTML = originalHtml;
                if (window.lucide) lucide.createIcons();
            }
        });
    }

    // Sorted Excel Download Action
    if (ageDownloadExcelBtn) {
        ageDownloadExcelBtn.addEventListener('click', async () => {
            if (!ageState.sessionId) {
                alert('कृपया पहले कोई एक्सेल फ़ाइल अपलोड करें या सत्र लोड करें');
                return;
            }

            ageDownloadExcelBtn.disabled = true;
            const originalHtml = ageDownloadExcelBtn.innerHTML;
            ageDownloadExcelBtn.innerHTML = '<i data-lucide="loader-2" class="w-4 h-4 animate-spin"></i> <span>Excel तैयार हो रहा है...</span>';
            if (window.lucide) lucide.createIcons();

            try {
                const wardTitle = agePanchayatInput ? agePanchayatInput.value.trim() : 'ग्राम पंचायत';
                const listTitle = ageListTitleInput ? ageListTitleInput.value.trim() : '';
                const boothAddr = ageBoothInput ? ageBoothInput.value.trim() : '';
                const selWard = ageWardSelect ? ageWardSelect.value : 'all';
                const filterActive = ageActiveOnlyCheck ? ageActiveOnlyCheck.checked : true;
                const jp = ageJilaParishadInput ? ageJilaParishadInput.value.trim() : (ageState.jilaParishad || '');
                const ps = agePanchayatSamitiInput ? agePanchayatSamitiInput.value.trim() : (ageState.panchayatSamiti || '');

                const payload = {
                    session_id: ageState.sessionId,
                    ward_title: wardTitle,
                    list_title: listTitle,
                    booth_address: boothAddr,
                    list_type: ageState.listType,
                    min_age: ageState.minAge,
                    max_age: ageState.maxAge,
                    sort_order: ageState.sortOrder,
                    selected_ward: selWard,
                    filter_active_only: filterActive,
                    jila_parishad: jp,
                    panchayat_samiti: ps
                };

                const response = await fetch('/api/age_list/export_excel', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                if (!response.ok) {
                    const err = await response.json();
                    throw new Error(err.error || 'Excel डाउनलोड विफल');
                }

                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                const safeName = (wardTitle || 'Age_List').replace(/[^\w\s-]/g, '').trim() || 'Age_List';
                const typeStr = ageState.listType === 'young' ? 'Young_18-26' : 'Senior_120-70';
                a.download = `Geam_Digital_${safeName}_${typeStr}_List.xlsx`;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
            } catch (err) {
                alert('Excel निर्यात त्रुटि: ' + err.message);
            } finally {
                ageDownloadExcelBtn.disabled = false;
                ageDownloadExcelBtn.innerHTML = originalHtml;
                refreshHistoryPills();
                if (window.lucide) lucide.createIcons();
            }
        });
    }

    // ==============================================================
    // CROSS-FEATURE ACTIVITY & FILE HISTORY SYSTEM (DRAWER & PILLS)
    // ==============================================================

    let currentHistoryFilter = 'all';

    function openHistoryDrawer(featureFilter = 'all') {
        currentHistoryFilter = featureFilter;
        updateHistoryFilterButtons(featureFilter);

        if (historyDrawerBackdrop) {
            historyDrawerBackdrop.classList.remove('hidden');
            setTimeout(() => {
                if (historyDrawerPanel) historyDrawerPanel.classList.remove('translate-x-full');
            }, 10);
        }
        loadHistoryList(featureFilter);
    }

    function closeHistoryDrawer() {
        if (historyDrawerPanel) {
            historyDrawerPanel.classList.add('translate-x-full');
            setTimeout(() => {
                if (historyDrawerBackdrop) historyDrawerBackdrop.classList.add('hidden');
            }, 300);
        }
    }

    function updateHistoryFilterButtons(filter) {
        historyFilterBtns.forEach(btn => {
            const btnFilter = btn.getAttribute('data-filter');
            if (btnFilter === filter) {
                btn.className = 'history-filter-btn px-3 py-1.5 rounded-lg bg-white text-slate-900 border border-slate-200 shadow-2xs font-extrabold cursor-pointer';
            } else {
                btn.className = 'history-filter-btn px-3 py-1.5 rounded-lg text-slate-600 hover:text-slate-900 hover:bg-white/80 font-bold cursor-pointer';
            }
        });
    }

    async function loadHistoryList(filter = 'all') {
        if (!historyListContainer) return;

        historyListContainer.innerHTML = `
            <div class="text-center py-8 text-slate-400">
                <div class="inline-block animate-spin rounded-full h-6 w-6 border-2 border-slate-300 border-t-sky-600"></div>
                <p class="text-xs mt-2 font-medium">इतिहास लोड हो रहा है...</p>
            </div>
        `;

        try {
            const url = filter && filter !== 'all' ? `/api/history?feature_id=${encodeURIComponent(filter)}` : '/api/history';
            const resp = await fetch(url);
            const data = await resp.json();

            if (!resp.ok) throw new Error(data.error || 'इतिहास लोड नहीं हो सका');

            const records = data.history || [];
            if (records.length === 0) {
                historyListContainer.innerHTML = `
                    <div class="text-center py-12 px-4 space-y-3">
                        <div class="w-12 h-12 mx-auto rounded-full bg-slate-100 flex items-center justify-center text-slate-400">
                            <i data-lucide="inbox" class="w-6 h-6"></i>
                        </div>
                        <p class="text-xs font-bold text-slate-700">कोई इतिहास रिकॉर्ड नहीं मिला</p>
                        <p class="text-[11px] text-slate-400">जब आप पीडीएफ/एक्सेल अपलोड या संपादित करेंगे, तो तारीख, समय व फ़ाइल का नाम यहाँ सुरक्षित दिखेगा।</p>
                    </div>
                `;
                if (window.lucide) lucide.createIcons();
                return;
            }

            historyListContainer.innerHTML = '';
            records.forEach(item => {
                const card = document.createElement('div');
                card.className = 'bg-white border border-slate-200 hover:border-slate-300 rounded-xl p-3.5 shadow-2xs space-y-2 transition-all';

                // Color accent based on feature
                let featureBadgeColor = 'bg-slate-100 text-slate-700 border-slate-200';
                if (item.feature_id === 'feature1') featureBadgeColor = 'bg-emerald-50 text-emerald-800 border-emerald-200';
                else if (item.feature_id === 'feature2') featureBadgeColor = 'bg-indigo-50 text-indigo-800 border-indigo-200';
                else if (item.feature_id === 'feature3') featureBadgeColor = 'bg-rose-50 text-rose-800 border-rose-200';
                else if (item.feature_id === 'feature4') featureBadgeColor = 'bg-amber-50 text-amber-800 border-amber-200';
                else if (item.feature_id === 'feature5') featureBadgeColor = 'bg-purple-50 text-purple-800 border-purple-200';
                else if (item.feature_id === 'feature6') featureBadgeColor = 'bg-teal-50 text-teal-800 border-teal-200';

                const fileNamesList = Array.isArray(item.file_names) ? item.file_names : [item.file_names || 'अज्ञात फ़ाइल'];

                card.innerHTML = `
                    <div class="flex items-center justify-between gap-2">
                        <span class="text-[10px] font-extrabold px-2 py-0.5 rounded-md border ${featureBadgeColor}">
                            ${item.feature_name || 'सिस्टम गतिविधि'}
                        </span>
                        <div class="flex items-center space-x-1.5">
                            <span class="text-[10px] font-semibold text-slate-500 flex items-center gap-1">
                                <i data-lucide="clock" class="w-3 h-3 text-slate-400"></i>
                                <span>${item.display_date || ''} • <strong>${item.display_time || ''}</strong></span>
                            </span>
                            <button type="button" class="delete-history-btn p-1 text-slate-400 hover:text-rose-500 rounded hover:bg-slate-100 transition cursor-pointer" data-id="${item.id}" title="यह प्रविष्टि हटाएं">
                                <i data-lucide="trash" class="w-3.5 h-3.5"></i>
                            </button>
                        </div>
                    </div>
                    <div class="space-y-1">
                        <div class="text-xs font-bold text-slate-900 flex items-start gap-1.5">
                            <i data-lucide="file-text" class="w-3.5 h-3.5 text-slate-500 shrink-0 mt-0.5"></i>
                            <span class="break-all leading-snug">${fileNamesList.join(', ')}</span>
                        </div>
                        <div class="text-[11px] text-slate-600 flex items-center gap-1.5 flex-wrap">
                            <span class="font-medium text-slate-400">क्रिया:</span>
                            <span class="font-bold text-slate-800">${item.action || 'अपलोड / संपादन'}</span>
                            ${item.details ? `<span class="text-slate-300">•</span><span class="text-slate-500">${item.details}</span>` : ''}
                        </div>
                    </div>
                `;

                // Single record delete
                const delBtn = card.querySelector('.delete-history-btn');
                if (delBtn) {
                    delBtn.addEventListener('click', async (e) => {
                        e.stopPropagation();
                        await deleteHistoryRecord(item.id);
                    });
                }

                historyListContainer.appendChild(card);
            });

            if (window.lucide) lucide.createIcons();

        } catch (err) {
            historyListContainer.innerHTML = `
                <div class="text-center py-6 text-rose-500 text-xs font-bold">
                    इतिहास प्राप्त करने में विफल: ${err.message}
                </div>
            `;
        }
    }

    async function deleteHistoryRecord(id) {
        try {
            const resp = await fetch('/api/history/delete', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ id })
            });
            if (resp.ok) {
                loadHistoryList(currentHistoryFilter);
                refreshHistoryPills();
            }
        } catch (e) {
            console.warn('Failed to delete history record', e);
        }
    }

    async function clearAllHistory() {
        if (!confirm('क्या आप वाकई सम्पूर्ण सिस्टम इतिहास साफ़ करना चाहते हैं?')) return;
        try {
            const resp = await fetch('/api/history/clear', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' }
            });
            if (resp.ok) {
                loadHistoryList(currentHistoryFilter);
                refreshHistoryPills();
            }
        } catch (e) {
            alert('इतिहास साफ़ करने में त्रुटि: ' + e.message);
        }
    }

    async function refreshHistoryPills() {
        try {
            const resp = await fetch('/api/history?limit=150');
            if (!resp.ok) return;
            const data = await resp.json();
            const records = data.history || [];

            // Update top header badge
            if (historyCountBadge) {
                historyCountBadge.textContent = records.length;
            }

            // Update feature pills
            const features = [
                { id: 'feature1', pill: 'feature1HistoryPill', fileEl: 'feature1HistFile', timeEl: 'feature1HistTime' },
                { id: 'feature2', pill: 'feature2HistoryPill', fileEl: 'feature2HistFile', timeEl: 'feature2HistTime' },
                { id: 'feature3', pill: 'feature3HistoryPill', fileEl: 'feature3HistFile', timeEl: 'feature3HistTime' },
                { id: 'feature4', pill: 'feature4HistoryPill', fileEl: 'feature4HistFile', timeEl: 'feature4HistTime' },
                { id: 'feature5', pill: 'feature5HistoryPill', fileEl: 'feature5HistFile', timeEl: 'feature5HistTime' },
                { id: 'feature6', pill: 'feature6HistoryPill', fileEl: 'feature6HistFile', timeEl: 'feature6HistTime' }
            ];

            features.forEach(f => {
                const rec = records.find(r => r.feature_id === f.id);
                const pillEl = document.getElementById(f.pill);
                const fileEl = document.getElementById(f.fileEl);
                const timeEl = document.getElementById(f.timeEl);

                if (rec && pillEl && fileEl && timeEl) {
                    const fn = Array.isArray(rec.file_names) ? rec.file_names[0] : (rec.file_names || '-');
                    fileEl.textContent = fn;
                    timeEl.textContent = `${rec.display_date} • ${rec.display_time}`;
                    pillEl.classList.remove('hidden');
                    pillEl.classList.add('flex');
                } else if (pillEl) {
                    pillEl.classList.add('hidden');
                    pillEl.classList.remove('flex');
                }
            });

            if (window.lucide) lucide.createIcons();
        } catch (e) {
            console.warn('Error refreshing history pills:', e);
        }
    }

    // Expose openHistoryDrawer globally for HTML onclick handlers
    window.openHistoryDrawer = openHistoryDrawer;

    // Attach Event Listeners
    if (historyModalBtn) {
        historyModalBtn.addEventListener('click', () => openHistoryDrawer('all'));
    }
    if (closeHistoryDrawerBtn) {
        closeHistoryDrawerBtn.addEventListener('click', closeHistoryDrawer);
    }
    if (clearHistoryBtn) {
        clearHistoryBtn.addEventListener('click', clearAllHistory);
    }
    if (historyDrawerBackdrop) {
        historyDrawerBackdrop.addEventListener('click', (e) => {
            if (e.target === historyDrawerBackdrop) closeHistoryDrawer();
        });
    }
    historyFilterBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            const filter = btn.getAttribute('data-filter') || 'all';
            openHistoryDrawer(filter);
        });
    });
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && historyDrawerBackdrop && !historyDrawerBackdrop.classList.contains('hidden')) {
            closeHistoryDrawer();
        }
    });

    // Initial load of history pills
    refreshHistoryPills();

});
