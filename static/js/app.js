// Geam Digital - Multi-Ward Panchayat Voter List to Excel Client Application Logic (with Status Column)

document.addEventListener('DOMContentLoaded', () => {
    // Application State
    const state = {
        sessionId: null,
        fileName: 'Geam_Digital_Panchayat_Voter_List.xlsx',
        summaryGrid: [],     // Sheet 1: वार्ड_सारांश
        masterGrid: [],      // Sheet 2: समस्त_मतदाता_सूची (All Wards Sequence-wise with 11 columns)
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
            convertBtnText.textContent = `अपलोड करें और "${pdfFiles[0].name}" की 11-कॉलम एक्सेल पाएं`;
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

    // 3. Form Submit - Upload & Process Multi-Ward PDFs
    uploadForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        if (!fileInput.files || fileInput.files.length === 0) {
            alert('कृपया पहले एक या अधिक वोटर लिस्ट पीडीएफ फ़ाइलें चुनें!');
            return;
        }

        const formData = new FormData();
        for (let i = 0; i < fileInput.files.length; i++) {
            formData.append('pdf_files', fileInput.files[i]);
        }

        const count = fileInput.files.length;
        setLoading(true, `सभी ${count} वार्डों की पीडीएफ का विश्लेषण किया जा रहा है... (वार्ड सॉर्टिंग, स्टेटस Active/Deleted पहचान, OCR)`);

        try {
            const response = await fetch('/api/upload', {
                method: 'POST',
                body: formData
            });
            const resData = await response.json();
            if (!response.ok) throw new Error(resData.error || 'Upload and conversion failed');

            handleBatchResponse(resData);
        } catch (err) {
            alert(`त्रुटि: ${err.message}`);
        } finally {
            setLoading(false);
        }
    });

    // 4. Handle Server Response for Multi-Ward Batch
    function handleBatchResponse(resData) {
        state.sessionId = resData.session_id;
        state.fileName = resData.file_name || 'Geam_Digital_Panchayat_Voter_List.xlsx';
        const data = resData.data;

        state.summaryGrid = data.summary_grid || [];
        state.masterGrid = data.master_grid || [];
        state.wards = data.wards || [];
        state.allTables = data.tables || [state.summaryGrid, state.masterGrid];
        state.allTableNames = data.table_names || ['वार्ड_सारांश', 'समस्त_मतदाता_सूची'];
        state.metadata = data.invoice_metadata || {};

        // Default view: Master Sequence Sheet with 11 columns
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
        quickDownloadTitle.textContent = `${totalWards} वार्डों की सम्पूर्ण पंचायत एक्सेल तैयार है! (${totalScanned} रिकॉर्ड्स, 11th Column Status)`;
        summaryWardsCountBadge.textContent = `${totalWards} वार्ड्स सम्मिलित (कुल ${totalScanned}: ${totalActive} Active, ${totalDeleted} Deleted)`;

        // Auto Download if checked
        if (autoDownloadCheck && autoDownloadCheck.checked) {
            setTimeout(() => {
                triggerExcelDownload();
            }, 600);
        }

        // Scroll to results
        resultsWorkspace.scrollIntoView({ behavior: 'smooth' });
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
                if (['क्र.सं.', 'वार्ड / भाग संख्या', 'कुल मतदाता', 'सक्रिय (Active)', 'विलोपित (Deleted)'].includes(hName)) {
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

            if (['भाग संख्या', 'क्रम संख्या', 'आयु', 'मोबाइल नो', 'Status'].includes(h)) {
                th.classList.add('text-center');
            }
            if (h === 'मोबाइल नो') {
                th.innerHTML = `${h} <span class="text-[10px] text-amber-300 font-normal">(खाली)</span>`;
            } else if (h === 'Status') {
                th.className = 'p-2.5 text-center border border-slate-700 text-xs font-extrabold tracking-tight whitespace-nowrap bg-slate-950 text-rose-300';
                th.innerHTML = `Status <span class="text-[9px] text-slate-300 font-normal">(Col 11)</span>`;
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

        rows.forEach((row, rowIdx) => {
            const tr = document.createElement('tr');
            const statusVal = String(row[10] || '').trim();
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
                if (['भाग संख्या', 'क्रम संख्या', 'आयु', 'मोबाइल नो', 'वोटर ID', 'हाउस नंबर', 'Status'].includes(hName)) {
                    td.classList.add('text-center');
                }
                if (hName === 'वोटर ID') {
                    td.classList.add('font-mono', 'font-semibold', 'text-sky-700');
                }
                if (hName === 'नाम') {
                    td.classList.add('font-bold', 'text-slate-900');
                }
                if (hName === 'भाग संख्या') {
                    td.classList.add('font-bold', 'text-slate-700');
                }

                // Render 11th Column "Status" with dedicated Badge
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

    // 10. Cell Value Update in State
    function updateCellValue(filteredRowIdx, colIdx, newVal) {
        if (!state.filteredGrid[filteredRowIdx]) return;
        state.filteredGrid[filteredRowIdx][colIdx] = newVal;

        const row = state.filteredGrid[filteredRowIdx];
        const origIdx = state.activeGrid.indexOf(row);
        if (origIdx !== -1) {
            state.activeGrid[origIdx][colIdx] = newVal;
        }
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
});
