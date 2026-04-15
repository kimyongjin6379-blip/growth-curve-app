/**
 * Growth Curve Data Automation Tool — Frontend Logic (v1.1)
 * Updated: SAMPLE_MAP 지원, 2-step 가공 플로우
 */

(function () {
    'use strict';

    // ── DOM Elements ──
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const fileInfo = document.getElementById('file-info');
    const fileName = document.getElementById('file-name');
    const fileSize = document.getElementById('file-size');
    const fileRemove = document.getElementById('file-remove');
    const btnProcess = document.getElementById('btn-process');
    const btnDownload = document.getElementById('btn-download');
    const statusArea = document.getElementById('status-area');
    const statusLoading = document.getElementById('status-loading');
    const statusSuccess = document.getElementById('status-success');
    const statusError = document.getElementById('status-error');
    const statusMessage = document.getElementById('status-message');
    const errorMessage = document.getElementById('error-message');
    const chartPlaceholder = document.getElementById('chart-placeholder');
    const chartContainer = document.getElementById('chart-container');
    const chartBadge = document.getElementById('chart-badge');
    const summaryCard = document.getElementById('summary-card');
    const summaryGrid = document.getElementById('summary-grid');
    const sampleMapCard = document.getElementById('sample-map-card');
    const mappingTbody = document.getElementById('mapping-tbody');

    // ── State ──
    let selectedFile = null;
    let downloadFileId = null;
    let downloadFilename = null;
    let detectedGroups = [];

    // ── Helpers ──
    function formatBytes(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
    }

    // ── Color palette for chart traces ──
    const CHART_COLORS = [
        '#6366f1', '#06b6d4', '#10b981', '#f59e0b', '#ef4444',
        '#8b5cf6', '#ec4899', '#14b8a6', '#f97316', '#84cc16',
        '#a78bfa', '#22d3ee', '#34d399', '#fbbf24', '#fb7185',
    ];

    // ── File Selection ──
    function handleFileSelect(file) {
        if (!file) return;

        const ext = file.name.split('.').pop().toLowerCase();
        if (!['xlsx', 'xls', 'csv'].includes(ext)) {
            showError('지원되지 않는 파일 형식입니다. .xlsx 또는 .csv 파일을 업로드해 주세요.');
            return;
        }

        selectedFile = file;
        fileName.textContent = file.name;
        fileSize.textContent = formatBytes(file.size);
        fileInfo.style.display = 'flex';
        dropZone.classList.add('has-file');
        btnProcess.disabled = false;

        // Reset previous results
        resetResults();
    }

    function removeFile() {
        selectedFile = null;
        fileInput.value = '';
        fileInfo.style.display = 'none';
        dropZone.classList.remove('has-file');
        btnProcess.disabled = true;
        resetResults();
    }

    function resetResults() {
        statusArea.style.display = 'none';
        statusLoading.style.display = 'flex';
        statusSuccess.style.display = 'none';
        statusError.style.display = 'none';
        btnDownload.style.display = 'none';
        btnDownload.disabled = true;
        chartPlaceholder.style.display = 'flex';
        chartContainer.style.display = 'none';
        chartBadge.style.display = 'none';
        summaryCard.style.display = 'none';
        sampleMapCard.style.display = 'none';
        downloadFileId = null;
        downloadFilename = null;
        detectedGroups = [];
    }

    // ── Drag & Drop ──
    dropZone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', (e) => handleFileSelect(e.target.files[0]));
    fileRemove.addEventListener('click', (e) => {
        e.stopPropagation();
        removeFile();
    });

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        const file = e.dataTransfer.files[0];
        handleFileSelect(file);
    });

    // ── Sample Mapping UI ──
    function buildMappingTable(groups) {
        mappingTbody.innerHTML = '';
        groups.forEach((grp) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="mapping-code">${grp}</td>
                <td><input type="text" class="form-input mapping-strain" data-code="${grp}" placeholder="균주명"></td>
                <td><input type="text" class="form-input mapping-peptone1" data-code="${grp}" placeholder="펩톤명"></td>
                <td><input type="number" class="form-input mapping-ratio1" data-code="${grp}" placeholder="100" step="1" min="0" max="100" value="100"></td>
                <td><input type="text" class="form-input mapping-peptone2" data-code="${grp}" placeholder="블렌딩 시 입력"></td>
                <td><input type="number" class="form-input mapping-ratio2" data-code="${grp}" placeholder="0" step="1" min="0" max="100"></td>
                <td><input type="number" class="form-input mapping-total-pct" data-code="${grp}" placeholder="2.0" step="0.1" min="0"></td>
            `;
            // 비율 자동 계산: ratio1 변경 시 ratio2 = 100 - ratio1
            const ratio1Input = tr.querySelector('.mapping-ratio1');
            const ratio2Input = tr.querySelector('.mapping-ratio2');
            ratio1Input.addEventListener('input', () => {
                const v = parseFloat(ratio1Input.value) || 0;
                const p2 = tr.querySelector('.mapping-peptone2').value.trim();
                if (p2) ratio2Input.value = Math.max(0, 100 - v);
            });
            ratio2Input.addEventListener('input', () => {
                const v = parseFloat(ratio2Input.value) || 0;
                ratio1Input.value = Math.max(0, 100 - v);
            });
            mappingTbody.appendChild(tr);
        });
        sampleMapCard.style.display = 'block';
    }

    function collectSampleMap() {
        const rows = mappingTbody.querySelectorAll('tr');
        const result = [];

        // 초기 '사용 균주' 값 가져오기
        const globalStrainInput = document.getElementById('strain');
        let currentStrain = globalStrainInput ? globalStrainInput.value.trim() : '';
        let currentPct = 0;

        rows.forEach((tr) => {
            const code = tr.querySelector('.mapping-code').textContent.trim();

            // 균주명 처리 (값이 있으면 갱신, 없으면 이전 값 유지)
            const strainInput = tr.querySelector('.mapping-strain');
            let strain = strainInput ? strainInput.value.trim() : '';
            if (strain !== '') {
                currentStrain = strain;
            } else {
                strain = currentStrain;
                if (strainInput && strain !== '') strainInput.value = strain;
            }

            // 펩톤1 (필수)
            const peptone1 = (tr.querySelector('.mapping-peptone1').value || '').trim();
            const ratio1 = parseFloat(tr.querySelector('.mapping-ratio1').value) || 100;

            // 펩톤2 (블렌딩 시)
            const peptone2 = (tr.querySelector('.mapping-peptone2').value || '').trim();
            const ratio2 = parseFloat(tr.querySelector('.mapping-ratio2').value) || 0;

            // 총 펩톤 농도 (carry-forward: 값이 있으면 갱신, 없으면 이전 값 유지)
            const pctRaw = (tr.querySelector('.mapping-total-pct').value || '').trim();
            if (pctRaw !== '') {
                currentPct = parseFloat(pctRaw) || 0;
            }
            const totalPct = currentPct;

            if (peptone1 || strain) {
                const entry = {
                    code: code,
                    strain: strain,
                    name: peptone1,
                    peptone_pct: totalPct,
                    peptone_1: peptone1,
                    ratio_1: ratio1,
                };
                // 블렌딩인 경우에만 peptone_2 추가
                if (peptone2) {
                    entry.peptone_2 = peptone2;
                    entry.ratio_2 = ratio2;
                    // display name: "PEA-1(60)+SOY-1(40)"
                    entry.name = `${peptone1}(${ratio1})+${peptone2}(${ratio2})`;
                }
                result.push(entry);
            }
        });
        return result;
    }

    // ── Process ──
    btnProcess.addEventListener('click', processFile);

    async function processFile() {
        if (!selectedFile) return;

        // Show loading
        btnProcess.disabled = true;
        statusArea.style.display = 'block';
        statusLoading.style.display = 'flex';
        statusSuccess.style.display = 'none';
        statusError.style.display = 'none';
        btnDownload.style.display = 'none';

        // Build FormData
        const formData = new FormData();
        formData.append('file', selectedFile);
        formData.append('experiment_date', document.getElementById('experiment-date').value || '');
        formData.append('goal', document.getElementById('goal').value || '');
        formData.append('strain', document.getElementById('strain').value || '');
        formData.append('base_media', document.getElementById('base-media').value || '');
        formData.append('media_type', document.getElementById('media-type').value || 'peptone_screening');

        // Build payload depending on experiment type
        const expType = document.getElementById('media-type').value || 'peptone_screening';
        let payload;
        if (expType === 'media_optimization') {
            payload = collectMediaOptimization();
        } else {
            payload = collectSampleMap();
        }
        formData.append('sample_map_json', JSON.stringify(payload));

        try {
            const response = await fetch('/api/process', {
                method: 'POST',
                body: formData,
            });

            const data = await response.json();

            if (!response.ok) {
                throw new Error(data.detail || '서버 오류가 발생했습니다.');
            }

            if (data.success) {
                showSuccess(data);
            } else {
                throw new Error(data.message || '처리 오류');
            }
        } catch (err) {
            showError(err.message);
        }
    }

    function showSuccess(data) {
        statusLoading.style.display = 'none';
        statusSuccess.style.display = 'flex';
        statusMessage.textContent = data.message || '가공이 완료되었습니다!';

        downloadFileId = data.file_id;
        downloadFilename = data.filename;
        btnDownload.style.display = 'flex';
        btnDownload.disabled = false;
        btnProcess.disabled = false;

        // Update button text for re-processing
        const btnSpan = btnProcess.querySelector('span');
        if (btnSpan) {
            btnSpan.textContent = '다시 가공하기';
        }

        // Show sample mapping if groups detected
        if (data.chart_data && data.chart_data.groups && data.chart_data.groups.length > 0) {
            // Only build mapping table if not already populated
            if (detectedGroups.length === 0) {
                detectedGroups = data.chart_data.groups;
                buildMappingTable(detectedGroups);
            }
        }

        // Render chart
        if (data.chart_data) {
            renderChart(data.chart_data);
        }
    }

    function showError(msg) {
        statusLoading.style.display = 'none';
        statusError.style.display = 'flex';
        errorMessage.textContent = msg;
        btnProcess.disabled = false;
    }

    // ── Download ──
    btnDownload.addEventListener('click', () => {
        if (!downloadFileId) return;

        const link = document.createElement('a');
        link.href = `/api/download/${downloadFileId}`;
        link.download = downloadFilename || 'processed.xlsx';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });

    // ── Chart Rendering (Plotly.js) ──
    function renderChart(chartData) {
        chartPlaceholder.style.display = 'none';
        chartContainer.style.display = 'block';
        chartBadge.style.display = 'inline-block';

        const traces = [];
        const { time_hours, series } = chartData;

        series.forEach((s, idx) => {
            const color = CHART_COLORS[idx % CHART_COLORS.length];

            // Build error_y config from SD data
            const errorYConfig = (s.sd && s.sd.some(v => v !== null && v > 0))
                ? {
                    type: 'data',
                    array: s.sd.map(v => v !== null ? v : 0),
                    visible: true,
                    color: hexToRgba(color, 0.4),
                    thickness: 1.5,
                    width: 3,
                }
                : undefined;

            // Mean line with error bars
            traces.push({
                x: time_hours,
                y: s.mean,
                name: s.name,
                type: 'scatter',
                mode: 'lines+markers',
                line: {
                    color: color,
                    width: 2.5,
                    shape: 'linear',
                },
                marker: {
                    color: color,
                    size: 5,
                },
                error_y: errorYConfig,
            });
        });

        const layout = {
            paper_bgcolor: 'rgba(0,0,0,0)',
            plot_bgcolor: 'rgba(0,0,0,0)',
            font: {
                family: 'Inter, sans-serif',
                color: '#94a3b8',
                size: 12,
            },
            xaxis: {
                title: {
                    text: 'Time (h)',
                    font: { color: '#94a3b8', size: 13 },
                },
                gridcolor: 'rgba(255,255,255,0.05)',
                zeroline: false,
                linecolor: 'rgba(255,255,255,0.1)',
            },
            yaxis: {
                title: {
                    text: 'OD₆₀₀',
                    font: { color: '#94a3b8', size: 13 },
                },
                gridcolor: 'rgba(255,255,255,0.05)',
                zeroline: false,
                linecolor: 'rgba(255,255,255,0.1)',
            },
            legend: {
                bgcolor: 'rgba(0,0,0,0)',
                font: { color: '#cbd5e1', size: 11 },
                orientation: 'h',
                y: -0.2,
                x: 0.5,
                xanchor: 'center',
            },
            margin: { l: 60, r: 20, t: 20, b: 60 },
            hovermode: 'x unified',
            hoverlabel: {
                bgcolor: '#1e293b',
                bordercolor: 'rgba(99,102,241,0.3)',
                font: { color: '#f1f5f9', size: 12 },
            },
        };

        const config = {
            responsive: true,
            displayModeBar: true,
            modeBarButtonsToRemove: ['sendDataToCloud', 'lasso2d', 'select2d'],
            displaylogo: false,
        };

        Plotly.newPlot('plotly-chart', traces, layout, config);

        // Show summary
        renderSummary(chartData);
    }

    function hexToRgba(hex, alpha) {
        const r = parseInt(hex.slice(1, 3), 16);
        const g = parseInt(hex.slice(3, 5), 16);
        const b = parseInt(hex.slice(5, 7), 16);
        return `rgba(${r}, ${g}, ${b}, ${alpha})`;
    }

    // ============================================================
    // Media Optimization UI (v2 — composition groups + SM mapping)
    // ============================================================
    const SM_COUNT = 12;  // Maximum supported SM groups (SM1..SM12)

    const BASE_MEDIA_PRESETS = {
        MRS: [
            { name: 'Glucose',          value: 20,   unit: 'g/L', category: 'carbon' },
            { name: 'Peptone',          value: 10,   unit: 'g/L', category: 'nitrogen' },
            { name: 'Beef Extract',     value: 10,   unit: 'g/L', category: 'other' },
            { name: 'Yeast Extract',    value: 5,    unit: 'g/L', category: 'other' },
            { name: 'K₂HPO₄',          value: 2,    unit: 'g/L', category: 'mineral' },
            { name: 'Sodium Acetate',   value: 5,    unit: 'g/L', category: 'mineral' },
            { name: 'Ammonium Citrate', value: 2,    unit: 'g/L', category: 'mineral' },
            { name: 'MgSO₄',           value: 0.1,  unit: 'g/L', category: 'mineral' },
            { name: 'MnSO₄',           value: 0.05, unit: 'g/L', category: 'mineral' },
            { name: 'Tween80',          value: 1,    unit: 'mL/L', category: 'other' },
        ],
        TSB: [
            { name: 'Tryptone',    value: 17,  unit: 'g/L', category: 'nitrogen' },
            { name: 'Soytone',     value: 3,   unit: 'g/L', category: 'nitrogen' },
            { name: 'Glucose',     value: 2.5, unit: 'g/L', category: 'carbon' },
            { name: 'NaCl',        value: 5,   unit: 'g/L', category: 'mineral' },
            { name: 'K₂HPO₄',     value: 2.5, unit: 'g/L', category: 'mineral' },
        ],
        LB: [
            { name: 'Tryptone',      value: 10, unit: 'g/L', category: 'nitrogen' },
            { name: 'Yeast Extract', value: 5,  unit: 'g/L', category: 'other' },
            { name: 'NaCl',          value: 10, unit: 'g/L', category: 'mineral' },
        ],
        CUSTOM: [],
    };

    const CARBON_SOURCES = [
        'Glucose', 'Arabinose', 'Fructose', 'Galactose', 'Lactose',
        'Maltose', 'Mannose', 'Raffinose', 'Sucrose', 'Xylose',
    ];

    const NITROGEN_SOURCES = [
        { alias: '1',  name: 'SOY-1' },
        { alias: 'N',  name: 'SOY-N+' },
        { alias: 'L',  name: 'SOY-L' },
        { alias: 'SP', name: 'SOY-P' },
        { alias: 'B',  name: 'SOY-B' },
        { alias: 'R',  name: 'RICE-1' },
        { alias: 'W',  name: 'WHEAT-1' },
        { alias: 'P',  name: 'PEA-1' },
        { alias: 'PP', name: 'PPR Type4' },
    ];

    const MINERALS_STANDARD = [
        'MgSO₄', 'MnSO₄', 'CaCl₂', 'FeSO₄', 'ZnSO₄',
        'K₂HPO₄', 'KH₂PO₄', 'NaCl', '(NH₄)₂SO₄', 'Sodium Acetate',
    ];

    const CATEGORY_LABEL = {
        carbon: '탄소원 (Carbon Sources)',
        nitrogen: '질소원 (Peptones)',
        mineral: '무기염류 (Minerals)',
        other: '기타 성분 (Vitamins / Surfactants / Indicators)',
    };

    const CATEGORY_CHIP_SOURCE = {
        carbon: CARBON_SOURCES,
        nitrogen: NITROGEN_SOURCES,
        mineral: MINERALS_STANDARD,
        other: null,  // no chip grid (only custom adds)
    };

    // ── Generic composition editor builder ──────────────────────────────
    // Each editor owns its own state object: { carbon: Map, nitrogen: Map, mineral: Map, other: Map }
    // (Map<name, {value, unit, custom}>)
    function makeEmptyState() {
        return {
            carbon: new Map(),
            nitrogen: new Map(),
            mineral: new Map(),
            other: new Map(),
        };
    }

    function buildCompositionEditor(rootEl, state, opts = {}) {
        rootEl.innerHTML = '';
        rootEl.classList.add('composition-editor');
        const refs = {
            sections: {},   // {cat: {sectionEl, chipGrid, rowsEl, customCounter}}
        };

        ['carbon', 'nitrogen', 'mineral', 'other'].forEach((cat) => {
            const section = document.createElement('div');
            section.className = 'ce-section';
            section.dataset.category = cat;
            section.innerHTML = `
                <div class="ce-section-title">${CATEGORY_LABEL[cat]}</div>
                <div class="chip-grid"></div>
                <div class="component-rows"></div>
            `;
            const chipGrid = section.querySelector('.chip-grid');
            const rowsEl = section.querySelector('.component-rows');

            // Chip grid (only for cats with predefined chips)
            const chipSource = CATEGORY_CHIP_SOURCE[cat];
            if (chipSource) {
                chipSource.forEach((item) => {
                    const name = typeof item === 'string' ? item : item.name;
                    const alias = typeof item === 'object' ? item.alias : null;
                    const label = document.createElement('label');
                    label.className = 'chip';
                    label.dataset.name = name;
                    label.innerHTML = `
                        <input type="checkbox">
                        <span>${name}</span>
                        ${alias ? `<span class="chip-alias">(${alias})</span>` : ''}
                    `;
                    const cb = label.querySelector('input');
                    cb.addEventListener('change', () => {
                        if (cb.checked) {
                            addComponentRowGeneric(state, refs, cat, name, '', 'g/L');
                            label.classList.add('checked');
                        } else {
                            removeComponentRowGeneric(state, refs, cat, name);
                            label.classList.remove('checked');
                        }
                    });
                    chipGrid.appendChild(label);
                });
            } else {
                // hide empty chip-grid
                chipGrid.style.display = 'none';
            }

            // Custom-add button (mineral / other)
            if (cat === 'mineral' || cat === 'other') {
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'btn-small';
                btn.textContent = cat === 'mineral'
                    ? '+ 사용자 정의 무기염류'
                    : '+ 기타 성분 추가';
                btn.addEventListener('click', () => {
                    const ref = refs.sections[cat];
                    ref.customCounter = (ref.customCounter || 0) + 1;
                    const placeholder = `새 ${cat === 'mineral' ? '무기염류' : '성분'} ${ref.customCounter}`;
                    addComponentRowGeneric(state, refs, cat, placeholder, '', 'g/L', { custom: true });
                });
                section.appendChild(btn);
            }

            rootEl.appendChild(section);
            refs.sections[cat] = { sectionEl: section, chipGrid, rowsEl, customCounter: 0 };
        });

        return refs;
    }

    function addComponentRowGeneric(state, refs, cat, name, value, unit, opts = {}) {
        const ref = refs.sections[cat];
        if (!ref) return;
        // Update existing
        if (state[cat].has(name)) {
            const existing = state[cat].get(name);
            if (value !== '' && value != null) existing.value = value;
            if (unit) existing.unit = unit;
            const row = ref.rowsEl.querySelector(`[data-row-name="${CSS.escape(name)}"]`);
            if (row) {
                const valInput = row.querySelector('.row-value');
                const unitSel = row.querySelector('.row-unit');
                if (valInput && value !== '' && value != null) valInput.value = value;
                if (unitSel && unit) unitSel.value = unit;
            }
            return;
        }
        const isCustom = opts.custom || false;
        state[cat].set(name, { value: value, unit: unit || 'g/L', custom: isCustom });

        const row = document.createElement('div');
        row.className = 'component-row';
        row.dataset.rowName = name;
        row.dataset.category = cat;
        row.innerHTML = `
            ${isCustom
                ? `<input type="text" class="row-name-input" value="${name}" placeholder="성분명">`
                : `<span class="row-name">${name}</span>`}
            <input type="number" class="row-value" value="${value}" placeholder="농도" step="0.001" min="0">
            <select class="row-unit">
                <option value="g/L"${unit === 'g/L' ? ' selected' : ''}>g/L</option>
                <option value="mL/L"${unit === 'mL/L' ? ' selected' : ''}>mL/L</option>
                <option value="mg/L"${unit === 'mg/L' ? ' selected' : ''}>mg/L</option>
                <option value="mM"${unit === 'mM' ? ' selected' : ''}>mM</option>
            </select>
            <button type="button" class="row-remove" title="제거">×</button>
        `;
        const valInput = row.querySelector('.row-value');
        const unitSel = row.querySelector('.row-unit');
        const nameInput = row.querySelector('.row-name-input');
        const removeBtn = row.querySelector('.row-remove');

        valInput.addEventListener('input', () => {
            const s = state[cat].get(name);
            if (s) s.value = valInput.value;
        });
        unitSel.addEventListener('change', () => {
            const s = state[cat].get(name);
            if (s) s.unit = unitSel.value;
        });
        if (nameInput) {
            nameInput.addEventListener('change', () => {
                const newName = nameInput.value.trim();
                if (newName && newName !== name) {
                    const s = state[cat].get(name);
                    state[cat].delete(name);
                    state[cat].set(newName, s);
                    row.dataset.rowName = newName;
                    name = newName;
                }
            });
        }
        removeBtn.addEventListener('click', () => {
            removeComponentRowGeneric(state, refs, cat, name);
            // Uncheck matching chip if exists
            const chip = ref.chipGrid.querySelector(`.chip[data-name="${CSS.escape(name)}"]`);
            if (chip) {
                const cb = chip.querySelector('input');
                if (cb) cb.checked = false;
                chip.classList.remove('checked');
            }
        });
        ref.rowsEl.appendChild(row);
    }

    function removeComponentRowGeneric(state, refs, cat, name) {
        state[cat].delete(name);
        const ref = refs.sections[cat];
        if (!ref) return;
        const row = ref.rowsEl.querySelector(`[data-row-name="${CSS.escape(name)}"]`);
        if (row) row.remove();
    }

    function clearEditor(state, refs) {
        ['carbon', 'nitrogen', 'mineral', 'other'].forEach((cat) => {
            state[cat].clear();
            const ref = refs.sections[cat];
            if (!ref) return;
            ref.rowsEl.innerHTML = '';
            ref.chipGrid.querySelectorAll('.chip').forEach((c) => {
                c.classList.remove('checked');
                const cb = c.querySelector('input');
                if (cb) cb.checked = false;
            });
        });
    }

    function applyPresetToEditor(state, refs, presetName) {
        clearEditor(state, refs);
        if (presetName === 'NONE') return;
        const preset = BASE_MEDIA_PRESETS[presetName] || [];
        preset.forEach((c) => {
            addComponentRowGeneric(state, refs, c.category, c.name, c.value, c.unit);
            const ref = refs.sections[c.category];
            if (!ref) return;
            const chip = ref.chipGrid.querySelector(`.chip[data-name="${CSS.escape(c.name)}"]`);
            if (chip) {
                const cb = chip.querySelector('input');
                if (cb) cb.checked = true;
                chip.classList.add('checked');
            }
        });
    }

    function snapshotComposition(state) {
        // → list of {name, value, unit, category}
        const list = [];
        ['carbon', 'nitrogen', 'mineral', 'other'].forEach((cat) => {
            state[cat].forEach((s, name) => {
                const v = parseFloat(s.value);
                list.push({
                    name: name,
                    value: isNaN(v) ? 0 : v,
                    unit: s.unit || 'g/L',
                    category: cat,
                });
            });
        });
        return list;
    }

    function loadCompositionIntoEditor(state, refs, compList) {
        clearEditor(state, refs);
        (compList || []).forEach((c) => {
            const cat = c.category || 'other';
            const ref = refs.sections[cat];
            const isPredefined = ref && ref.chipGrid.querySelector(`.chip[data-name="${CSS.escape(c.name)}"]`);
            addComponentRowGeneric(state, refs, cat, c.name, c.value, c.unit, { custom: !isPredefined });
            if (isPredefined) {
                const chip = ref.chipGrid.querySelector(`.chip[data-name="${CSS.escape(c.name)}"]`);
                const cb = chip.querySelector('input');
                if (cb) cb.checked = true;
                chip.classList.add('checked');
            }
        });
    }

    // ── Base Medium (preset-only, no editor) ──────────────────────────
    // Base Medium is a pure *template selector*. The user cannot edit its
    // composition directly — instead, the selected preset (MRS/TSB/LB)
    // auto-fills every new composition group, and provides the reference
    // for "Base 대비 차이" highlighting in the Excel output.
    //
    // preset = "NONE" → no reference base; groups start empty and are
    // treated as standalone in the Excel.
    function getBasePreset() {
        const sel = document.getElementById('base-medium-preset');
        return sel ? sel.value : 'NONE';
    }

    function getBaseComposition() {
        const preset = getBasePreset();
        if (preset === 'NONE') return [];
        const tpl = BASE_MEDIA_PRESETS[preset] || [];
        // Return a deep-copy so groups can safely mutate their own
        return tpl.map(c => ({ ...c }));
    }

    function renderBasePreview() {
        const wrap = document.getElementById('base-preview-wrap');
        const content = document.getElementById('base-preview-content');
        if (!wrap || !content) return;
        const preset = getBasePreset();
        const comp = getBaseComposition();
        if (preset === 'NONE' || comp.length === 0) {
            // Preview is useless when no reference composition
            wrap.style.display = 'none';
            return;
        }
        wrap.style.display = 'block';
        content.classList.remove('bp-empty');
        content.innerHTML = comp.map(c => {
            const catLabel = ({carbon:'탄소원', nitrogen:'질소원', mineral:'무기염류', other:'기타'})[c.category] || '';
            return `<div class="bp-item"><span class="bp-cat">${catLabel}</span>${escapeHtml(c.name)}<span class="bp-val">${c.value} ${c.unit || ''}</span></div>`;
        }).join('');
    }

    // Small util — HTML-escape for inline injection
    function escapeHtml(s) {
        if (s == null) return '';
        return String(s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    }

    // ── Composition Groups (cards) ────────────────────────────────────
    const compositionGroups = [];   // [{id, state, refs, ...}] (refs filled when card built)
    let cgCounter = 0;

    function newGroupId() {
        cgCounter += 1;
        return `cg-${cgCounter}`;
    }

    function addCompositionGroup(initial = {}) {
        const group = {
            id: initial.id || newGroupId(),
            name: initial.name || '',
            strain: initial.strain || '',
            description: initial.description || '',
            state: makeEmptyState(),
            appliedSamples: new Set(initial.applied_samples || []),
            cardEl: null,
            editorRefs: null,
        };
        compositionGroups.push(group);
        renderCompositionGroup(group);
        // Initialize composition: passed-in > Base preset > empty
        if (initial.composition) {
            loadCompositionIntoEditor(group.state, group.editorRefs, initial.composition);
        } else {
            const baseComp = getBaseComposition();   // from preset (MRS/TSB/LB) or []
            if (baseComp.length > 0) {
                loadCompositionIntoEditor(group.state, group.editorRefs, baseComp);
            }
        }
        renumberGroups();
        refreshSmChipsAcrossGroups();
        updateEmptyState();
        return group;
    }

    function removeCompositionGroup(groupId) {
        const idx = compositionGroups.findIndex((g) => g.id === groupId);
        if (idx === -1) return;
        const group = compositionGroups[idx];
        if (group.cardEl) group.cardEl.remove();
        compositionGroups.splice(idx, 1);
        renumberGroups();
        refreshSmChipsAcrossGroups();
        updateEmptyState();
    }

    function duplicateCompositionGroup(groupId) {
        const orig = compositionGroups.find((g) => g.id === groupId);
        if (!orig) return;
        addCompositionGroup({
            name: orig.name ? `${orig.name} (복제)` : '',
            strain: orig.strain,
            description: orig.description,
            composition: snapshotComposition(orig.state),
            applied_samples: [],   // user picks SMs for the new copy
        });
    }

    function copyBaseToGroup(groupId) {
        const group = compositionGroups.find((g) => g.id === groupId);
        if (!group) return;
        const baseComp = getBaseComposition();
        if (baseComp.length === 0) {
            alert('Base Medium이 "직접 입력"이라 복사할 조성이 없습니다.\n(MRS/TSB/LB 를 선택하면 해당 조성으로 초기화 가능)');
            return;
        }
        if (!confirm(`이 그룹의 조성을 ${getBasePreset()} 기본 조성으로 초기화할까요? (현재 수정 내용은 덮어써집니다)`)) return;
        loadCompositionIntoEditor(group.state, group.editorRefs, baseComp);
    }

    function renumberGroups() {
        compositionGroups.forEach((g, i) => {
            const numEl = g.cardEl && g.cardEl.querySelector('.cg-num');
            if (numEl) numEl.textContent = `조성 그룹 ${i + 1}`;
        });
    }

    function getClaimedSmMap() {
        // Returns Map<SM, [groupId, ...]> for cross-group duplicate hint
        const m = new Map();
        compositionGroups.forEach((g) => {
            g.appliedSamples.forEach((sm) => {
                if (!m.has(sm)) m.set(sm, []);
                m.get(sm).push(g.id);
            });
        });
        return m;
    }

    function refreshSmChipsAcrossGroups() {
        const claimed = getClaimedSmMap();
        compositionGroups.forEach((g) => {
            if (!g.cardEl) return;
            g.cardEl.querySelectorAll('.sm-chip').forEach((btn) => {
                const sm = btn.dataset.sm;
                const inThisGroup = g.appliedSamples.has(sm);
                btn.classList.toggle('sm-chip-active', inThisGroup);
                const owners = claimed.get(sm) || [];
                const claimedByOther = !inThisGroup && owners.length > 0;
                btn.classList.toggle('sm-chip-claimed', claimedByOther);
                btn.title = claimedByOther
                    ? `이미 다른 조성 그룹에 매핑됨`
                    : (inThisGroup ? '클릭 시 해제' : '클릭하여 이 조성에 매핑');
            });
            // Update summary line
            const summary = g.cardEl.querySelector('.sm-summary');
            if (summary) {
                const arr = Array.from(g.appliedSamples);
                arr.sort((a, b) => parseInt(a.slice(2)) - parseInt(b.slice(2)));
                summary.textContent = arr.length
                    ? `매핑된 실험구: ${arr.join(', ')}  (${arr.length}개)`
                    : '아직 매핑된 실험구 없음';
            }
        });
    }

    function updateEmptyState() {
        const container = document.getElementById('comp-groups-container');
        if (!container) return;
        let empty = container.querySelector('.cg-empty-state');
        if (compositionGroups.length === 0) {
            if (!empty) {
                empty = document.createElement('div');
                empty.className = 'cg-empty-state';
                empty.innerHTML = '아직 조성 그룹이 없습니다.<br>아래 <b>"+ 조성 그룹 추가"</b> 버튼을 눌러 첫 번째 조성을 만드세요.';
                container.appendChild(empty);
            }
        } else if (empty) {
            empty.remove();
        }
    }

    function renderCompositionGroup(group) {
        const container = document.getElementById('comp-groups-container');
        if (!container) return;

        const card = document.createElement('div');
        card.className = 'comp-group-card';
        card.dataset.groupId = group.id;

        // Header
        const header = document.createElement('div');
        header.className = 'cg-header';
        header.innerHTML = `
            <span class="cg-num">조성 그룹</span>
            <div class="cg-header-actions">
                <button type="button" class="cg-action-btn cg-base-copy-btn" title="Base Medium 조성을 이 그룹에 복사">⤴ Base 복사</button>
                <button type="button" class="cg-action-btn cg-dup-btn" title="이 그룹을 복제">⎘ 복제</button>
                <button type="button" class="cg-action-btn cg-remove-btn" title="이 그룹 삭제">×</button>
            </div>
        `;
        card.appendChild(header);

        // Meta row 1: 조성 이름 + 균주
        const meta1 = document.createElement('div');
        meta1.className = 'cg-meta-row';
        meta1.innerHTML = `
            <input type="text" class="form-input cg-name" placeholder="조성 이름 (예: Control, +Mg 2x, No Glucose)">
            <input type="text" class="form-input cg-strain" placeholder="균주 (예: LR / 공배양: LR, LP)">
        `;
        card.appendChild(meta1);

        // Meta row 2: 설명 (full width)
        const meta2 = document.createElement('div');
        meta2.className = 'cg-meta-row-full';
        meta2.innerHTML = `
            <input type="text" class="form-input cg-desc" placeholder="설명 (선택, 예: Mg 2배 증량 / 탄소원 제거)">
        `;
        card.appendChild(meta2);

        // Composition editor
        const editorWrap = document.createElement('div');
        editorWrap.className = 'cg-composition-wrap';
        const editorLabel = document.createElement('div');
        editorLabel.className = 'cg-section-label';
        editorLabel.textContent = '배지 조성';
        editorWrap.appendChild(editorLabel);
        const editor = document.createElement('div');
        editorWrap.appendChild(editor);
        card.appendChild(editorWrap);

        // SM mapping
        const smWrap = document.createElement('div');
        smWrap.innerHTML = `
            <div class="cg-section-label">적용 실험구 (SM1 ~ SM${SM_COUNT}) — 클릭하여 토글</div>
            <div class="sm-chip-row"></div>
            <div class="sm-summary">아직 매핑된 실험구 없음</div>
        `;
        const chipRow = smWrap.querySelector('.sm-chip-row');
        for (let i = 1; i <= SM_COUNT; i += 1) {
            const sm = `SM${i}`;
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'sm-chip';
            btn.dataset.sm = sm;
            btn.textContent = sm;
            btn.addEventListener('click', () => {
                if (group.appliedSamples.has(sm)) {
                    group.appliedSamples.delete(sm);
                } else {
                    group.appliedSamples.add(sm);
                }
                refreshSmChipsAcrossGroups();
            });
            chipRow.appendChild(btn);
        }
        card.appendChild(smWrap);

        // Bind meta inputs
        const nameInput = meta1.querySelector('.cg-name');
        const strainInput = meta1.querySelector('.cg-strain');
        const descInput = meta2.querySelector('.cg-desc');
        nameInput.value = group.name;
        strainInput.value = group.strain;
        descInput.value = group.description;
        nameInput.addEventListener('input', () => { group.name = nameInput.value.trim(); });
        strainInput.addEventListener('input', () => { group.strain = strainInput.value.trim(); });
        descInput.addEventListener('input', () => { group.description = descInput.value.trim(); });

        // Action buttons
        header.querySelector('.cg-remove-btn').addEventListener('click', () => {
            if (compositionGroups.length === 1 && !confirm('마지막 조성 그룹을 삭제하시겠습니까?')) return;
            removeCompositionGroup(group.id);
        });
        header.querySelector('.cg-dup-btn').addEventListener('click', () => duplicateCompositionGroup(group.id));
        header.querySelector('.cg-base-copy-btn').addEventListener('click', () => copyBaseToGroup(group.id));

        // Build editor
        group.editorRefs = buildCompositionEditor(editor, group.state);

        container.appendChild(card);
        group.cardEl = card;
    }

    // ── Payload collection (new format) ──────────────────────────────
    function collectMediaOptimization() {
        const preset = getBasePreset();
        const baseComposition = getBaseComposition();   // derived from preset

        const groups = compositionGroups.map((g) => ({
            id: g.id,
            name: g.name,
            strain: g.strain,
            description: g.description,
            composition: snapshotComposition(g.state),
            applied_samples: Array.from(g.appliedSamples).sort(
                (a, b) => parseInt(a.slice(2)) - parseInt(b.slice(2))
            ),
        }));

        return {
            experiment_type: 'media_optimization',
            base_medium: {
                preset: preset,                       // "MRS" / "TSB" / "LB" / "NONE"
                custom_name: '',                      // (deprecated, kept for server back-compat)
                composition: baseComposition,         // empty [] when preset="NONE"
            },
            composition_groups: groups,
        };
    }

    function initMediaOptimization() {
        const presetSelect = document.getElementById('base-medium-preset');
        if (presetSelect) {
            presetSelect.value = 'MRS';   // default template
            presetSelect.addEventListener('change', renderBasePreview);
        }

        // Add composition group button
        const addBtn = document.getElementById('btn-add-comp-group');
        if (addBtn) {
            addBtn.addEventListener('click', () => addCompositionGroup());
        }

        // Initial render of Base preview
        renderBasePreview();

        // Start with empty-state hint
        updateEmptyState();
    }

    initMediaOptimization();

    // ── Experiment type toggle ──
    const mediaTypeSelect = document.getElementById('media-type');
    const mediaOptCard = document.getElementById('media-opt-card');

    function updateExperimentTypeView() {
        const v = mediaTypeSelect.value;
        if (v === 'media_optimization') {
            sampleMapCard.style.display = 'none';
            mediaOptCard.style.display = 'block';
        } else {
            mediaOptCard.style.display = 'none';
            // sampleMapCard visibility handled by detection logic (appears after process)
        }
    }

    mediaTypeSelect.addEventListener('change', updateExperimentTypeView);
    updateExperimentTypeView();

    // ── Data Summary ──
    function renderSummary(chartData) {
        summaryCard.style.display = 'block';
        summaryGrid.innerHTML = '';

        const nGroups = chartData.series.length;
        const nTimepoints = chartData.time_hours.length;
        const totalTime = chartData.time_hours[chartData.time_hours.length - 1] || 0;

        // Max OD across all groups
        let maxOD = 0;
        chartData.series.forEach(s => {
            const localMax = Math.max(...s.mean.filter(v => v !== null));
            if (localMax > maxOD) {
                maxOD = localMax;
            }
        });

        const summaryItems = [
            { value: nGroups, label: '샘플 그룹' },
            { value: nTimepoints, label: 'Timepoints' },
            { value: totalTime.toFixed(1) + 'h', label: '총 배양시간' },
            { value: maxOD.toFixed(3), label: 'Max OD₆₀₀' },
        ];

        summaryItems.forEach(item => {
            const el = document.createElement('div');
            el.className = 'summary-item';
            el.innerHTML = `
                <div class="summary-value">${item.value}</div>
                <div class="summary-label">${item.label}</div>
            `;
            summaryGrid.appendChild(el);
        });
    }
})();
