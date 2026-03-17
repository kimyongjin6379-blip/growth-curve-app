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
                <td><input type="text" class="form-input mapping-name" data-code="${grp}" placeholder="샘플명 입력"></td>
                <td><input type="number" class="form-input mapping-pct" data-code="${grp}" placeholder="0.0" step="0.1" min="0"></td>
            `;
            mappingTbody.appendChild(tr);
        });
        sampleMapCard.style.display = 'block';
    }

    function collectSampleMap() {
        const rows = mappingTbody.querySelectorAll('tr');
        const result = [];
        rows.forEach((tr) => {
            const code = tr.querySelector('.mapping-code').textContent.trim();
            const name = tr.querySelector('.mapping-name').value.trim();
            const pctInput = tr.querySelector('.mapping-pct').value.trim();
            if (name) {
                result.push({
                    code: code,
                    name: name,
                    peptone_pct: pctInput !== '' ? parseFloat(pctInput) : 0.0,
                });
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

        // Collect sample mapping if available
        const sampleMap = collectSampleMap();
        formData.append('sample_map_json', JSON.stringify(sampleMap));

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
