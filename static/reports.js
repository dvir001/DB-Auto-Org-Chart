const API_BASE_URL = window.location.origin;
let latestRecords = [];

function qs(id) {
    return document.getElementById(id);
}

function formatDate(value) {
    if (!value) {
        return null;
    }
    try {
        const date = new Date(value);
        if (Number.isNaN(date.getTime())) {
            return value;
        }
        const pad = (num) => String(num).padStart(2, '0');
        const year = date.getFullYear();
        const month = pad(date.getMonth() + 1);
        const day = pad(date.getDate());
        const hours = pad(date.getHours());
        const minutes = pad(date.getMinutes());
        const seconds = pad(date.getSeconds());

        const offsetMinutes = -date.getTimezoneOffset();
        const offsetSign = offsetMinutes >= 0 ? '+' : '-';
        const absOffset = Math.abs(offsetMinutes);
        const offsetHours = pad(Math.floor(absOffset / 60));
        const offsetMins = pad(absOffset % 60);

        return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}${offsetSign}${offsetHours}:${offsetMins}`;
    } catch (error) {
        return value;
    }
}

function toggleLoading(isLoading) {
    const refreshBtn = qs('refreshReportBtn');
    const exportBtn = qs('exportReportBtn');
    const statusEl = qs('tableStatus');
    const t = window.i18n?.t || ((key) => key);

    if (refreshBtn) {
        refreshBtn.disabled = isLoading;
        refreshBtn.setAttribute('aria-busy', String(isLoading));
    }
    if (exportBtn) {
        exportBtn.disabled = isLoading || latestRecords.length === 0;
    }
    if (statusEl) {
        const key = isLoading ? 'reports.table.loading' : 'reports.table.updated';
        statusEl.textContent = t(key);
    }
}

function showError(messageKey, detail) {
    const banner = qs('errorBanner');
    const t = window.i18n?.t || ((key) => key);
    if (!banner) {
        return;
    }
    const message = t(messageKey);
    banner.textContent = detail ? `${message} ${detail}` : message;
    banner.classList.remove('is-hidden');
}

function clearError() {
    const banner = qs('errorBanner');
    if (banner) {
        banner.classList.add('is-hidden');
        banner.textContent = '';
    }
}

function renderSummary(records, generatedAt) {
    const countEl = qs('missingCount');
    const generatedEl = qs('generatedAt');
    const t = window.i18n?.t || ((key) => key);

    if (countEl) {
        countEl.textContent = records.length.toLocaleString();
    }
    if (generatedEl) {
        if (generatedAt) {
            generatedEl.textContent = formatDate(generatedAt);
        } else {
            generatedEl.textContent = t('reports.summary.generatedPending');
        }
    }
}

function reasonBadgeClass(reason) {
    switch (reason) {
        case 'manager_not_found':
            return 'badge badge--danger';
        case 'detached':
            return 'badge badge--info';
        default:
            return 'badge badge--warning';
    }
}

function renderTable(records) {
    const tbody = qs('reportTableBody');
    const t = window.i18n?.t || ((key) => key);

    if (!tbody) {
        return;
    }

    tbody.innerHTML = '';

    if (!records.length) {
        const emptyRow = document.createElement('tr');
        emptyRow.className = 'empty-row';
        const cell = document.createElement('td');
        cell.colSpan = 8;
        cell.textContent = t('reports.table.empty');
        emptyRow.appendChild(cell);
        tbody.appendChild(emptyRow);
        return;
    }

    const reasonLabelFor = (reason) => {
        const labels = {
            no_manager: 'reports.table.reasonLabels.no_manager',
            manager_not_found: 'reports.table.reasonLabels.manager_not_found',
            detached: 'reports.table.reasonLabels.detached',
        };
        const key = labels[reason] || 'reports.table.reasonLabels.unknown';
        return t(key);
    };

    records.forEach((record) => {
        const row = document.createElement('tr');
        const cells = [
            record.name || '—',
            record.title || '—',
            record.department || '—',
            record.email || '—',
            record.phone || '—',
            record.location || '—',
            record.managerName || '—',
        ];

        cells.forEach((value) => {
            const cell = document.createElement('td');
            cell.textContent = value;
            row.appendChild(cell);
        });

        const reasonCell = document.createElement('td');
        const badge = document.createElement('span');
        badge.className = reasonBadgeClass(record.reason);
        badge.textContent = reasonLabelFor(record.reason);
        reasonCell.appendChild(badge);
        row.appendChild(reasonCell);

        tbody.appendChild(row);
    });
}

async function loadReport({ refresh = false } = {}) {
    const t = window.i18n?.t || ((key) => key);
    toggleLoading(true);
    clearError();

    try {
        const url = new URL('/api/reports/missing-manager', API_BASE_URL);
        if (refresh) {
            url.searchParams.set('refresh', 'true');
        }
        const response = await fetch(url, { credentials: 'include' });
        if (!response.ok) {
            throw new Error(`${response.status}`);
        }
        const payload = await response.json();
        latestRecords = Array.isArray(payload.records) ? payload.records : [];
        renderSummary(latestRecords, payload.generatedAt);
        renderTable(latestRecords);
        toggleLoading(false);
        const statusEl = qs('tableStatus');
        if (statusEl) {
            const countText = t('reports.table.countSummary', { count: latestRecords.length });
            statusEl.textContent = countText;
        }
    } catch (error) {
        toggleLoading(false);
        showError('reports.errors.loadFailed', error.message);
        renderSummary([], null);
        renderTable([]);
    }
}

async function exportReport() {
    const t = window.i18n?.t || ((key) => key);
    clearError();

    try {
        const url = new URL('/api/reports/missing-manager/export', API_BASE_URL);
        const response = await fetch(url, { credentials: 'include' });
        if (!response.ok) {
            throw new Error(`${response.status}`);
        }

        const blob = await response.blob();
        let filename = `missing-managers-${new Date().toISOString().slice(0, 10)}.xlsx`;
        const disposition = response.headers.get('Content-Disposition') || response.headers.get('content-disposition');
        if (disposition) {
            const match = disposition.match(/filename="?([^";]+)"?/i);
            if (match && match[1]) {
                filename = match[1];
            }
        }

        const blobUrl = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = blobUrl;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        link.remove();
        window.URL.revokeObjectURL(blobUrl);
    } catch (error) {
        showError('reports.errors.exportFailed', error.message);
    }
}

async function initializeReportsPage() {
    await window.i18n?.ready;
    const t = window.i18n?.t || ((key) => key);

    const refreshBtn = qs('refreshReportBtn');
    if (refreshBtn) {
        refreshBtn.addEventListener('click', () => loadReport({ refresh: true }));
        refreshBtn.title = t('reports.buttons.refreshTooltip');
    }

    const exportBtn = qs('exportReportBtn');
    if (exportBtn) {
        exportBtn.addEventListener('click', exportReport);
        exportBtn.title = t('reports.buttons.exportTooltip');
        exportBtn.disabled = true;
    }

    await loadReport();
}

document.addEventListener('DOMContentLoaded', () => {
    initializeReportsPage().catch((error) => {
        console.error('Failed to initialize reports page:', error);
        showError('reports.errors.initializationFailed');
    });
});
