const API_BASE_URL = window.location.origin;
let currentReportKey = 'missing-manager';
let latestRecords = [];

const FILTER_REASON_I18N_KEYS = {
    filter_disabled: 'reports.types.filteredLicensed.reason.disabled',
    filter_guest: 'reports.types.filteredLicensed.reason.guest',
    filter_no_title: 'reports.types.filteredLicensed.reason.noTitle',
    filter_ignored_title: 'reports.types.filteredLicensed.reason.ignoredTitle',
    filter_ignored_department: 'reports.types.filteredLicensed.reason.ignoredDepartment',
    filter_ignored_employee: 'reports.types.filteredLicensed.reason.ignoredEmployee',
};

const REPORT_CONFIGS = {
    'missing-manager': {
        dataPath: '/api/reports/missing-manager',
        exportPath: '/api/reports/missing-manager/export',
        summaryLabelKey: 'reports.summary.totalLabel',
        tableTitleKey: 'reports.table.title',
        emptyKey: 'reports.table.empty',
        countSummaryKey: 'reports.table.countSummary',
        buildStatusParams: (records) => ({ count: records.length }),
        columns: [
            { key: 'name', labelKey: 'reports.table.columns.name' },
            { key: 'title', labelKey: 'reports.table.columns.title' },
            { key: 'department', labelKey: 'reports.table.columns.department' },
            { key: 'email', labelKey: 'reports.table.columns.email' },
            { key: 'managerName', labelKey: 'reports.table.columns.manager' },
            {
                key: 'reason',
                labelKey: 'reports.table.columns.reason',
                render: (record, t) => createReasonBadge(record.reason, t),
            },
        ],
    },
    'disabled-licensed': {
        dataPath: '/api/reports/disabled-licensed',
        exportPath: '/api/reports/disabled-licensed/export',
        summaryLabelKey: 'reports.types.disabledLicensed.summaryLabel',
        tableTitleKey: 'reports.types.disabledLicensed.tableTitle',
        emptyKey: 'reports.types.disabledLicensed.empty',
        countSummaryKey: 'reports.types.disabledLicensed.countSummary',
        showLicenseSummary: true,
        licenseSummaryLabelKey: 'reports.summary.licensesLabel',
        buildStatusParams: (records) => ({
            count: records.length,
            licenses: records.reduce((total, item) => total + (item.licenseCount || 0), 0),
        }),
        columns: [
            { key: 'name', labelKey: 'reports.table.columns.name' },
            { key: 'title', labelKey: 'reports.table.columns.title' },
            { key: 'department', labelKey: 'reports.table.columns.department' },
            { key: 'email', labelKey: 'reports.table.columns.email' },
            { key: 'licenseCount', labelKey: 'reports.table.columns.licenseCount' },
            {
                key: 'licenseSkus',
                labelKey: 'reports.table.columns.licenses',
                render: (record) => (record.licenseSkus || []).join(', '),
            },
        ],
    },
    'filtered-licensed': {
        dataPath: '/api/reports/filtered-licensed',
        exportPath: '/api/reports/filtered-licensed/export',
        summaryLabelKey: 'reports.types.filteredLicensed.summaryLabel',
        tableTitleKey: 'reports.types.filteredLicensed.tableTitle',
        emptyKey: 'reports.types.filteredLicensed.empty',
        countSummaryKey: 'reports.types.filteredLicensed.countSummary',
        showLicenseSummary: true,
        licenseSummaryLabelKey: 'reports.summary.licensesLabel',
        buildStatusParams: (records) => ({
            count: records.length,
            licenses: records.reduce((total, item) => total + (item.licenseCount || 0), 0),
        }),
        columns: [
            { key: 'name', labelKey: 'reports.table.columns.name' },
            { key: 'title', labelKey: 'reports.table.columns.title' },
            { key: 'department', labelKey: 'reports.table.columns.department' },
            { key: 'email', labelKey: 'reports.table.columns.email' },
            { key: 'licenseCount', labelKey: 'reports.table.columns.licenseCount' },
            {
                key: 'licenseSkus',
                labelKey: 'reports.table.columns.licenses',
                render: (record) => (record.licenseSkus || []).join(', '),
            },
            {
                key: 'filterReasons',
                labelKey: 'reports.types.filteredLicensed.columns.filterReasons',
                render: renderFilterReasonsCell,
            },
        ],
    },
    'filtered-users': {
        dataPath: '/api/reports/filtered-users',
        exportPath: '/api/reports/filtered-users/export',
        summaryLabelKey: 'reports.types.filteredUsers.summaryLabel',
        tableTitleKey: 'reports.types.filteredUsers.tableTitle',
        emptyKey: 'reports.types.filteredUsers.empty',
        countSummaryKey: 'reports.types.filteredUsers.countSummary',
        buildStatusParams: (records) => ({ count: records.length }),
        columns: [
            { key: 'name', labelKey: 'reports.table.columns.name' },
            { key: 'title', labelKey: 'reports.table.columns.title' },
            { key: 'department', labelKey: 'reports.table.columns.department' },
            { key: 'email', labelKey: 'reports.table.columns.email' },
            {
                key: 'filterReasons',
                labelKey: 'reports.types.filteredLicensed.columns.filterReasons',
                render: renderFilterReasonsCell,
            },
        ],
    },
};

function renderFilterReasonsCell(record, t) {
    const reasons = record.filterReasons || [];
    if (!reasons.length) {
        return defaultCellValue([]);
    }

    const container = document.createElement('div');
    container.className = 'reason-badges';

    reasons.forEach((reasonKey) => {
        const badge = document.createElement('span');
        badge.className = 'badge badge--neutral';
        badge.textContent = t(FILTER_REASON_I18N_KEYS[reasonKey] || reasonKey);
        container.appendChild(badge);
    });

    return container;
}

function qs(id) {
    return document.getElementById(id);
}

function getTranslator() {
    return window.i18n?.t || ((key) => key);
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

function applyReportContext(config) {
    const t = getTranslator();
    const labelEl = qs('primarySummaryLabel');
    if (labelEl) {
        labelEl.textContent = t(config.summaryLabelKey);
    }
    const titleEl = qs('tableTitle');
    if (titleEl) {
        titleEl.textContent = t(config.tableTitleKey);
    }
}

function toggleLoading(isLoading, config, records = []) {
    const refreshBtn = qs('refreshReportBtn');
    const exportBtn = qs('exportReportBtn');
    const statusEl = qs('tableStatus');
    const t = getTranslator();

    if (refreshBtn) {
        refreshBtn.disabled = isLoading;
        refreshBtn.setAttribute('aria-busy', String(isLoading));
    }
    if (exportBtn) {
        exportBtn.disabled = isLoading || records.length === 0;
    }
    if (statusEl) {
        statusEl.textContent = isLoading
            ? t('reports.table.loading')
            : t(config.countSummaryKey, config.buildStatusParams(records));
    }
}

function showError(messageKey, detail) {
    const banner = qs('errorBanner');
    const t = getTranslator();
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

function renderSummary(records, generatedAt, config) {
    applyReportContext(config);
    const countEl = qs('primarySummaryValue');
    const generatedEl = qs('generatedAt');
    const licenseCard = qs('licenseSummaryCard');
    const licenseLabel = qs('licenseSummaryLabel');
    const licenseValue = qs('licenseSummaryValue');
    const t = getTranslator();
    const summaryMetrics = config.buildStatusParams ? config.buildStatusParams(records) : { count: records.length };

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

    if (licenseCard && licenseLabel && licenseValue) {
        if (config.showLicenseSummary) {
            const labelKey = config.licenseSummaryLabelKey || 'reports.summary.licensesLabel';
            licenseLabel.textContent = t(labelKey);
            const totalLicenses = summaryMetrics.licenses ?? 0;
            licenseValue.textContent = Number.isFinite(totalLicenses)
                ? totalLicenses.toLocaleString()
                : '—';
            licenseCard.classList.remove('is-hidden');
        } else {
            licenseCard.classList.add('is-hidden');
        }
    }
}

function defaultCellValue(value) {
    if (Array.isArray(value)) {
        return value.length ? value.join(', ') : '—';
    }
    if (value === null || value === undefined || value === '') {
        return '—';
    }
    return value;
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

function createReasonBadge(reason, t) {
    const badge = document.createElement('span');
    badge.className = reasonBadgeClass(reason);
    const labels = {
        no_manager: 'reports.table.reasonLabels.no_manager',
        manager_not_found: 'reports.table.reasonLabels.manager_not_found',
        detached: 'reports.table.reasonLabels.detached',
    };
    const labelKey = labels[reason] || 'reports.table.reasonLabels.unknown';
    badge.textContent = t(labelKey);
    return badge;
}

function renderTable(records, config) {
    const thead = qs('reportTableHead');
    const tbody = qs('reportTableBody');
    const t = getTranslator();

    if (!thead || !tbody) {
        return;
    }

    thead.innerHTML = '';
    const headerRow = document.createElement('tr');
    config.columns.forEach((column) => {
        const th = document.createElement('th');
        th.textContent = t(column.labelKey);
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    tbody.innerHTML = '';

    if (!records.length) {
        const emptyRow = document.createElement('tr');
        emptyRow.className = 'empty-row';
        const cell = document.createElement('td');
        cell.colSpan = config.columns.length;
        cell.textContent = t(config.emptyKey);
        emptyRow.appendChild(cell);
        tbody.appendChild(emptyRow);
        return;
    }

    records.forEach((record) => {
        const row = document.createElement('tr');
        config.columns.forEach((column) => {
            const cell = document.createElement('td');
            let value;
            if (column.render) {
                value = column.render(record, t);
            } else {
                value = defaultCellValue(record[column.key]);
            }

            if (value instanceof HTMLElement) {
                cell.appendChild(value);
            } else {
                cell.textContent = value || '—';
            }

            row.appendChild(cell);
        });
        tbody.appendChild(row);
    });
}

async function loadReport({ refresh = false } = {}) {
    const config = REPORT_CONFIGS[currentReportKey] || REPORT_CONFIGS['missing-manager'];
    toggleLoading(true, config);
    clearError();

    try {
        const url = new URL(config.dataPath, API_BASE_URL);
        if (refresh) {
            url.searchParams.set('refresh', 'true');
        }
        const response = await fetch(url, { credentials: 'include' });
        if (!response.ok) {
            throw new Error(`${response.status}`);
        }
        const payload = await response.json();
        latestRecords = Array.isArray(payload.records) ? payload.records : [];
        renderSummary(latestRecords, payload.generatedAt, config);
        renderTable(latestRecords, config);
        toggleLoading(false, config, latestRecords);
    } catch (error) {
        toggleLoading(false, config, []);
        showError('reports.errors.loadFailed', error.message);
        renderSummary([], null, config);
        renderTable([], config);
    }
}

async function exportReport() {
    const config = REPORT_CONFIGS[currentReportKey] || REPORT_CONFIGS['missing-manager'];
    clearError();

    try {
        const url = new URL(config.exportPath, API_BASE_URL);
        const response = await fetch(url, { credentials: 'include' });
        if (!response.ok) {
            throw new Error(`${response.status}`);
        }

        const blob = await response.blob();
        let filename = `report-${new Date().toISOString().slice(0, 10)}.xlsx`;
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
    const htmlElement = document.documentElement;
    const i18nReadyPromise = window.i18n?.ready;

    try {
        if (i18nReadyPromise && typeof i18nReadyPromise.then === 'function') {
            try {
                await i18nReadyPromise;
            } catch (error) {
                console.error('Failed to load translations for reports page:', error);
            }
        }

        const reportSelect = qs('reportTypeSelect');
        if (reportSelect) {
            reportSelect.value = currentReportKey;
            reportSelect.addEventListener('change', () => {
                currentReportKey = reportSelect.value;
                const config = REPORT_CONFIGS[currentReportKey] || REPORT_CONFIGS['missing-manager'];
                renderSummary([], null, config);
                renderTable([], config);
                loadReport().catch((error) => {
                    console.error('Failed to load report:', error);
                });
            });
        }

        const refreshBtn = qs('refreshReportBtn');
        if (refreshBtn) {
            refreshBtn.addEventListener('click', () => loadReport({ refresh: true }));
            const t = getTranslator();
            refreshBtn.title = t('reports.buttons.refreshTooltip');
        }

        const exportBtn = qs('exportReportBtn');
        if (exportBtn) {
            exportBtn.addEventListener('click', exportReport);
            const t = getTranslator();
            exportBtn.title = t('reports.buttons.exportTooltip');
            exportBtn.disabled = true;
        }

        const initialConfig = REPORT_CONFIGS[currentReportKey];
        renderSummary([], null, initialConfig);
        renderTable([], initialConfig);

        await loadReport();
    } finally {
        htmlElement.classList.remove('i18n-loading');
    }
}

document.addEventListener('DOMContentLoaded', () => {
    initializeReportsPage().catch((error) => {
        console.error('Failed to initialize reports page:', error);
        showError('reports.errors.initializationFailed');
    });
});
