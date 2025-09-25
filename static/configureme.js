const API_BASE_URL = window.location.origin;
let currentSettings = {};

const EXPORT_COLUMN_KEYS = [
    { key: 'name', selectId: 'exportColumnName' },
    { key: 'title', selectId: 'exportColumnTitle' },
    { key: 'department', selectId: 'exportColumnDepartment' },
    { key: 'email', selectId: 'exportColumnEmail' },
    { key: 'phone', selectId: 'exportColumnPhone' },
    { key: 'hireDate', selectId: 'exportColumnHireDate' },
    { key: 'country', selectId: 'exportColumnCountry' },
    { key: 'state', selectId: 'exportColumnState' },
    { key: 'city', selectId: 'exportColumnCity' },
    { key: 'office', selectId: 'exportColumnOffice' },
    { key: 'manager', selectId: 'exportColumnManager' }
];

const EXPORT_COLUMN_DEFAULTS = {
    name: 'show',
    title: 'show',
    department: 'show',
    email: 'show',
    phone: 'show',
    hireDate: 'admin',
    country: 'show',
    state: 'show',
    city: 'show',
    office: 'show',
    manager: 'show'
};

async function loadSettings() {
    try {
        const response = await fetch(`${API_BASE_URL}/api/settings`);
        if (response.ok) {
            currentSettings = await response.json();
            applySettings(currentSettings);
        }
    } catch (error) {
        console.error('Error loading settings:', error);
    }
}

function initTimePicker() {
    const hourSel = document.getElementById('updateHour');
    const minSel = document.getElementById('updateMinute');
    if (!hourSel || !minSel) return;
    if (!hourSel.options.length) {
        for (let h = 0; h < 24; h++) {
            const opt = document.createElement('option');
            opt.value = opt.textContent = h.toString().padStart(2, '0');
            hourSel.appendChild(opt);
        }
    }
    if (!minSel.options.length) {
        for (let m = 0; m < 60; m++) {
            const opt = document.createElement('option');
            opt.value = opt.textContent = m.toString().padStart(2, '0');
            minSel.appendChild(opt);
        }
    }
    function syncHidden() {
        const hidden = document.getElementById('updateTime');
        if (hidden) hidden.value = `${hourSel.value}:${minSel.value}`;
    }
    hourSel.addEventListener('change', syncHidden);
    minSel.addEventListener('change', syncHidden);
    if (!hourSel.value) hourSel.value = '20';
    if (!minSel.value) minSel.value = '00';
    syncHidden();
}

function applySettings(settings) {
    document.title = `Configuration - ${settings.chartTitle || 'DB AutoOrgChart'}`;
    if (settings.chartTitle) {
        document.getElementById('chartTitle').value = settings.chartTitle;
    }

    if (settings.headerColor) {
        document.getElementById('headerColor').value = settings.headerColor;
        document.getElementById('headerColorHex').value = settings.headerColor;
        updateHeaderPreview(settings.headerColor);
    }

    const logoPath = settings.logoPath || '/static/icon.png';
    document.getElementById('currentLogo').src = `${logoPath}?t=${Date.now()}`;
    const logoStatus = document.getElementById('logoStatus');
    if (logoStatus) {
        logoStatus.textContent = logoPath.includes('icon_custom_') ? 'Custom uploaded' : 'Using default';
    }

    const faviconPath = settings.faviconPath || '/favicon.ico';
    document.getElementById('currentFavicon').src = `${faviconPath}?t=${Date.now()}`;
    const favStatus = document.getElementById('faviconStatus');
    if (favStatus) {
        favStatus.textContent = faviconPath.includes('favicon_custom_') ? 'Custom uploaded' : 'Using default';
    }

    if (settings.nodeColors) {
        ['level0', 'level1', 'level2', 'level3', 'level4', 'level5'].forEach(level => {
            if (settings.nodeColors[level]) {
                const colorInput = document.getElementById(`${level}Color`);
                const hexInput = document.getElementById(`${level}ColorHex`);
                if (colorInput && hexInput) {
                    colorInput.value = settings.nodeColors[level];
                    hexInput.value = settings.nodeColors[level];
                }
            }
        });
    }

    if (settings.autoUpdateEnabled !== undefined) {
        document.getElementById('autoUpdateEnabled').checked = settings.autoUpdateEnabled;
    }

    if (settings.updateTime) {
        const [h, m] = settings.updateTime.split(':');
        const hourSel = document.getElementById('updateHour');
        const minSel = document.getElementById('updateMinute');
        if (hourSel && minSel) {
            hourSel.value = h.padStart(2, '0');
            minSel.value = m.padStart(2, '0');
        }
        const hidden = document.getElementById('updateTime');
        if (hidden) hidden.value = `${h.padStart(2, '0')}:${m.padStart(2, '0')}`;
    }

    if (settings.collapseLevel) {
        document.getElementById('collapseLevel').value = settings.collapseLevel;
    }

    if (settings.searchAutoExpand !== undefined) {
        document.getElementById('searchAutoExpand').checked = settings.searchAutoExpand;
    }

    if (settings.searchHighlight !== undefined) {
        document.getElementById('searchHighlight').checked = settings.searchHighlight;
    }

    if (settings.showDepartments !== undefined) {
        document.getElementById('showDepartments').checked = settings.showDepartments;
    }

    if (settings.showEmployeeCount !== undefined) {
        document.getElementById('showEmployeeCount').checked = settings.showEmployeeCount;
    }

    if (settings.highlightNewEmployees !== undefined) {
        document.getElementById('highlightNewEmployees').checked = settings.highlightNewEmployees;
    }

    if (settings.newEmployeeMonths !== undefined) {
        document.getElementById('newEmployeeMonths').value = settings.newEmployeeMonths;
    }

    if (settings.hideDisabledUsers !== undefined) {
        document.getElementById('hideDisabledUsers').checked = settings.hideDisabledUsers;
    }

    if (settings.hideGuestUsers !== undefined) {
        document.getElementById('hideGuestUsers').checked = settings.hideGuestUsers;
    }

    if (settings.hideNoTitle !== undefined) {
        document.getElementById('hideNoTitle').checked = settings.hideNoTitle;
    }

    if (settings.ignoredDepartments !== undefined) {
        const el = document.getElementById('ignoredDepartmentsInput');
        if (el) el.value = settings.ignoredDepartments;
    }
    if (settings.ignoredTitles !== undefined) {
        const el = document.getElementById('ignoredTitlesInput');
        if (el) el.value = settings.ignoredTitles;
    }

    if (settings.printOrientation) {
        document.getElementById('printOrientation').value = settings.printOrientation;
    }

    if (settings.printSize) {
        document.getElementById('printSize').value = settings.printSize;
    }

    if (settings.multiLineChildrenThreshold !== undefined) {
        const el2 = document.getElementById('multiLineChildrenThreshold');
        if (el2) el2.value = settings.multiLineChildrenThreshold;
    }

    applyExportColumnSettings(settings);
}

function applyExportColumnSettings(settings) {
    const modeFor = (key) => {
        const value = (settings.exportXlsxColumns || {})[key];
        if (value === 'hide' || value === 'admin') {
            return value;
        }
        return 'show';
    };

    EXPORT_COLUMN_KEYS.forEach(({ key, selectId }) => {
        const select = document.getElementById(selectId);
        if (select) {
            select.value = modeFor(key);
        }
    });
}

function getExportColumnSettings() {
    const result = {};
    EXPORT_COLUMN_KEYS.forEach(({ key, selectId }) => {
        const select = document.getElementById(selectId);
        if (select) {
            const value = select.value;
            result[key] = (value === 'hide' || value === 'admin') ? value : 'show';
        }
    });
    return result;
}

function resetExportColumns() {
    EXPORT_COLUMN_KEYS.forEach(({ key, selectId }) => {
        const select = document.getElementById(selectId);
        if (select) {
            select.value = EXPORT_COLUMN_DEFAULTS[key] || 'show';
        }
    });
}

document.getElementById('headerColor').addEventListener('input', event => {
    document.getElementById('headerColorHex').value = event.target.value;
    updateHeaderPreview(event.target.value);
});

document.getElementById('headerColorHex').addEventListener('input', event => {
    if (event.target.value.match(/^#[0-9A-Fa-f]{6}$/)) {
        document.getElementById('headerColor').value = event.target.value;
        updateHeaderPreview(event.target.value);
    }
});

['level0', 'level1', 'level2', 'level3', 'level4', 'level5'].forEach(level => {
    const colorInput = document.getElementById(`${level}Color`);
    const hexInput = document.getElementById(`${level}ColorHex`);

    if (!colorInput || !hexInput) {
        return;
    }

    colorInput.addEventListener('input', event => {
        hexInput.value = event.target.value;
    });

    hexInput.addEventListener('input', event => {
        if (event.target.value.match(/^#[0-9A-Fa-f]{6}$/)) {
            colorInput.value = event.target.value;
        }
    });
});

function updateHeaderPreview(color) {
    const darker = adjustColor(color, -30);
    const stylesheet = Array.from(document.styleSheets).find(sheet => {
        try {
            return sheet.href && sheet.href.includes('configureme.css');
        } catch (error) {
            return false;
        }
    });

    if (!stylesheet) {
        return;
    }

    let rootRule = Array.from(stylesheet.cssRules).find(rule => rule.selectorText === ':root');
    if (!rootRule) {
        const index = stylesheet.cssRules.length;
        stylesheet.insertRule(':root {}', index);
        rootRule = stylesheet.cssRules[index];
    }

    rootRule.style.setProperty('--header-preview-primary', color);
    rootRule.style.setProperty('--header-preview-secondary', darker);
}

function adjustColor(color, amount) {
    const num = parseInt(color.replace('#', ''), 16);
    const r = Math.max(0, Math.min(255, (num >> 16) + amount));
    const g = Math.max(0, Math.min(255, ((num >> 8) & 0x00ff) + amount));
    const b = Math.max(0, Math.min(255, (num & 0x0000ff) + amount));
    return `#${((r << 16) | (g << 8) | b).toString(16).padStart(6, '0')}`;
}

initTimePicker();
loadSettings();

async function uploadLogoFile(file) {
    if (!file) return;
    const maxSize = 5 * 1024 * 1024;
    if (file.size > maxSize) {
        showStatus('File too large. Maximum size is 5MB.', 'error');
        return;
    }
    const allowedTypes = ['image/png', 'image/jpeg', 'image/jpg'];
    if (!allowedTypes.includes(file.type)) {
        showStatus('Invalid file type. Please use PNG, JPG, or JPEG.', 'error');
        return;
    }
    showStatus('Uploading logo...', 'info');
    try {
        const authCheck = await fetch(`${API_BASE_URL}/api/settings`);
        if (!authCheck.ok && authCheck.status === 401) {
            showStatus('Session expired. Please refresh the page and log in again.', 'error');
            return;
        }
    } catch {
        // ignore auth check errors
    }
    const formData = new FormData();
    formData.append('logo', file);
    try {
        const response = await fetch(`${API_BASE_URL}/api/upload-logo`, { method: 'POST', body: formData });
        if (response.ok) {
            const result = await response.json();
            const logoImg = document.getElementById('currentLogo');
            const newLogoPath = `${result.path}?t=${Date.now()}`;
            logoImg.onload = () => {
                showStatus('Logo uploaded successfully!', 'success');
                const status = document.getElementById('logoStatus');
                if (status) status.textContent = 'Custom uploaded';
            };
            logoImg.onerror = () => {
                console.error('Logo failed to load from path:', newLogoPath);
                showStatus('Logo uploaded but failed to display. Try refreshing the page.', 'warning');
            };
            logoImg.src = newLogoPath;
        } else {
            let errorMessage = 'Unknown error';
            try {
                const errorData = await response.json();
                errorMessage = errorData.error || errorData.message || 'Server error';
            } catch {
                const errorText = await response.text();
                errorMessage = errorText || `HTTP ${response.status} error`;
            }
            if (response.status === 401) showStatus('Authentication required. Please log in again.', 'error');
            else if (response.status === 413) showStatus('File too large. Please use a smaller image.', 'error');
            else if (response.status === 429) showStatus('Too many upload attempts. Please wait a moment.', 'error');
            else showStatus(errorMessage, 'error');
        }
    } catch (error) {
        console.error('Upload error:', error);
        showStatus('Network error uploading logo. Please try again.', 'error');
    }
}

document.getElementById('logoUpload').addEventListener('change', async event => {
    const file = event.target.files[0];
    await uploadLogoFile(file);
    event.target.value = '';
});

async function uploadFaviconFile(file) {
    if (!file) return;
    const maxSize = 5 * 1024 * 1024;
    if (file.size > maxSize) {
        showStatus('File too large. Maximum size is 5MB.', 'error');
        return;
    }
    const allowedTypes = ['image/x-icon', 'image/png', 'image/jpeg', 'image/jpg'];
    if (!allowedTypes.includes(file.type) && !file.name.toLowerCase().endsWith('.ico')) {
        showStatus('Invalid file type. Please use ICO, PNG, JPG, or JPEG.', 'error');
        return;
    }
    showStatus('Uploading favicon...', 'info');
    try {
        const formData = new FormData();
        formData.append('favicon', file);
        const response = await fetch(`${API_BASE_URL}/api/upload-favicon`, { method: 'POST', body: formData });
        if (response.ok) {
            const result = await response.json();
            const faviconImg = document.getElementById('currentFavicon');
            const newFaviconPath = `${result.path}?t=${Date.now()}`;
            faviconImg.onload = () => {
                showStatus('Favicon uploaded successfully!', 'success');
                updatePageFavicon(newFaviconPath);
                const status = document.getElementById('faviconStatus');
                if (status) status.textContent = 'Custom uploaded';
            };
            faviconImg.onerror = () => {
                showStatus('Favicon uploaded but failed to display.', 'warning');
            };
            faviconImg.src = newFaviconPath;
        } else {
            let errorMessage = 'Unknown error';
            try {
                const errorData = await response.json();
                errorMessage = errorData.error || errorData.message || 'Server error';
            } catch {
                const errorText = await response.text();
                errorMessage = errorText || `HTTP ${response.status} error`;
            }
            showStatus(errorMessage, 'error');
        }
    } catch (error) {
        console.error('Favicon upload error:', error);
        showStatus('Network error uploading favicon. Please try again.', 'error');
    }
}

document.getElementById('faviconUpload').addEventListener('change', async event => {
    const file = event.target.files[0];
    await uploadFaviconFile(file);
    event.target.value = '';
});

function updatePageFavicon(faviconPath) {
    let favicon = document.querySelector('link[rel="icon"]') || document.querySelector('link[rel="shortcut icon"]');
    if (!favicon) {
        favicon = document.createElement('link');
        favicon.rel = 'icon';
        document.head.appendChild(favicon);
    }
    favicon.href = faviconPath;
}

function resetFavicon() {
    if (confirm('Are you sure you want to reset the favicon to default?')) {
        fetch(`${API_BASE_URL}/api/reset-favicon`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' }
        })
            .then(response => response.json())
            .then(result => {
                if (result.success) {
                    document.getElementById('currentFavicon').src = `/favicon.ico?t=${Date.now()}`;
                    updatePageFavicon('/favicon.ico');
                    showStatus('Favicon reset to default successfully!', 'success');
                    const status = document.getElementById('faviconStatus');
                    if (status) status.textContent = 'Using default';
                } else {
                    showStatus(`Failed to reset favicon: ${result.error || 'Unknown error'}`, 'error');
                }
            })
            .catch(error => {
                console.error('Error resetting favicon:', error);
                showStatus('Error resetting favicon. Please try again.', 'error');
            });
    }
}

function resetChartTitle() {
    document.getElementById('chartTitle').value = 'Organization Chart';
}

function resetHeaderColor() {
    document.getElementById('headerColor').value = '#0078d4';
    document.getElementById('headerColorHex').value = '#0078d4';
    updateHeaderPreview('#0078d4');
}

function resetLogo() {
    document.getElementById('currentLogo').src = `/static/icon.png?t=${Date.now()}`;
    const status = document.getElementById('logoStatus');
    if (status) status.textContent = 'Using default';
    fetch(`${API_BASE_URL}/api/reset-logo`, { method: 'POST' });
}

function resetNodeColors() {
    const defaults = {
        level0: '#90EE90',
        level1: '#FFFFE0',
        level2: '#E0F2FF',
        level3: '#FFE4E1',
        level4: '#E8DFF5',
        level5: '#FFEAA7'
    };

    Object.keys(defaults).forEach(level => {
        document.getElementById(`${level}Color`).value = defaults[level];
        document.getElementById(`${level}ColorHex`).value = defaults[level];
    });
}

function resetUpdateTime() {
    const hourSel = document.getElementById('updateHour');
    const minSel = document.getElementById('updateMinute');
    if (hourSel) hourSel.value = '20';
    if (minSel) minSel.value = '00';
    const hidden = document.getElementById('updateTime');
    if (hidden) hidden.value = '20:00';
    document.getElementById('autoUpdateEnabled').checked = true;
}

function resetCollapseLevel() {
    document.getElementById('collapseLevel').value = '2';
}

function resetMultiLineSettings() {
    const thresholdEl = document.getElementById('multiLineChildrenThreshold');
    if (thresholdEl) thresholdEl.value = 20;
}

function resetIgnoredDepartments() {
    const el = document.getElementById('ignoredDepartmentsInput');
    if (el) el.value = '';
}

function resetIgnoredTitles() {
    const el = document.getElementById('ignoredTitlesInput');
    if (el) el.value = '';
}

async function resetAllSettings() {
    if (confirm('Are you sure you want to reset all settings to defaults? This cannot be undone.')) {
        document.getElementById('chartTitle').value = 'Organization Chart';
        resetHeaderColor();
        resetLogo();
        resetNodeColors();
        resetUpdateTime();
        resetCollapseLevel();
        document.getElementById('searchAutoExpand').checked = true;
        document.getElementById('searchHighlight').checked = true;
        document.getElementById('showDepartments').checked = true;
        document.getElementById('showEmployeeCount').checked = true;
        document.getElementById('highlightNewEmployees').checked = true;
        document.getElementById('newEmployeeMonths').value = '3';
        document.getElementById('hideDisabledUsers').checked = true;
        document.getElementById('hideGuestUsers').checked = true;
        document.getElementById('hideNoTitle').checked = true;
        document.getElementById('ignoredDepartmentsInput').value = '';
        const titlesInput = document.getElementById('ignoredTitlesInput');
        if (titlesInput) titlesInput.value = '';
        document.getElementById('printOrientation').value = 'landscape';
        document.getElementById('printSize').value = 'a4';
        const mlThreshold = document.getElementById('multiLineChildrenThreshold');
        if (mlThreshold) mlThreshold.value = 20;
        resetExportColumns();

        try {
            const response = await fetch(`${API_BASE_URL}/api/reset-all-settings`, { method: 'POST' });
            if (response.ok) {
                showStatus('All settings reset to defaults!', 'success');
                setTimeout(() => location.reload(), 1500);
            }
        } catch (error) {
            showStatus('Error resetting settings', 'error');
        }
    }
}

async function logout() {
    try {
        const response = await fetch(`${API_BASE_URL}/logout`, {
            method: 'POST',
            credentials: 'include'
        });

        if (response.ok) {
            window.location.href = '/';
        } else {
            sessionStorage.clear();
            localStorage.clear();
            window.location.href = '/';
        }
    } catch (error) {
        console.error('Logout error:', error);
        window.location.href = '/';
    }
}

async function saveAllSettings() {
    const settings = {
        chartTitle: document.getElementById('chartTitle').value || 'Organization Chart',
        headerColor: document.getElementById('headerColor').value,
        nodeColors: {
            level0: document.getElementById('level0Color').value,
            level1: document.getElementById('level1Color').value,
            level2: document.getElementById('level2Color').value,
            level3: document.getElementById('level3Color').value,
            level4: document.getElementById('level4Color').value,
            level5: document.getElementById('level5Color').value
        },
        autoUpdateEnabled: document.getElementById('autoUpdateEnabled').checked,
        updateTime: document.getElementById('updateTime').value,
        collapseLevel: document.getElementById('collapseLevel').value,
        searchAutoExpand: document.getElementById('searchAutoExpand').checked,
        searchHighlight: document.getElementById('searchHighlight').checked,
        showDepartments: document.getElementById('showDepartments').checked,
        showEmployeeCount: document.getElementById('showEmployeeCount').checked,
        highlightNewEmployees: document.getElementById('highlightNewEmployees').checked,
        newEmployeeMonths: parseInt(document.getElementById('newEmployeeMonths').value, 10),
        hideDisabledUsers: document.getElementById('hideDisabledUsers').checked,
        hideGuestUsers: document.getElementById('hideGuestUsers').checked,
        hideNoTitle: document.getElementById('hideNoTitle').checked,
        ignoredDepartments: (document.getElementById('ignoredDepartmentsInput')?.value || '').trim(),
        ignoredTitles: (document.getElementById('ignoredTitlesInput')?.value || '').trim(),
        printOrientation: document.getElementById('printOrientation').value,
        printSize: document.getElementById('printSize').value,
        multiLineChildrenThreshold: parseInt(document.getElementById('multiLineChildrenThreshold')?.value || '20', 10),
        exportXlsxColumns: getExportColumnSettings()
    };

    try {
        const response = await fetch(`${API_BASE_URL}/api/settings`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(settings)
        });

        if (response.ok) {
            showStatus('Settings saved successfully!', 'success');
        } else {
            const errorData = await response.json().catch(() => ({}));
            const errorMsg = errorData.error || `Error saving settings (Status: ${response.status})`;
            console.error('Settings save error:', errorData);
            showStatus(errorMsg, 'error');
        }
    } catch (error) {
        console.error('Network error saving settings:', error);
        showStatus(`Error saving settings: ${error.message}`, 'error');
    }
}

async function triggerUpdate() {
    const statusEl = document.getElementById('updateStatus');
    statusEl.textContent = 'Updating...';

    try {
        const response = await fetch(`${API_BASE_URL}/api/update-now`, { method: 'POST' });

        if (response.ok) {
            statusEl.textContent = '✔ Update started';
            setTimeout(() => {
                statusEl.textContent = '';
            }, 3000);
        } else {
            statusEl.textContent = '✗ Update failed';
        }
    } catch (error) {
        statusEl.textContent = '✗ Update failed';
    }
}

function showStatus(message, type) {
    const statusEl = document.getElementById('statusMessage');
    statusEl.textContent = message;
    statusEl.className = `status-message ${type}`;
    setTimeout(() => {
        statusEl.className = 'status-message';
    }, 3000);
}

function registerConfigActions() {
    const actionHandlers = {
        'reset-chart-title': resetChartTitle,
        'reset-header-color': resetHeaderColor,
        'reset-logo': resetLogo,
        'reset-favicon': resetFavicon,
        'reset-node-colors': resetNodeColors,
        'reset-update-time': resetUpdateTime,
        'reset-collapse-level': resetCollapseLevel,
        'reset-ignored-titles': resetIgnoredTitles,
        'reset-ignored-departments': resetIgnoredDepartments,
        'reset-multiline-settings': resetMultiLineSettings,
        'reset-export-columns': resetExportColumns,
        'trigger-update': triggerUpdate,
        'save-all': saveAllSettings,
        'reset-all': resetAllSettings,
        'logout': logout
    };

    document.querySelectorAll('[data-config-action]').forEach(button => {
        const handler = actionHandlers[button.dataset.configAction];
        if (typeof handler === 'function') {
            button.addEventListener('click', handler);
        }
    });
}

document.addEventListener('DOMContentLoaded', () => {
    loadSettings();
    registerConfigActions();

    const logoDropZone = document.getElementById('logoDropZone');
    if (logoDropZone) {
        ['dragenter', 'dragover'].forEach(evt => logoDropZone.addEventListener(evt, event => {
            event.preventDefault();
            event.stopPropagation();
            logoDropZone.classList.add('dragover');
        }));
        ['dragleave', 'drop'].forEach(evt => logoDropZone.addEventListener(evt, event => {
            event.preventDefault();
            event.stopPropagation();
            logoDropZone.classList.remove('dragover');
        }));
        logoDropZone.addEventListener('drop', event => {
            const file = event.dataTransfer.files && event.dataTransfer.files[0];
            if (file) uploadLogoFile(file);
        });
    }

    const faviconDropZone = document.getElementById('faviconDropZone');
    if (faviconDropZone) {
        ['dragenter', 'dragover'].forEach(evt => faviconDropZone.addEventListener(evt, event => {
            event.preventDefault();
            event.stopPropagation();
            faviconDropZone.classList.add('dragover');
        }));
        ['dragleave', 'drop'].forEach(evt => faviconDropZone.addEventListener(evt, event => {
            event.preventDefault();
            event.stopPropagation();
            faviconDropZone.classList.remove('dragover');
        }));
        faviconDropZone.addEventListener('drop', event => {
            const file = event.dataTransfer.files && event.dataTransfer.files[0];
            if (file) uploadFaviconFile(file);
        });
    }
});
