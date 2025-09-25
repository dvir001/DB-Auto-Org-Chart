let currentData = null;
let allEmployees = [];
const employeeById = new Map();
let root = null;
let svg = null;
let g = null;
let linkLayer = null;
let nodeLayer = null;
let zoom = null;
let appSettings = {};
let currentLayout = 'vertical'; // Default layout
const hiddenNodeIds = new Set(JSON.parse(localStorage.getItem('hiddenNodeIds') || '[]'));
let isAuthenticated = false;
const COMPACT_PREFERENCE_KEY = 'orgChart.compactLargeTeams';
let userCompactPreference = null;
const PROFILE_IMAGE_PREFERENCE_KEY = 'orgChart.showProfileImages';
let userProfileImagesPreference = null;
let serverShowProfileImages = null;

function loadStoredCompactPreference() {
    userCompactPreference = null;
    try {
        const stored = localStorage.getItem(COMPACT_PREFERENCE_KEY);
        if (stored === 'true') {
            userCompactPreference = true;
        } else if (stored === 'false') {
            userCompactPreference = false;
        }
    } catch (error) {
        console.warn('Unable to access compact layout preference storage:', error);
        userCompactPreference = null;
    }
}

function storeCompactPreference(value) {
    userCompactPreference = value;
    try {
        localStorage.setItem(COMPACT_PREFERENCE_KEY, String(value));
    } catch (error) {
        console.warn('Unable to persist compact layout preference:', error);
    }
}

function clearCompactPreferenceStorage() {
    userCompactPreference = null;
    try {
        localStorage.removeItem(COMPACT_PREFERENCE_KEY);
    } catch (error) {
        console.warn('Unable to clear compact layout preference storage:', error);
    }
}

function getEffectiveCompactEnabled() {
    const serverEnabled = !appSettings || appSettings.multiLineChildrenEnabled !== false;
    if (!isAuthenticated && userCompactPreference !== null) {
        return userCompactPreference;
    }
    return serverEnabled;
}

function loadStoredProfileImagePreference() {
    userProfileImagesPreference = null;
    try {
        const stored = localStorage.getItem(PROFILE_IMAGE_PREFERENCE_KEY);
        if (stored === 'true') {
            userProfileImagesPreference = true;
        } else if (stored === 'false') {
            userProfileImagesPreference = false;
        }
    } catch (error) {
        console.warn('Unable to access profile image preference storage:', error);
        userProfileImagesPreference = null;
    }
}

function storeProfileImagePreference(value) {
    userProfileImagesPreference = value;
    try {
        localStorage.setItem(PROFILE_IMAGE_PREFERENCE_KEY, String(value));
    } catch (error) {
        console.warn('Unable to persist profile image preference:', error);
    }
}

function clearProfileImagePreference() {
    userProfileImagesPreference = null;
    try {
        localStorage.removeItem(PROFILE_IMAGE_PREFERENCE_KEY);
    } catch (error) {
        console.warn('Unable to clear profile image preference storage:', error);
    }
}

function getEffectiveProfileImagesEnabled() {
    const serverEnabled = (serverShowProfileImages != null)
        ? serverShowProfileImages
        : (!appSettings || appSettings.showProfileImages !== false);
    if (userProfileImagesPreference !== null) {
        return userProfileImagesPreference;
    }
    return serverEnabled;
}

function persistHiddenIds() {
    localStorage.setItem('hiddenNodeIds', JSON.stringify(Array.from(hiddenNodeIds)));
}

function isHiddenNode(node) {
    // A node is hidden if itself or any ancestor is marked
    let cur = node;
    while (cur) {
        if (hiddenNodeIds.has(cur.data.id)) return true;
        cur = cur.parent;
    }
    return false;
}

function toggleHideNode(d) {
    if (!d || !d.data || !d.data.id) return;
    if (hiddenNodeIds.has(d.data.id)) {
        hiddenNodeIds.delete(d.data.id);
    } else {
        hiddenNodeIds.add(d.data.id);
    }
    persistHiddenIds();
    update(d);
}

function resetHiddenSubtrees() {
    if (hiddenNodeIds.size === 0) return;
    hiddenNodeIds.clear();
    persistHiddenIds();
    update(root);
}

const API_BASE_URL = window.location.origin;
const nodeWidth = 220;
const nodeHeight = 80;
const levelHeight = 130; // slightly taller to afford multi-line rows

// Zoom tracking and helpers
let userAdjustedZoom = false;
let programmaticZoomActive = false;
let resizeTimer = null;
const RESIZE_DEBOUNCE_MS = 180;

function applyZoomTransform(transform, { duration = 750, resetUser = false } = {}) {
    if (!svg || !zoom) return;
    programmaticZoomActive = true;

    const finalize = () => {
        programmaticZoomActive = false;
        if (resetUser) {
            userAdjustedZoom = false;
        }
    };

    if (duration > 0) {
        const transition = svg.transition().duration(duration).call(zoom.transform, transform);
        transition.on('end', finalize);
        transition.on('interrupt', finalize);
    } else {
        svg.call(zoom.transform, transform);
        finalize();
    }
}

function updateSvgSize() {
    if (!svg) return;
    const container = document.getElementById('orgChart');
    if (!container) return;
    const width = container.clientWidth;
    const height = container.clientHeight || 800;
    svg.attr('width', width).attr('height', height);
}

function createTreeLayout() {
    const layout = d3.tree()
        .nodeSize(currentLayout === 'vertical'
            ? [nodeWidth + 26, levelHeight]
            : [levelHeight, nodeWidth + 26])
        .separation((a, b) => {
            const sameParent = a.parent && b.parent && a.parent === b.parent;
            // Keep siblings at full spacing to avoid overlap even for large teams
            return sameParent ? 1.0 : 1.2;
        });
    return layout;
}
const userIconUrl = window.location.origin + '/static/usericon.png';


// Security: HTML escaping function to prevent XSS
function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function renderAvatar({ imageUrl, name, initials, imageClass, fallbackClass }) {
    const safeInitials = escapeHtml(initials || '');
    if (imageUrl && appSettings.showProfileImages !== false) {
        const safeUrl = escapeHtml(imageUrl);
        return `
            <img class="${imageClass}" src="${safeUrl}" alt="${escapeHtml(name || '')}" data-role="avatar-image">
            <div class="${fallbackClass}" data-role="avatar-fallback" hidden>${safeInitials}</div>
        `;
    }
    return `<div class="${fallbackClass}" data-role="avatar-fallback">${safeInitials}</div>`;
}

// Date formatting function to format hire dates to yyyy-MM-dd
function formatHireDate(dateString) {
    if (!dateString) return '';
    try {
        const date = new Date(dateString);
        if (isNaN(date.getTime())) return dateString; // Return original if invalid
        return date.toISOString().split('T')[0]; // Returns yyyy-MM-dd format
    } catch (e) {
        return dateString; // Return original if parsing fails
    }
}

// Dynamic font scaling based on text length
function calculateFontSize(text, baseSize, maxLength, minSize = 9) {
    if (!text) return baseSize;
    const length = text.length;
    if (length <= maxLength * 0.7) return baseSize; // Normal size for short text
    if (length <= maxLength) return Math.max(baseSize * 0.9, minSize); // Slightly smaller for medium text
    return Math.max(baseSize * 0.75, minSize); // Smaller for long text
}

function getLabelOffsetX() {
    return appSettings.showProfileImages !== false ? -nodeWidth / 2 + 50 : 0;
}

function getLabelAnchor() {
    return appSettings.showProfileImages !== false ? 'start' : 'middle';
}

function getNameFontSizePx(name) {
    const maxLength = appSettings.showProfileImages !== false ? 25 : 30;
    return calculateFontSize(name, 14, maxLength) + 'px';
}

function getTitleFontSizePx(title) {
    const maxLength = appSettings.showProfileImages !== false ? 25 : 30;
    return calculateFontSize(title, 11, maxLength, 8) + 'px';
}

function getDepartmentFontSizePx(dept) {
    const maxLength = appSettings.showProfileImages !== false ? 25 : 35;
    return calculateFontSize(dept, 9, maxLength, 7) + 'px';
}

function getTrimmedTitle(title = '') {
    const charLimit = appSettings.showProfileImages !== false ? 45 : 50;
    return title.length > charLimit ? title.substring(0, charLimit) + '...' : title;
}

// Security: Safely set innerHTML with escaped content
function safeInnerHTML(element, htmlContent) {
    element.innerHTML = htmlContent;
}

function applyProfileImageAttributes(selection) {
    selection
        .attr('class', 'profile-image')
        .attr('xlink:href', userIconUrl)
        .attr('x', -nodeWidth / 2 + 8)
        .attr('y', -18)
        .attr('width', 36)
        .attr('height', 36)
        .attr('clip-path', 'circle(18px at 18px 18px)')
        .attr('preserveAspectRatio', 'xMidYMid slice')
        .each(function(d) {
            if (d.data.photoUrl && d.data.photoUrl.includes('/api/photo/')) {
                const element = d3.select(this);
                const img = new Image();

                img.onload = function() {
                    element.attr('xlink:href', d.data.photoUrl);
                    console.log(`Photo loaded for ${d.data.name}`);
                };

                img.onerror = function() {
                    console.log(`Photo failed for ${d.data.name}, keeping default icon`);
                };

                img.src = d.data.photoUrl;
            }
        });
}

async function loadSettings() {
    try {
        const response = await fetch(`${API_BASE_URL}/api/settings`);
        if (response.ok) {
            appSettings = await response.json();
            serverShowProfileImages = appSettings.showProfileImages !== false;
            applySettings();
        } else {
            // If settings fail to load, still show header content with defaults
            showHeaderContent();
        }
    } catch (error) {
        console.error('Error loading settings:', error);
        // If settings fail to load, still show header content with defaults
        showHeaderContent();
    }
}

function showHeaderContent() {
    // Show header content even if settings failed to load
    const headerContent = document.querySelector('.header-content');
    if (headerContent) {
        headerContent.classList.remove('loading');
    }
    
    // Also show default logo if settings failed
    const logo = document.querySelector('.header-logo');
    ensureLogoFallback(logo);
    if (logo && !logo.src) {
        logo.src = logo.dataset.defaultSrc || '/static/icon.png';
        logo.classList.remove('loading');
        logo.style.display = '';
    }
}

function ensureLogoFallback(logo) {
    if (!logo || logo.dataset.fallbackBound === 'true') {
        return;
    }

    const handleLoad = () => {
        logo.style.display = '';
    };

    const handleError = () => {
        logo.style.display = 'none';
    };

    logo.addEventListener('load', handleLoad);
    logo.addEventListener('error', handleError);
    logo.dataset.fallbackBound = 'true';
}

function setupStaticEventListeners() {
    const configBtn = document.getElementById('configBtn');
    if (configBtn) {
        configBtn.addEventListener('click', () => {
            window.location.href = '/configure';
        });
    }

    const logoutBtn = document.getElementById('logoutBtn');
    if (logoutBtn) {
        logoutBtn.addEventListener('click', () => {
            logout();
        });
    }

    document.querySelectorAll('[data-layout]').forEach(button => {
        button.addEventListener('click', () => {
            if (button.dataset.layout) {
                setLayoutOrientation(button.dataset.layout);
            }
        });
    });

    const controls = document.querySelector('.controls');
    if (controls) {
        controls.addEventListener('click', event => {
            const button = event.target.closest('[data-control]');
            if (!button) return;
            handleControlAction(button.dataset.control);
        });
    }

    const saveTopUserBtn = document.getElementById('saveTopUserBtn');
    if (saveTopUserBtn) {
        saveTopUserBtn.addEventListener('click', saveTopUser);
    }

    const resetTopUserBtn = document.getElementById('resetTopUserBtn');
    if (resetTopUserBtn) {
        resetTopUserBtn.addEventListener('click', resetTopUser);
    }

    const compactBtn = document.getElementById('compactToggleBtn');
    if (compactBtn) {
        compactBtn.addEventListener('click', toggleCompactLargeTeams);
    }

    const profileBtn = document.getElementById('profileImageToggleBtn');
    if (profileBtn) {
        profileBtn.addEventListener('click', toggleProfileImages);
    }

    const closeDetailBtn = document.getElementById('employeeDetailCloseBtn');
    if (closeDetailBtn) {
        closeDetailBtn.addEventListener('click', closeEmployeeDetail);
    }

    const resultsContainer = document.getElementById('searchResults');
    if (resultsContainer) {
        resultsContainer.addEventListener('click', event => {
            const item = event.target.closest('.search-result-item');
            if (!item) return;
            const employeeId = item.dataset.employeeId;
            if (employeeId) {
                selectSearchResult(employeeId);
            }
        });
    }

    const infoPanel = document.getElementById('employeeInfo');
    if (infoPanel) {
        infoPanel.addEventListener('click', event => {
            const target = event.target.closest('[data-employee-id]');
            if (!target) return;
            showEmployeeDetailById(target.dataset.employeeId);
        });
    }

    const logo = document.querySelector('.header-logo');
    if (logo) {
        ensureLogoFallback(logo);
    }

    setLayoutOrientation(currentLayout);
}

function handleControlAction(action) {
    switch (action) {
        case 'zoom-in':
            zoomIn();
            break;
        case 'zoom-out':
            zoomOut();
            break;
        case 'reset-zoom':
            resetZoom();
            break;
        case 'fit':
            fitToScreen();
            break;
        case 'expand':
            expandAll();
            break;
        case 'collapse':
            collapseAll();
            break;
        case 'reset-hidden':
            resetHiddenSubtrees();
            break;
        case 'print':
            printChart();
            break;
        case 'export-visible-svg':
            exportToImage('svg', false);
            break;
        case 'export-visible-png':
            exportToImage('png', false);
            break;
        case 'export-visible-pdf':
            exportToPDF(false);
            break;
        case 'export-xlsx':
            exportToXLSX();
            break;
        default:
            break;
    }
}

function applySettings() {
    if (appSettings.chartTitle) {
        document.querySelector('.header-text h1').textContent = appSettings.chartTitle;
        // Update the browser tab title to match the custom title
        document.title = appSettings.chartTitle;
    } else {
        // Fallback to default title if no custom title is set
        document.title = 'DB AutoOrgChart';
    }

    if (appSettings.headerColor) {
        const header = document.querySelector('.header');
        const darker = adjustColor(appSettings.headerColor, -30);
        header.style.background = `linear-gradient(135deg, ${appSettings.headerColor} 0%, ${darker} 100%)`;
    }

    // Handle logo loading
    const logo = document.querySelector('.header-logo');
    ensureLogoFallback(logo);
    if (appSettings.logoPath) {
        logo.src = appSettings.logoPath + '?t=' + Date.now();
    } else {
        // Use default logo if no custom logo is set
        logo.src = logo.dataset.defaultSrc || '/static/icon.png';
    }
    logo.classList.remove('loading'); // Show logo after src is set
    logo.style.display = '';

    if (appSettings.updateTime) {
        const timeText = appSettings.autoUpdateEnabled ? 
            `Updates daily @ ${appSettings.updateTime}` : 
            'Auto-update disabled';
        const headerP = document.querySelector('.header-text p');
        if (headerP) headerP.textContent = timeText;
    }
    
    // Show header content after settings are applied to prevent flash of default content
    const headerContent = document.querySelector('.header-content');
    if (headerContent) {
        headerContent.classList.remove('loading');
    }

    // Reflect Compact Teams toggle state
    try {
        const btn = document.getElementById('compactToggleBtn');
        if (btn && appSettings) {
            const enabled = getEffectiveCompactEnabled();
            appSettings.multiLineChildrenEnabled = enabled;
            btn.classList.toggle('active', enabled);

            // Update label to include threshold number
            const threshold = (appSettings.multiLineChildrenThreshold != null)
                ? appSettings.multiLineChildrenThreshold
                : 20;
            btn.innerHTML = '<span class="layout-icon">▦</span> ' +
                'Compact Teams (' + threshold + ')';
        }
    } catch (e) { /* no-op */ }

    const showProfileImages = getEffectiveProfileImagesEnabled();
    appSettings.showProfileImages = showProfileImages;
    const profileBtn = document.getElementById('profileImageToggleBtn');
    if (profileBtn) {
        profileBtn.classList.toggle('active', showProfileImages);
        profileBtn.setAttribute('aria-pressed', String(showProfileImages));
        profileBtn.title = showProfileImages ? 'Hide profile images' : 'Show profile images';
    }

    updateAuthDependentUI();
}

function adjustColor(color, amount) {
    const num = parseInt(color.replace('#', ''), 16);
    const r = Math.max(0, Math.min(255, (num >> 16) + amount));
    const g = Math.max(0, Math.min(255, ((num >> 8) & 0x00FF) + amount));
    const b = Math.max(0, Math.min(255, (num & 0x0000FF) + amount));
    return '#' + ((r << 16) | (g << 8) | b).toString(16).padStart(6, '0');
}

// No conversion needed; we display and store updateTime in 24-hour HH:MM

function setLayoutOrientation(orientation) {
    currentLayout = orientation;

    document.querySelectorAll('[data-layout]').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.layout === orientation);
    });

    if (root) {
        update(root);
        fitToScreen();
    }
}

function updateAdminActions() {
    const configBtn = document.getElementById('configBtn');
    if (configBtn) {
        configBtn.classList.remove('is-hidden');
    }

    const logoutBtn = document.getElementById('logoutBtn');
    if (logoutBtn) {
        logoutBtn.classList.toggle('is-hidden', !isAuthenticated);
    }
}

function updateAuthDependentUI() {
    updateAdminActions();

    const compactBtn = document.getElementById('compactToggleBtn');
    if (compactBtn) {
        compactBtn.disabled = false;
        compactBtn.removeAttribute('aria-disabled');
        const enabled = getEffectiveCompactEnabled();
        compactBtn.classList.toggle('active', enabled);
        compactBtn.title = isAuthenticated
            ? 'Toggle compact layout for large teams'
            : 'Toggle compact layout for your view (not saved globally)';
    }

    const profileBtn = document.getElementById('profileImageToggleBtn');
    if (profileBtn) {
        const showImages = getEffectiveProfileImagesEnabled();
        profileBtn.classList.toggle('active', showImages);
        profileBtn.setAttribute('aria-pressed', String(showImages));
        profileBtn.title = showImages ? 'Hide profile images' : 'Show profile images';
    }
}

async function checkAuthentication() {
    try {
        const response = await fetch(`${API_BASE_URL}/api/auth-check`, {
            credentials: 'same-origin'
        });
        return response.ok;
    } catch (error) {
        console.error('Authentication check failed:', error);
        return false;
    }
}

async function init() {
    isAuthenticated = await checkAuthentication();
    if (isAuthenticated) {
        userCompactPreference = null;
    } else {
        loadStoredCompactPreference();
    }
    loadStoredProfileImagePreference();
    updateAuthDependentUI();
    await loadSettings();
    
    try {
        const response = await fetch(`${API_BASE_URL}/api/employees`);
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        currentData = await response.json();
        window.currentOrgData = currentData; // Store globally for manager lookup
        
        if (currentData) {
            employeeById.clear();
            allEmployees = flattenTree(currentData);
            initializeTopUserSearch();
            preloadEmployeeImages(allEmployees);
            renderOrgChart(currentData);
        } else {
            throw new Error('No data received from server');
        }
    } catch (error) {
        console.error('Error loading employee data:', error);
        document.getElementById('orgChart').innerHTML = '<div class="loading">Error loading data. Please refresh the page.</div>';
    }
}

// Toggle Compact Teams from main page
async function toggleCompactLargeTeams() {
    const btn = document.getElementById('compactToggleBtn');
    const previousValue = getEffectiveCompactEnabled();
    const newValue = !previousValue;

    if (btn) {
        btn.classList.toggle('active', newValue);
    }

    if (!appSettings) {
        appSettings = {};
    }

    if (!isAuthenticated) {
        storeCompactPreference(newValue);
        appSettings.multiLineChildrenEnabled = newValue;
        if (root) {
            update(root);
            fitToScreen();
        }
        updateAuthDependentUI();
        return;
    }

    if (btn) {
        btn.disabled = true;
    }

    try {
        const res = await fetch(`${API_BASE_URL}/api/set-multiline-enabled`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ multiLineChildrenEnabled: newValue })
        });
        if (res.status === 401) {
            if (btn) {
                btn.classList.toggle('active', previousValue);
                btn.disabled = false;
            }
            isAuthenticated = false;
            updateAuthDependentUI();
            alert('Admin login expired. Please log in again to change settings.');
            return;
        }
        if (!res.ok) {
            throw new Error('Failed to save Compact Teams');
        }

        clearCompactPreferenceStorage();
        appSettings.multiLineChildrenEnabled = newValue;
        if (btn) {
            btn.disabled = false;
        }
        updateAuthDependentUI();
        if (root) {
            update(root);
            fitToScreen();
        }
    } catch (err) {
    console.error('Error toggling Compact Teams:', err);
        if (btn) {
            btn.classList.toggle('active', previousValue);
            btn.disabled = false;
        }
        appSettings.multiLineChildrenEnabled = previousValue;
        updateAuthDependentUI();
    }
}

function toggleProfileImages() {
    const btn = document.getElementById('profileImageToggleBtn');
    const currentValue = getEffectiveProfileImagesEnabled();
    const newValue = !currentValue;

    if (!appSettings) {
        appSettings = {};
    }

    appSettings.showProfileImages = newValue;

    if (serverShowProfileImages != null && newValue === serverShowProfileImages) {
        clearProfileImagePreference();
    } else {
        storeProfileImagePreference(newValue);
    }

    if (btn) {
        btn.classList.toggle('active', newValue);
        btn.setAttribute('aria-pressed', String(newValue));
        btn.title = newValue ? 'Hide profile images' : 'Show profile images';
    }

    if (root) {
        update(root);
    }

    updateAuthDependentUI();
}

function preloadEmployeeImages(employees) {
    // Preload employee images to improve loading performance
    if (appSettings.showProfileImages !== false) {
        employees.forEach(employee => {
            if (employee.photoUrl && employee.photoUrl.includes('/api/photo/')) {
                const img = new Image();
                img.onload = () => {
                    console.log(`Preloaded photo for ${employee.name}`);
                };
                img.onerror = () => {
                    console.log(`No photo available for ${employee.name} - will use default icon`);
                };
                // Load the photo URL without cache-busting for preload
                img.src = employee.photoUrl;
            }
        });
    }
}

function flattenTree(node, list = []) {
    if (!node) return list;
    list.push(node);
    if (node.id) {
        employeeById.set(node.id, node);
    }
    if (node.children && Array.isArray(node.children)) {
        node.children.forEach(child => flattenTree(child, list));
    }
    return list;
}

// Initialize top-level user search functionality
function initializeTopUserSearch() {
    const searchInput = document.getElementById('topUserSearch');
    const resultsContainer = document.getElementById('topUserResults');
    
    if (!searchInput || !resultsContainer) return;
    
    let selectedUser = null;
    
    // Set initial value if there's a configured top user
    if (appSettings.topUserEmail) {
        const currentUser = allEmployees.find(emp => emp.email === appSettings.topUserEmail);
        if (currentUser) {
            searchInput.value = currentUser.name;
            selectedUser = currentUser;
        }
    }
    
    // Search functionality
    searchInput.addEventListener('input', function() {
        const query = this.value.trim().toLowerCase();
        
        if (query.length < 2) {
            resultsContainer.classList.remove('active');
            selectedUser = null;
            return;
        }
        
        const matches = allEmployees.filter(employee => {
            if (!employee.name || !employee.email) return false;
            
            const name = employee.name.toLowerCase();
            const title = (employee.title || '').toLowerCase();
            const department = (employee.department || '').toLowerCase();
            
            return name.includes(query) || title.includes(query) || department.includes(query);
        }).slice(0, 10); // Limit to 10 results
        
        displayTopUserResults(matches, resultsContainer, searchInput);
    });
    
    // Handle clicking outside to close results
    document.addEventListener('click', function(e) {
        if (!searchInput.contains(e.target) && !resultsContainer.contains(e.target)) {
            resultsContainer.classList.remove('active');
        }
    });
    
    // Handle escape key
    searchInput.addEventListener('keydown', function(e) {
        if (e.key === 'Escape') {
            resultsContainer.classList.remove('active');
        }
    });
    
    // Store selected user reference
    searchInput._selectedUser = selectedUser;
}

function displayTopUserResults(employees, container, input) {
    container.innerHTML = '';
    
    if (employees.length === 0) {
        container.classList.remove('active');
        return;
    }
    
    employees.forEach(employee => {
        const item = document.createElement('div');
        item.className = 'search-result-item';
        item.innerHTML = `
            <div class="search-result-name">${escapeHtml(employee.name)}</div>
            <div class="search-result-title">${escapeHtml(employee.title || 'No Title')} ${employee.department ? '– ' + escapeHtml(employee.department) : ''}</div>
        `;
        
        item.addEventListener('click', function() {
            input.value = employee.name;
            input._selectedUser = employee;
            container.classList.remove('active');
        });
        
        container.appendChild(item);
    });
    
    container.classList.add('active');
}

// Save the selected top-level user
async function saveTopUser() {
    const searchInput = document.getElementById('topUserSearch');
    const saveBtn = document.getElementById('saveTopUserBtn');
    
    if (!searchInput) return;

    const selectedUser = searchInput._selectedUser;
    const inputValue = searchInput.value.trim();
    
    // If input is empty, save as auto-detect
    const emailToSave = inputValue === '' ? '' : (selectedUser ? selectedUser.email : '');
    
    // Debug logging
    console.log('SaveTopUser Debug:');
    console.log('- inputValue:', inputValue);
    console.log('- selectedUser:', selectedUser);
    console.log('- emailToSave:', emailToSave);
    
    try {
        if (saveBtn) {
            saveBtn.disabled = true;
            saveBtn.textContent = 'Saving...';
        }
        
        // Update the setting using the public endpoint
        const response = await fetch(`${API_BASE_URL}/api/set-top-user`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                topUserEmail: emailToSave
            })
        });
        
        if (response.ok) {
            // Update app settings
            appSettings.topUserEmail = emailToSave;
            
            // Show success feedback and force refresh
            if (saveBtn) {
                saveBtn.textContent = 'Saved!';
            }
            setTimeout(() => {
                // Force a full page refresh to ensure clean state and preserve toolbar
                window.location.reload();
            }, 1000);
        } else {
            throw new Error('Failed to save setting');
        }
    } catch (error) {
        console.error('Error saving top user:', error);
        if (saveBtn) {
            saveBtn.textContent = 'Error';
            setTimeout(() => {
                saveBtn.textContent = 'Save';
                saveBtn.disabled = false;
            }, 2000);
        }
    }
}

// Reset top-level user to auto-detect
async function resetTopUser() {
    const searchInput = document.getElementById('topUserSearch');
    const resultsContainer = document.getElementById('topUserResults');
    const resetBtn = document.getElementById('resetTopUserBtn');
    
    try {
        if (resetBtn) {
            resetBtn.disabled = true;
            resetBtn.textContent = 'Resetting...';
        }
        
        // Update the setting to empty string (auto-detect) using the public endpoint
        const response = await fetch(`${API_BASE_URL}/api/set-top-user`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                topUserEmail: ''
            })
        });
        
        if (response.ok) {
            // Update app settings
            appSettings.topUserEmail = '';
            
            // Clear the search input
            searchInput.value = '';
            searchInput._selectedUser = null;
            resultsContainer.classList.remove('active');
            
            // Show success feedback and force refresh
            if (resetBtn) {
                resetBtn.textContent = 'Reset!';
            }
            setTimeout(() => {
                // Force a full page refresh to ensure clean state
                window.location.reload();
            }, 1000);
        } else {
            throw new Error('Failed to reset setting');
        }
    } catch (error) {
        console.error('Error resetting top user:', error);
        if (resetBtn) {
            resetBtn.textContent = 'Error';
            setTimeout(() => {
                resetBtn.textContent = 'Reset';
                resetBtn.disabled = false;
            }, 2000);
        }
    }
}

// Reload employee data and re-render chart
async function reloadEmployeeData() {
    try {
        // Show loading state
        document.getElementById('orgChart').innerHTML = '<div class="loading"><div class="spinner"></div><p>Updating organization chart...</p></div>';
        
        const response = await fetch(`${API_BASE_URL}/api/employees`);
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        currentData = await response.json();
        window.currentOrgData = currentData; // Store globally for manager lookup
        
        if (currentData) {
            employeeById.clear();
            allEmployees = flattenTree(currentData);
            preloadEmployeeImages(allEmployees);
            renderOrgChart(currentData);
        } else {
            throw new Error('No data received from server');
        }
    } catch (error) {
        console.error('Error reloading employee data:', error);
        document.getElementById('orgChart').innerHTML = '<div class="loading">Error loading data. Please refresh the page.</div>';
    }
}

function renderOrgChart(data) {
    if (!data) {
        console.error('No data to render');
        return;
    }

    const container = document.getElementById('orgChart');
    container.querySelector('.loading').style.display = 'none';

    const width = container.clientWidth;
    const height = container.clientHeight || 800;

    svg = d3.select('#orgChart')
        .append('svg')
        .attr('width', width)
        .attr('height', height);

    updateSvgSize();

    zoom = d3.zoom()
        .scaleExtent([0.1, 3])
        .on('zoom', (event) => {
            g.attr('transform', event.transform);
            if (!programmaticZoomActive && event.sourceEvent) {
                userAdjustedZoom = true;
            }
        });

    svg.call(zoom);

    g = svg.append('g');
    linkLayer = g.append('g').attr('class', 'links');
    nodeLayer = g.append('g').attr('class', 'nodes');

    const initialTransform = d3.zoomIdentity.translate(width/2, 100);
    applyZoomTransform(initialTransform, { duration: 0, resetUser: true });

    root = d3.hierarchy(data);

    root.x0 = 0;
    root.y0 = 0;

    const treeLayout = createTreeLayout();

    const collapseLevel = appSettings.collapseLevel || '2';
    if (collapseLevel !== 'all') {
        const level = parseInt(collapseLevel);
        root.each(d => {
            if (d.depth >= level - 1 && d.children) {
                d._children = d.children;
                d.children = null;
            }
        });
    }

    update(root);
    fitToScreen({ duration: 0 });
}

function update(source) {
    const treeLayout = createTreeLayout();

    const treeData = treeLayout(root);
    const nodes = treeData.descendants();
    const links = treeData.links();

    // Swap x and y coordinates for horizontal layout
    if (currentLayout === 'horizontal') {
        nodes.forEach(d => {
            const temp = d.x;
            d.x = d.y;
            d.y = temp;
        });
    }

    // Apply multi-line wrap for large children groups (client-side layout tweak)
    applyMultiLineChildrenLayout(nodes);

    // Identify multi-line parents to render bus-style connectors (include root)
    const enabled = appSettings.multiLineChildrenEnabled !== false;
    const threshold = appSettings.multiLineChildrenThreshold || 20;
    const mlParents = enabled
        ? nodes.filter(p => (p.children || []).length >= threshold)
        : [];
    const excludedTargets = new Set();
    mlParents.forEach(p => (p.children || []).forEach(c => excludedTargets.add(c.data.id)));

    const stdLinks = links.filter(d => !excludedTargets.has(d.target.data.id));

    const link = linkLayer.selectAll('.std-link')
        .data(stdLinks, d => d.target.data.id);

    const linkEnter = link.enter()
        .append('path')
        .attr('class', 'link std-link')
        .attr('d', d => {
            const o = {x: source.x0 || source.x, y: source.y0 || source.y};
            return diagonal(o, o);
        });

    link.merge(linkEnter)
        .transition()
        .duration(500)
        .attr('d', d => diagonal(d.source, d.target));

    link.exit()
        .transition()
        .duration(500)
        .attr('d', d => {
            const o = {x: source.x, y: source.y};
            return diagonal(o, o);
        })
        .remove();

    // Render bus-style connectors for multi-line parents
    function buildBusPath(parent) {
        const children = (parent.children || []).slice().sort((a, b) => a.x - b.x);
        if (!children.length) return '';
        const rowsMap = new Map();
        children.forEach(ch => {
            const key = currentLayout === 'vertical' ? Math.round(ch.y) : Math.round(ch.x);
            if (!rowsMap.has(key)) rowsMap.set(key, []);
            rowsMap.get(key).push(ch);
        });
        const rows = Array.from(rowsMap.entries()).sort((a, b) => a[0] - b[0]).map(e => e[1]);
        let d = '';
        if (currentLayout === 'vertical') {
            const spineYs = rows.map(r => Math.min(...r.map(ch => ch.y - nodeHeight / 2)) - 12);
            const topSpineY = Math.min(...spineYs);
            const bottomSpineY = Math.max(...spineYs);
            d += `M ${parent.x} ${parent.y + nodeHeight/2} L ${parent.x} ${bottomSpineY}`;
            rows.forEach((row, i) => {
                const spineY = spineYs[i];
                const xs = row.map(ch => ch.x);
                const left = Math.min(...xs);
                const right = Math.max(...xs);
                d += ` M ${left} ${spineY} L ${right} ${spineY}`;
                row.forEach(ch => {
                    const childTop = ch.y - nodeHeight/2;
                    d += ` M ${ch.x} ${spineY} L ${ch.x} ${childTop}`;
                });
            });
        } else {
            const spineXs = rows.map(r => Math.min(...r.map(ch => ch.x - nodeWidth / 2)) - 12);
            const leftMostSpineX = Math.min(...spineXs);
            const rightMostSpineX = Math.max(...spineXs);
            d += `M ${parent.x + nodeWidth/2} ${parent.y} L ${rightMostSpineX} ${parent.y}`;
            rows.forEach((row, i) => {
                const spineX = spineXs[i];
                const ys = row.map(ch => ch.y);
                const top = Math.min(...ys);
                const bottom = Math.max(...ys);
                d += ` M ${spineX} ${top} L ${spineX} ${bottom}`;
                row.forEach(ch => {
                    const childLeft = ch.x - nodeWidth/2;
                    d += ` M ${spineX} ${ch.y} L ${childLeft} ${ch.y}`;
                });
            });
        }
        return d;
    }

    const bus = linkLayer.selectAll('path.bus-group')
        .data(mlParents, d => d.data.id);
    bus.enter()
        .append('path')
        .attr('class', 'link bus-group')
        .merge(bus)
        .transition()
        .duration(500)
        .attr('d', d => buildBusPath(d));
    bus.exit().remove();

    const node = nodeLayer.selectAll('.node')
        .data(nodes, d => d.data.id);

    const nodeEnter = node.enter()
        .append('g')
        .attr('class', d => {
            let cls = d.depth === 0 ? 'node ceo' : 'node';
            if (isHiddenNode(d)) cls += ' hidden-subtree';
            return cls;
        })
        .attr('transform', d => `translate(${source.x0 || source.x}, ${source.y0 || source.y})`)
        .on('click', (event, d) => {
            event.stopPropagation();
            showEmployeeDetail(d.data);
        });

    nodeEnter.append('rect')
        .attr('class', d => {
            let classes = 'node-rect';
            if (appSettings.highlightNewEmployees !== false && d.data.isNewEmployee) {
                classes += ' new-employee';
            }
            return classes;
        })
        .attr('x', -nodeWidth/2)
        .attr('y', -nodeHeight/2)
        .attr('width', nodeWidth)
        .attr('height', nodeHeight)
        .style('fill', d => {
            const nodeColors = appSettings.nodeColors || {};
            switch(d.depth) {
                case 0: return nodeColors.level0 || '#90EE90';
                case 1: return nodeColors.level1 || '#FFFFE0';
                case 2: return nodeColors.level2 || '#E0F2FF';
                case 3: return nodeColors.level3 || '#FFE4E1';
                case 4: return nodeColors.level4 || '#E8DFF5';
                case 5: return nodeColors.level5 || '#FFEAA7';
                default: return '#F0F0F0'; 
            }
        })
        .style('stroke', d => {
            if (appSettings.highlightNewEmployees !== false && d.data.isNewEmployee) {
                return null;
            }
            const nodeColors = appSettings.nodeColors || {};
            let fillColor;
            switch(d.depth) {
                case 0: fillColor = nodeColors.level0 || '#90EE90'; break;
                case 1: fillColor = nodeColors.level1 || '#FFFFE0'; break;
                case 2: fillColor = nodeColors.level2 || '#E0F2FF'; break;
                case 3: fillColor = nodeColors.level3 || '#FFE4E1'; break;
                case 4: fillColor = nodeColors.level4 || '#E8DFF5'; break;
                case 5: fillColor = nodeColors.level5 || '#FFEAA7'; break;
                default: fillColor = '#F0F0F0';
            }
            return adjustColor(fillColor, -50);
        })
        .style('stroke-width', '2px');

    if (appSettings.showProfileImages !== false) {
        applyProfileImageAttributes(nodeEnter.append('image'));
    }

    nodeEnter.append('text')
        .attr('class', 'node-text')
        .attr('x', getLabelOffsetX())
        .attr('y', -10)
        .attr('text-anchor', getLabelAnchor())
        .style('font-weight', 'bold')
        .style('font-size', d => getNameFontSizePx(d.data.name))
        .text(d => d.data.name);

    nodeEnter.append('text')
        .attr('class', 'node-title')
        .attr('x', getLabelOffsetX())
        .attr('y', 5)
        .attr('text-anchor', getLabelAnchor())
        .style('font-size', d => getTitleFontSizePx(d.data.title || ''))
        .text(d => getTrimmedTitle(d.data.title || ''));

    if (appSettings.showDepartments !== false) {
        nodeEnter.append('text')
            .attr('class', 'node-department')
            .attr('x', getLabelOffsetX())
            .attr('y', 18)
            .attr('text-anchor', getLabelAnchor())
            .style('font-size', d => getDepartmentFontSizePx(d.data.department || 'Not specified'))
            .style('font-style', 'italic')
            .style('fill', '#666')
            .text(d => d.data.department || 'Not specified');
    }

    if (appSettings.showEmployeeCount !== false) {
        const countGroup = nodeEnter.append('g')
            .attr('class', 'count-badge')
            .style('display', d => {
                const totalCount = d._children?.length || d.children?.length || 0;
                return totalCount > 0 ? 'block' : 'none';
            });

        countGroup.append('circle')
            .attr('cx', -nodeWidth/2 + 15)
            .attr('cy', -nodeHeight/2 + 15)
            .attr('r', 12)
            .style('fill', '#ff6b6b')
            .style('stroke', 'white')
            .style('stroke-width', '2px');

        countGroup.append('text')
            .attr('x', -nodeWidth/2 + 15)
            .attr('y', -nodeHeight/2 + 19)
            .attr('text-anchor', 'middle')
            .style('fill', 'white')
            .style('font-size', '11px')
            .style('font-weight', 'bold')
            .text(d => {
                const count = d._children?.length || d.children?.length || 0;
                return count > 99 ? '99+' : count;
            });
    }

    const expandBtn = nodeEnter.append('g')
        .attr('class', 'expand-group')
        .style('display', d => (d._children?.length || d.children?.length) ? 'block' : 'none')
        .on('click', (event, d) => {
            event.stopPropagation();
            toggle(d);
        });

    expandBtn.append('circle')
        .attr('class', 'expand-btn')
        .attr('cy', currentLayout === 'vertical' ? nodeHeight/2 + 10 : 0)
        .attr('cx', currentLayout === 'horizontal' ? nodeWidth/2 + 10 : 0)
        .attr('r', 10);

    expandBtn.append('text')
        .attr('class', 'expand-text')
        .attr('y', currentLayout === 'vertical' ? nodeHeight/2 + 15 : 4)
        .attr('x', currentLayout === 'horizontal' ? nodeWidth/2 + 10 : 0)
        .attr('text-anchor', 'middle')
        .text(d => d._children?.length ? '+' : '-');

    // Eye icon toggle (placed top-right inside node)
    nodeEnter.append('text')
        .attr('class', 'hide-toggle')
        .attr('x', nodeWidth/2 - 14)
        .attr('y', -nodeHeight/2 + 14)
        .attr('text-anchor', 'middle')
        .text(d => hiddenNodeIds.has(d.data.id) ? '🙈' : '👁')
        .on('click', (event, d) => {
            event.stopPropagation();
            toggleHideNode(d);
        })
        .append('title')
        .text(d => hiddenNodeIds.has(d.data.id) ? 'Show this subtree' : 'Hide this subtree');

    if (appSettings.highlightNewEmployees !== false) {
        const newBadgeGroup = nodeEnter.append('g')
            .attr('class', 'new-employee-badge')
            .style('display', d => d.data.isNewEmployee ? 'block' : 'none');

        newBadgeGroup.append('rect')
            .attr('class', 'new-badge')
            .attr('x', nodeWidth/2 - 45)
            .attr('y', -nodeHeight/2 - 10)
            .attr('width', 35)
            .attr('height', 18)
            .attr('rx', 9)
            .attr('ry', 9);

        newBadgeGroup.append('text')
            .attr('class', 'new-badge-text')
            .attr('x', nodeWidth/2 - 27)
            .attr('y', -nodeHeight/2 + 2)
            .attr('text-anchor', 'middle')
            .text('NEW');
    }


    const nodeMerge = node.merge(nodeEnter);

    const nodeUpdate = nodeMerge
        .attr('class', d => {
            let cls = d.depth === 0 ? 'node ceo' : 'node';
            if (isHiddenNode(d)) cls += ' hidden-subtree';
            return cls;
        })
        .transition()
        .duration(500)
        .attr('transform', d => `translate(${d.x}, ${d.y})`);

    // Update eye icons and tooltips on merged selection (after transition start)
    nodeMerge.selectAll('text.hide-toggle')
        .text(d => hiddenNodeIds.has(d.data.id) ? '🙈' : '👁')
        .each(function(d){
            const titleEl = this.querySelector('title');
            if (titleEl) titleEl.textContent = hiddenNodeIds.has(d.data.id) ? 'Show this subtree' : 'Hide this subtree';
        });

    nodeUpdate.select('.expand-text')
        .text(d => d._children?.length ? '+' : '-')
        .attr('y', currentLayout === 'vertical' ? nodeHeight/2 + 15 : 4)
        .attr('x', currentLayout === 'horizontal' ? nodeWidth/2 + 10 : 0);

    nodeUpdate.select('.expand-btn')
        .attr('cy', currentLayout === 'vertical' ? nodeHeight/2 + 10 : 0)
        .attr('cx', currentLayout === 'horizontal' ? nodeWidth/2 + 10 : 0);

    nodeUpdate.select('.expand-group')
        .style('display', d => (d._children?.length || d.children?.length) ? 'block' : 'none');

    if (appSettings.showEmployeeCount !== false) {
        nodeUpdate.select('.count-badge')
            .style('display', d => {
                const totalCount = d._children?.length || d.children?.length || 0;
                return totalCount > 0 ? 'block' : 'none';
            });

        nodeUpdate.select('.count-badge text')
            .text(d => {
                const count = d._children?.length || d.children?.length || 0;
                return count > 99 ? '99+' : count;
            });
    }

    if (appSettings.showProfileImages !== false) {
        nodeMerge.each(function(d) {
            const nodeSel = d3.select(this);
            let img = nodeSel.select('image.profile-image');
            if (img.empty()) {
                img = nodeSel.insert('image', 'text');
            }
            applyProfileImageAttributes(img);
        });
    } else {
        nodeMerge.selectAll('image.profile-image').remove();
    }

    nodeMerge.select('.node-text')
        .attr('x', getLabelOffsetX())
        .attr('text-anchor', getLabelAnchor())
        .style('font-size', d => getNameFontSizePx(d.data.name))
        .text(d => d.data.name);

    nodeMerge.select('.node-title')
        .attr('x', getLabelOffsetX())
        .attr('text-anchor', getLabelAnchor())
        .style('font-size', d => getTitleFontSizePx(d.data.title || ''))
        .text(d => getTrimmedTitle(d.data.title || ''));

    if (appSettings.showDepartments !== false) {
        nodeMerge.select('.node-department')
            .attr('x', getLabelOffsetX())
            .attr('text-anchor', getLabelAnchor())
            .style('font-size', d => getDepartmentFontSizePx(d.data.department || 'Not specified'))
            .text(d => d.data.department || 'Not specified');
    }

    node.exit()
        .transition()
        .duration(500)
        .attr('transform', d => `translate(${source.x}, ${source.y})`)
        .remove();

    nodes.forEach(d => {
        d.x0 = d.x;
        d.y0 = d.y;
    });
}

// Arrange many direct reports in multiple rows to avoid overlap
function applyMultiLineChildrenLayout(nodes) {
    const enabled = appSettings.multiLineChildrenEnabled !== false;
    const threshold = appSettings.multiLineChildrenThreshold || 20;
    if (!enabled) return;

    // Helper: shift an entire subtree rooted at node by dx, dy
    function shiftSubtree(rootNode, dx, dy) {
        if ((dx === 0 && dy === 0) || !rootNode) return;
        nodes.forEach(n => {
            let cur = n;
            while (cur) {
                if (cur === rootNode) {
                    n.x += dx;
                    n.y += dy;
                    break;
                }
                cur = cur.parent;
            }
        });
    }

    // Helper: check if anc is an ancestor of node
    function isAncestor(anc, node) {
        let cur = node;
        while (cur) {
            if (cur === anc) return true;
            cur = cur.parent;
        }
        return false;
    }

    // Helper: compute subtree bounds for a node across provided nodes
    function getSubtreeBounds(rootNode) {
        let minX = Infinity, maxX = -Infinity, minY = Infinity, maxY = -Infinity;
        nodes.forEach(n => {
            if (isAncestor(rootNode, n)) {
                const left = n.x - nodeWidth / 2;
                const right = n.x + nodeWidth / 2;
                const top = n.y - nodeHeight / 2;
                const bottom = n.y + nodeHeight / 2;
                if (left < minX) minX = left;
                if (right > maxX) maxX = right;
                if (top < minY) minY = top;
                if (bottom > maxY) maxY = bottom;
            }
        });
        return { minX, maxX, minY, maxY };
    }


    nodes.forEach(parent => {
        const kids = parent.children || [];
        if (!kids.length) return;
    if (kids.length < threshold) return;

        // Determine row/column layout and minimize empty slots in last row
        // Preserve D3's left-to-right ordering to reduce crossing
        const orderedKids = kids.slice().sort((a, b) => a.x - b.x);
        const n = orderedKids.length;
        let columns = Math.ceil(Math.sqrt(n));
        let rows = Math.ceil(n / columns);
        columns = Math.ceil(n / rows); // refine to reduce underfill
        const hSpacing = nodeWidth + 36;
        const vSpacing = levelHeight; // keep consistent per-level spacing
        const totalHeight = (rows - 1) * vSpacing;

        orderedKids.forEach((child, idx) => {
            const col = idx % columns;
            const row = Math.floor(idx / columns);

            let targetX, targetY;
            if (currentLayout === 'vertical') {
                // Center based on actual items in this row
                let itemsInRow = (row < rows - 1) ? columns : (n - (rows - 1) * columns);
                if (itemsInRow <= 0) itemsInRow = columns;
                const rowWidth = (itemsInRow - 1) * hSpacing;
                const colInRow = col % itemsInRow;
                targetX = parent.x - rowWidth / 2 + colInRow * hSpacing;
                targetY = parent.y + (row + 1) * vSpacing;
            } else {
                let itemsInRow = (row < rows - 1) ? columns : (n - (rows - 1) * columns);
                if (itemsInRow <= 0) itemsInRow = columns;
                const rowWidth = (itemsInRow - 1) * hSpacing;
                const colInRow = col % itemsInRow;
                targetX = parent.x + (row + 1) * vSpacing;
                targetY = parent.y - rowWidth / 2 + colInRow * hSpacing;
            }

            const dx = targetX - child.x;
            const dy = targetY - child.y;
            shiftSubtree(child, dx, dy);
        });

        // Keep compaction logic light for stability
    });

    // Pass 2: Compress ancestors around multi-lined groups, do not alter the group itself
    const gapBetweenSubtrees = (appSettings.multiLineCompactGap ?? 36); // edge-to-edge gap
    // Identify multi-lined parents (where wrapping occurred)
    const mlParents = nodes.filter(p => (p.children || []).length >= threshold);
    const ancestorSet = new Set();
    mlParents.forEach(mlp => {
        let anc = mlp.parent;
        while (anc) {
            ancestorSet.add(anc);
            anc = anc.parent;
        }
    });

    // Process ancestors deep-to-shallow so higher levels can adapt after recentering
    const ancestors = Array.from(ancestorSet).sort((a, b) => b.depth - a.depth);
    ancestors.forEach(parent => {
        const kids = parent.children || [];
        if (kids.length < 2) return;
        // Build intervals for each child's subtree as fixed blocks
        const intervals = kids.map(child => {
            const b = getSubtreeBounds(child);
            if (currentLayout === 'vertical') {
                const width = Math.max(b.maxX - b.minX, nodeWidth);
                const center = (b.maxX + b.minX) / 2;
                return { child, width, center };
            } else {
                const width = Math.max(b.maxY - b.minY, nodeWidth);
                const center = (b.maxY + b.minY) / 2;
                return { child, width, center };
            }
        });
        // Preserve order by current center
        intervals.sort((a, b) => a.center - b.center);
        const totalWidth = intervals.reduce((sum, it) => sum + it.width, 0) + gapBetweenSubtrees * (intervals.length - 1);
        const groupCenter = currentLayout === 'vertical' ? parent.x : parent.y;
        const start = groupCenter - totalWidth / 2;
        let cursor = start;
        intervals.forEach(it => {
            const targetCenter = cursor + it.width / 2;
            const delta = targetCenter - it.center;
            if (currentLayout === 'vertical') {
                shiftSubtree(it.child, delta, 0);
            } else {
                shiftSubtree(it.child, 0, delta);
            }
            cursor += it.width + gapBetweenSubtrees;
        });

        // After packing children, recenter parent directly above/between them to avoid one-sided gaps
        const childBounds = kids.map(ch => getSubtreeBounds(ch));
        if (currentLayout === 'vertical') {
            const left = Math.min(...childBounds.map(b => b.minX));
            const right = Math.max(...childBounds.map(b => b.maxX));
            const desired = (left + right) / 2;
            const deltaParent = desired - parent.x;
            if (Math.abs(deltaParent) > 0.1) shiftSubtree(parent, deltaParent, 0);
        } else {
            const top = Math.min(...childBounds.map(b => b.minY));
            const bottom = Math.max(...childBounds.map(b => b.maxY));
            const desired = (top + bottom) / 2;
            const deltaParent = desired - parent.y;
            if (Math.abs(deltaParent) > 0.1) shiftSubtree(parent, 0, deltaParent);
        }
    });
}

function diagonal(s, d) {
    if (currentLayout === 'vertical') {
        const midY = (s.y + d.y) / 2;
        return `M ${s.x} ${s.y + nodeHeight/2}
                L ${s.x} ${midY}
                L ${d.x} ${midY}
                L ${d.x} ${d.y - nodeHeight/2}`;
    } else {
        const midX = (s.x + d.x) / 2;
        return `M ${s.x + nodeWidth/2} ${s.y}
                L ${midX} ${s.y}
                L ${midX} ${d.y}
                L ${d.x - nodeWidth/2} ${d.y}`;
    }
}

function toggle(d) {
    if (d.children) {
        d._children = d.children;
        d.children = null;
    } else {
        d.children = d._children;
        d._children = null;
        
        if (d.children) {
            d.children.forEach(child => {
                if (child.depth >= 2 && child.children) {
                    child._children = child.children;
                    child.children = null;
                }
            });
        }
    }
    update(d);
}

function expandAll() {
    root.each(d => {
        if (d._children) {
            d.children = d._children;
            d._children = null;
        }
    });
    update(root);
}

function collapseAll() {
    root.each(d => {
        if (d.depth >= 1 && d.children) {
            d._children = d.children;
            d.children = null;
        }
    });
    update(root);
}

function resetZoom() {
    fitToScreen({ duration: 500, resetUser: true });
}

function fitToScreen(options = {}) {
    if (!root || !svg) return;
    updateSvgSize();

    const { duration = 750, resetUser = true } = options;
    const treeLayout = createTreeLayout();
    const treeData = treeLayout(root);
    const nodes = treeData.descendants();
    if (currentLayout === 'horizontal') {
        nodes.forEach(d => { const t = d.x; d.x = d.y; d.y = t; });
    }
    applyMultiLineChildrenLayout(nodes);
    
    if (nodes.length === 0) return;
    
    const minX = d3.min(nodes, d => d.x) - nodeWidth / 2;
    const maxX = d3.max(nodes, d => d.x) + nodeWidth / 2;
    const minY = d3.min(nodes, d => d.y) - nodeHeight / 2;
    const maxY = d3.max(nodes, d => d.y) + nodeHeight / 2;
    
    const width = maxX - minX;
    const height = maxY - minY;
    
    const container = document.getElementById('orgChart');
    const containerWidth = Math.max(container.clientWidth, 1);
    const containerHeight = Math.max(container.clientHeight || 0, 1);
    
    const scale = Math.min(
        width === 0 ? 1 : (containerWidth * 0.9) / width,
        height === 0 ? 1 : (containerHeight * 0.9) / height,
        1 
    );
    
    const centerX = (minX + maxX) / 2;
    const centerY = (minY + maxY) / 2;
    
    const transform = d3.zoomIdentity
        .translate(containerWidth / 2, containerHeight / 2)
        .scale(scale)
        .translate(-centerX, -centerY);

    applyZoomTransform(transform, { duration, resetUser });
}

function zoomIn() {
    if (!svg) return;
    const transition = svg.transition().call(zoom.scaleBy, 1.2);
    transition.on('end', () => { userAdjustedZoom = true; });
    transition.on('interrupt', () => { userAdjustedZoom = true; });
}

function zoomOut() {
    if (!svg) return;
    const transition = svg.transition().call(zoom.scaleBy, 0.8);
    transition.on('end', () => { userAdjustedZoom = true; });
    transition.on('interrupt', () => { userAdjustedZoom = true; });
}

function getBounds(printRoot) {
    const treeLayout = createTreeLayout();
    const treeData = treeLayout(printRoot);
    const nodes = treeData.descendants();
    applyMultiLineChildrenLayout(nodes);
    const minX = d3.min(nodes, d => d.x) - nodeWidth / 2 - 20;
    const maxX = d3.max(nodes, d => d.x) + nodeWidth / 2 + 20;
    const minY = d3.min(nodes, d => d.y) - nodeHeight / 2 - 20;
    const maxY = d3.max(nodes, d => d.y) + nodeHeight / 2 + 50;
    return { minX, maxX, minY, maxY };
}

function buildExpandedData(node) {
    const copy = { data: node.data, depth: node.depth };
    const allKids = [];
    if (node.children) allKids.push(...node.children);
    if (node._children) allKids.push(...node._children);
    if (allKids.length) {
        copy.children = allKids.map(child => buildExpandedData(child));
    }
    copy.hasCollapsedChildren = !!(node._children && node._children.length);
    return copy;
}

function printChart() {
    createExportSVG(false).then(svgElement => {
        const printWin = window.open('', '_blank');
        printWin.document.write('<html><head><title>Org Chart Print</title>');
        printWin.document.write('<style>@page { margin: 0.5cm; } body { margin:0; padding:0; }</style>');
        printWin.document.write('</head><body>');
        // Clone so we don't mutate original
        const clone = svgElement.cloneNode(true);
        // Fit to page via CSS width 100%
        clone.removeAttribute('width');
        clone.removeAttribute('height');
        clone.style.width = '100%';
        clone.style.height = 'auto';
        printWin.document.body.appendChild(clone);
        printWin.document.write('</body></html>');
        printWin.document.close();
        printWin.focus();
        printWin.print();
    }).catch(err => console.error('Print failed:', err));
}

async function exportToImage(format = 'svg', exportFullChart = false) {
    const svgElement = await createExportSVG(exportFullChart);
    const svgString = new XMLSerializer().serializeToString(svgElement);
    
    if (format === 'svg') {
        // Export as SVG
        const svgBlob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
        const url = URL.createObjectURL(svgBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `org-chart-${new Date().toISOString().split('T')[0]}.svg`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    } else if (format === 'png') {
        // Convert to PNG using HTML5 Canvas approach
        console.log('Starting PNG export...');
        
        try {
            // Parse the SVG to get dimensions
            const parser = new DOMParser();
            const svgDoc = parser.parseFromString(svgString, 'image/svg+xml');
            const svgElement = svgDoc.documentElement;
            
            // Extract dimensions from SVG
            const svgWidth = parseFloat(svgElement.getAttribute('width')) || 800;
            const svgHeight = parseFloat(svgElement.getAttribute('height')) || 600;
            
            console.log('SVG dimensions extracted:', svgWidth, 'x', svgHeight);
            
            // Create canvas with better scaling
            const scale = window.devicePixelRatio || 2;
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            
            canvas.width = svgWidth * scale;
            canvas.height = svgHeight * scale;
            
            // Set CSS size for proper scaling
            canvas.style.width = svgWidth + 'px';
            canvas.style.height = svgHeight + 'px';
            
            // Scale context
            ctx.scale(scale, scale);
            
            // White background
            ctx.fillStyle = 'white';
            ctx.fillRect(0, 0, svgWidth, svgHeight);
            
            console.log('Canvas created with dimensions:', canvas.width, 'x', canvas.height);
            
            // Try multiple SVG loading methods for better compatibility
            const tryMethod1 = () => {
                const svgDataUrl = `data:image/svg+xml;base64,${btoa(svgString)}`;
                const img = new Image();
                
                img.onload = function() {
                    console.log('Method 1 - SVG loaded successfully');
                    ctx.drawImage(img, 0, 0, svgWidth, svgHeight);
                    downloadCanvas();
                };
                
                img.onerror = function(error) {
                    console.warn('Method 1 failed, trying method 2:', error);
                    tryMethod2();
                };
                
                img.crossOrigin = 'anonymous';
                img.src = svgDataUrl;
            };
            
            const tryMethod2 = () => {
                const svgDataUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svgString)}`;
                const img = new Image();
                
                img.onload = function() {
                    console.log('Method 2 - SVG loaded successfully');
                    ctx.drawImage(img, 0, 0, svgWidth, svgHeight);
                    downloadCanvas();
                };
                
                img.onerror = function(error) {
                    console.warn('Method 2 failed, trying fallback method:', error);
                    tryFallback();
                };
                
                img.crossOrigin = 'anonymous';
                img.src = svgDataUrl;
            };
            
            const tryFallback = () => {
                console.log('Using fallback: rendering SVG directly to canvas');
                try {
                    // Fallback: Create a simplified version without external resources
                    const simplifiedSvg = svgString
                        .replace(/xlink:href="[^"]*api\/photo[^"]*"/g, 'xlink:href=""')
                        .replace(/<image[^>]*api\/photo[^>]*\/>/g, '');
                    
                    const svgDataUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(simplifiedSvg)}`;
                    const img = new Image();
                    
                    img.onload = function() {
                        console.log('Fallback - SVG loaded successfully');
                        ctx.drawImage(img, 0, 0, svgWidth, svgHeight);
                        downloadCanvas();
                    };
                    
                    img.onerror = function(error) {
                        console.error('All methods failed:', error);
                        alert('Error loading chart for PNG export. Please try using SVG export instead.');
                    };
                    
                    img.crossOrigin = 'anonymous';
                    img.src = svgDataUrl;
                } catch (error) {
                    console.error('Fallback method failed:', error);
                    alert('Error creating PNG export. Please try using SVG export instead.');
                }
            };
            
            const downloadCanvas = () => {
                canvas.toBlob(function(blob) {
                    if (blob) {
                        console.log('PNG blob created successfully, size:', blob.size);
                        const pngUrl = URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        a.href = pngUrl;
                        a.download = `org-chart-${new Date().toISOString().split('T')[0]}.png`;
                        document.body.appendChild(a);
                        a.click();
                        document.body.removeChild(a);
                        URL.revokeObjectURL(pngUrl);
                    } else {
                        console.error('Failed to create PNG blob');
                        alert('Error creating PNG file. Please try again.');
                    }
                }, 'image/png', 0.95);
            };
            
            // Start with method 1
            tryMethod1();
            
        } catch (error) {
            console.error('Error in PNG export setup:', error);
            alert('Error setting up PNG export: ' + error.message);
        }
    }
}

async function exportToPDF(exportFullChart = false) {
    try {
        console.log('Starting PDF export...');
        
        // Check if data is loaded
        if (!currentData) {
            alert('No organizational chart data available. Please wait for the data to load or refresh the page.');
            return;
        }
        
        // Check if root is available for visible chart export
        if (!exportFullChart && !root) {
            alert('No chart is currently visible. Please ensure the organizational chart is loaded.');
            return;
        }
        
        if (typeof window.jspdf === 'undefined') {
            console.error('jsPDF library not loaded');
            alert('PDF library not loaded. Please refresh the page.');
            return;
        }
        
        console.log('Libraries loaded successfully, currentData available:', !!currentData, 'root available:', !!root);

        // Use the existing SVG creation function and scale it to PDF
        const svgElement = await createExportSVG(exportFullChart);
        console.log('SVG created successfully');
        
        // Get SVG dimensions
        const svgWidth = parseFloat(svgElement.getAttribute('width'));
        const svgHeight = parseFloat(svgElement.getAttribute('height'));
        
        console.log(`SVG dimensions: ${svgWidth} x ${svgHeight}`);
        
        // Create PDF with appropriate orientation
        const { jsPDF } = window.jspdf;
        const isLandscape = svgWidth > svgHeight;
        const pdf = new jsPDF(isLandscape ? 'l' : 'p', 'mm', 'a4');
        
        // Get PDF page dimensions
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const margin = 10;
        const availableWidth = pageWidth - (2 * margin);
        const availableHeight = pageHeight - (2 * margin);
        
        // Calculate scale to fit SVG in PDF page
        const scaleX = availableWidth / (svgWidth * 0.264583); // Convert px to mm
        const scaleY = availableHeight / (svgHeight * 0.264583);
        const scale = Math.min(scaleX, scaleY);
        
        // Calculate final dimensions and position
        const finalWidth = svgWidth * 0.264583 * scale;
        const finalHeight = svgHeight * 0.264583 * scale;
        const x = margin + (availableWidth - finalWidth) / 2;
        const y = margin + (availableHeight - finalHeight) / 2;
        
        console.log(`PDF: ${pageWidth}x${pageHeight}mm, Final: ${finalWidth.toFixed(1)}x${finalHeight.toFixed(1)}mm, Scale: ${scale.toFixed(3)}`);
        
        // Convert SVG to data URL
        const svgString = new XMLSerializer().serializeToString(svgElement);
        const svgDataUrl = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgString)));
        
        // Add SVG as image to PDF
        try {
            pdf.addImage(svgDataUrl, 'SVG', x, y, finalWidth, finalHeight);
        } catch (error) {
            console.warn('SVG addImage failed, trying PNG conversion:', error);
            
            // Fallback: convert SVG to canvas first
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            const img = new Image();
            
            await new Promise((resolve, reject) => {
                img.onload = () => {
                    canvas.width = svgWidth;
                    canvas.height = svgHeight;
                    ctx.fillStyle = 'white';
                    ctx.fillRect(0, 0, canvas.width, canvas.height);
                    ctx.drawImage(img, 0, 0);
                    resolve();
                };
                img.onerror = reject;
                img.src = svgDataUrl;
            });
            
            const pngDataUrl = canvas.toDataURL('image/png', 0.9);
            pdf.addImage(pngDataUrl, 'PNG', x, y, finalWidth, finalHeight);
        }

        const fileName = `org-chart-${new Date().toISOString().split('T')[0]}.pdf`;
        pdf.save(fileName);
        console.log('PDF exported successfully:', fileName);

    } catch (error) {
        console.error('Error in exportToPDF:', error);
        // Ensure cleanup of any temporary elements
        try {
            const orphanSvg = document.querySelector('svg[style*="position: absolute"]');
            if (orphanSvg) {
                document.body.removeChild(orphanSvg);
            }
            const orphanImg = document.querySelector('img[style*="position: absolute"]');
            if (orphanImg) {
                document.body.removeChild(orphanImg);
            }
        } catch (cleanupError) {
            console.warn('Cleanup error:', cleanupError);
        }
        alert('An error occurred during PDF export: ' + (error.message || error || 'Unknown error'));
    }
}

async function imageToDataUrl(url) {
    try {
        const response = await fetch(url);
        if (!response.ok) {
            console.warn(`Failed to fetch image: ${url}, status: ${response.status}`);
            return null;
        }
        const blob = await response.blob();
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result);
            reader.onerror = (err) => {
                console.error('FileReader error:', err);
                reject(err);
            };
            reader.readAsDataURL(blob);
        });
    } catch (error) {
        console.error('Error converting image to data URL:', url, error);
        return null;
    }
}

async function createExportSVG(exportFullChart = false) {
    let nodesToExport, linksToExport;
    
    // Validate data availability
    if (!currentData) {
        throw new Error('No organizational chart data available');
    }
    
    if (!exportFullChart && !root) {
        throw new Error('No visible chart data available');
    }
    
    // Determine which nodes to export
    if (exportFullChart) {
        const fullHierarchy = d3.hierarchy(currentData);
    const treeLayout = createTreeLayout();
        const treeData = treeLayout(fullHierarchy);
        nodesToExport = treeData.descendants();
        linksToExport = treeData.links();
    } else {
    const treeLayout = createTreeLayout();
        const treeData = treeLayout(root);
        nodesToExport = treeData.descendants();
        linksToExport = treeData.links();
    }

    // Filter out hidden nodes and descendants
    nodesToExport = nodesToExport.filter(n => !isHiddenNode(n));
    const hiddenIdSet = new Set(nodesToExport.map(n => n.data.id));
    linksToExport = linksToExport.filter(l => hiddenIdSet.has(l.source.data.id) && hiddenIdSet.has(l.target.data.id));
    
    // Validate that we have nodes to export
    if (!nodesToExport || nodesToExport.length === 0) {
        throw new Error('No chart nodes available for export');
    }
    
    // Adjust for horizontal layout
    if (currentLayout === 'horizontal') {
        nodesToExport.forEach(d => { [d.x, d.y] = [d.y, d.x]; });
    }

    // Apply our layout adjustments before bounds and export
    applyMultiLineChildrenLayout(nodesToExport);
    // No extra compaction
    
    // Calculate bounds
    const padding = 50;
    const minX = d3.min(nodesToExport, d => d.x) - nodeWidth/2 - padding;
    const maxX = d3.max(nodesToExport, d => d.x) + nodeWidth/2 + padding;
    const minY = d3.min(nodesToExport, d => d.y) - nodeHeight/2 - padding;
    const maxY = d3.max(nodesToExport, d => d.y) + nodeHeight/2 + padding;
    const width = maxX - minX;
    const height = maxY - minY;
    
    // Create SVG element
    const exportSvg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    Object.assign(exportSvg.style, { fontFamily: 'Arial, sans-serif' });
    exportSvg.setAttribute('width', width);
    exportSvg.setAttribute('height', height);
    exportSvg.setAttribute('viewBox', `${minX} ${minY} ${width} ${height}`);
    exportSvg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
    exportSvg.setAttribute('xmlns:xlink', 'http://www.w3.org/1999/xlink');
    
    // Add styles
    const style = document.createElementNS('http://www.w3.org/2000/svg', 'style');
    style.textContent = `
        .link { fill: none; stroke: #999; stroke-width: 2px; }
        .node-rect { rx: 4; ry: 4; }
        .node-text { font-size: 14px; fill: #333; font-weight: 600; }
        .node-title { font-size: 11px; fill: #555; }
        .node-department { font-size: 9px; fill: #666; font-style: italic; }
    `;
    exportSvg.appendChild(style);

    // Add white background
    const bg = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
    bg.setAttribute('x', minX);
    bg.setAttribute('y', minY);
    bg.setAttribute('width', width);
    bg.setAttribute('height', height);
    bg.setAttribute('fill', 'white');
    exportSvg.appendChild(bg);
    
    const g = document.createElementNS('http://www.w3.org/2000/svg', 'g');
    const linksGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
    linksGroup.setAttribute('class', 'links');
    const nodesGroup = document.createElementNS('http://www.w3.org/2000/svg', 'g');
    nodesGroup.setAttribute('class', 'nodes');
    
    // Identify multi-line parents for bus connectors
    const enabled = appSettings.multiLineChildrenEnabled !== false;
    const threshold = appSettings.multiLineChildrenThreshold || 20;
    const mlParents = enabled
        ? nodesToExport.filter(p => (p.children || []).length >= threshold)
        : [];
    const excludedTargets = new Set();
    mlParents.forEach(p => (p.children || []).forEach(c => excludedTargets.add(c.data.id)));
    const stdLinks = linksToExport.filter(d => !excludedTargets.has(d.target.data.id));

    // Draw standard links
    stdLinks.forEach(link => {
        const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        path.setAttribute('class', 'link');
        path.setAttribute('d', diagonal(link.source, link.target));
        linksGroup.appendChild(path);
    });

    // Draw bus connectors
    function buildBusPath(parent) {
        const children = (parent.children || []).slice().sort((a, b) => a.x - b.x);
        if (!children.length) return '';
        const rowsMap = new Map();
        children.forEach(ch => {
            const key = currentLayout === 'vertical' ? Math.round(ch.y) : Math.round(ch.x);
            if (!rowsMap.has(key)) rowsMap.set(key, []);
            rowsMap.get(key).push(ch);
        });
        const rows = Array.from(rowsMap.entries()).sort((a, b) => a[0] - b[0]).map(e => e[1]);
        let d = '';
        if (currentLayout === 'vertical') {
            const spineYs = rows.map(r => Math.min(...r.map(ch => ch.y - nodeHeight / 2)) - 12);
            const bottomSpineY = Math.max(...spineYs);
            d += `M ${parent.x} ${parent.y + nodeHeight/2} L ${parent.x} ${bottomSpineY}`;
            rows.forEach((row, i) => {
                const spineY = spineYs[i];
                const xs = row.map(ch => ch.x);
                const left = Math.min(...xs);
                const right = Math.max(...xs);
                d += ` M ${left} ${spineY} L ${right} ${spineY}`;
                row.forEach(ch => {
                    const childTop = ch.y - nodeHeight/2;
                    d += ` M ${ch.x} ${spineY} L ${ch.x} ${childTop}`;
                });
            });
        } else {
            const spineXs = rows.map(r => Math.min(...r.map(ch => ch.x - nodeWidth / 2)) - 12);
            const rightMostSpineX = Math.max(...spineXs);
            d += `M ${parent.x + nodeWidth/2} ${parent.y} L ${rightMostSpineX} ${parent.y}`;
            rows.forEach((row, i) => {
                const spineX = spineXs[i];
                const ys = row.map(ch => ch.y);
                const top = Math.min(...ys);
                const bottom = Math.max(...ys);
                d += ` M ${spineX} ${top} L ${spineX} ${bottom}`;
                row.forEach(ch => {
                    const childLeft = ch.x - nodeWidth/2;
                    d += ` M ${spineX} ${ch.y} L ${childLeft} ${ch.y}`;
                });
            });
        }
        return d;
    }

    mlParents.forEach(parent => {
        const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
        path.setAttribute('class', 'link');
        path.setAttribute('d', buildBusPath(parent));
        linksGroup.appendChild(path);
    });
    
    // Pre-fetch all images and convert to data URLs
    const imageCache = new Map();
    const defaultIconDataUrl = await imageToDataUrl(userIconUrl);
    const imagePromises = nodesToExport.map(async (d) => {
        if (appSettings.showProfileImages !== false && d.data.photoUrl && d.data.photoUrl.includes('/api/photo/')) {
            const dataUrl = await imageToDataUrl(window.location.origin + d.data.photoUrl);
            if (dataUrl) imageCache.set(d.data.id, dataUrl);
        }
    });
    await Promise.all(imagePromises);

    // Draw nodes
    for (const d of nodesToExport) {
        const nodeG = document.createElementNS('http://www.w3.org/2000/svg', 'g');
        nodeG.setAttribute('transform', `translate(${d.x}, ${d.y})`);
        
        // Node rectangle
        const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        rect.setAttribute('class', 'node-rect');
        rect.setAttribute('x', -nodeWidth/2);
        rect.setAttribute('y', -nodeHeight/2);
        rect.setAttribute('width', nodeWidth);
        rect.setAttribute('height', nodeHeight);
        
        const nodeColors = appSettings.nodeColors || {};
        let fillColor;
        switch(d.depth) {
            case 0: fillColor = nodeColors.level0 || '#90EE90'; break;
            case 1: fillColor = nodeColors.level1 || '#FFFFE0'; break;
            case 2: fillColor = nodeColors.level2 || '#E0F2FF'; break;
            case 3: fillColor = nodeColors.level3 || '#FFE4E1'; break;
            case 4: fillColor = nodeColors.level4 || '#E8DFF5'; break;
            case 5: fillColor = nodeColors.level5 || '#FFEAA7'; break;
            default: fillColor = '#F0F0F0';
        }
        rect.setAttribute('fill', fillColor);
        rect.setAttribute('stroke', adjustColor(fillColor, -50));
        rect.setAttribute('stroke-width', '2');
    nodeG.appendChild(rect);
        
        // Profile image with circular clipping
        if (appSettings.showProfileImages !== false) {
            // Create unique clip path ID for this image
            const clipId = `clip-${d.data.id || Math.random().toString(36).substr(2, 9)}`;
            
            // Create clip path definition
            const clipPath = document.createElementNS('http://www.w3.org/2000/svg', 'clipPath');
            clipPath.setAttribute('id', clipId);
            const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
            circle.setAttribute('cx', -nodeWidth/2 + 26); // Center of image (8 + 18)
            circle.setAttribute('cy', 0); // Center vertically (-18 + 18)
            circle.setAttribute('r', 18); // Radius for circular crop
            clipPath.appendChild(circle);
            
            // Add clip path to defs (create defs if it doesn't exist)
            let defs = exportSvg.querySelector('defs');
            if (!defs) {
                defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
                exportSvg.insertBefore(defs, exportSvg.firstChild);
            }
            defs.appendChild(clipPath);
            
            // Create the image with clipping applied
            const image = document.createElementNS('http://www.w3.org/2000/svg', 'image');
            const imageUrl = imageCache.get(d.data.id) || defaultIconDataUrl;
            if (imageUrl) {
                image.setAttributeNS('http://www.w3.org/1999/xlink', 'href', imageUrl);
            }
            image.setAttribute('x', -nodeWidth/2 + 8);
            image.setAttribute('y', -18);
            image.setAttribute('width', 36);
            image.setAttribute('height', 36);
            image.setAttribute('clip-path', `url(#${clipId})`);
            nodeG.appendChild(image);
        }

        const textX = appSettings.showProfileImages !== false ? -nodeWidth/2 + 50 : 0;
        const textAnchor = appSettings.showProfileImages !== false ? 'start' : 'middle';
        const textWidth = appSettings.showProfileImages !== false ? nodeWidth - 58 : nodeWidth - 20;
        
        // Name
        const nameText = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        nameText.setAttribute('class', 'node-text');
        nameText.setAttribute('x', textX);
        nameText.setAttribute('y', -10);
        nameText.setAttribute('text-anchor', textAnchor);
        nameText.textContent = d.data.name;
        nodeG.appendChild(nameText);
        
        // Title
        const title = d.data.title || '';
        const titleElement = document.createElementNS('http://www.w3.org/2000/svg', 'text');
        titleElement.setAttribute('class', 'node-title');
        titleElement.setAttribute('x', textX);
        titleElement.setAttribute('y', 5);
        titleElement.setAttribute('text-anchor', textAnchor);
        
        // Manual text wrapping for title
        const words = title.split(' ');
        let currentLine = '';
        let lineCount = 0;
        const maxLines = 2;
        for (let i = 0; i < words.length; i++) {
            const testLine = currentLine ? `${currentLine} ${words[i]}` : words[i];
            // Simple length check, not perfect but avoids complex measurement
            if (testLine.length > (textWidth / 6) && lineCount < maxLines - 1) {
                const tspan = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
                tspan.setAttribute('x', textX);
                tspan.setAttribute('dy', `${lineCount === 0 ? 0 : 1.2}em`);
                tspan.textContent = currentLine;
                titleElement.appendChild(tspan);
                currentLine = words[i];
                lineCount++;
            } else {
                currentLine = testLine;
            }
        }
        const lastTspan = document.createElementNS('http://www.w3.org/2000/svg', 'tspan');
        lastTspan.setAttribute('x', textX);
        lastTspan.setAttribute('dy', `${lineCount === 0 ? 0 : 1.2}em`);
        lastTspan.textContent = currentLine;
        titleElement.appendChild(lastTspan);
    nodeG.appendChild(titleElement);

        // Department
        if (appSettings.showDepartments !== false && d.data.department) {
            const deptText = document.createElementNS('http://www.w3.org/2000/svg', 'text');
            deptText.setAttribute('class', 'node-department');
            deptText.setAttribute('x', textX);
            deptText.setAttribute('y', 28);
            deptText.setAttribute('text-anchor', textAnchor);
            deptText.textContent = d.data.department;
            nodeG.appendChild(deptText);
        }
        
        nodesGroup.appendChild(nodeG);
    }
    
    g.appendChild(linksGroup);
    g.appendChild(nodesGroup);
    exportSvg.appendChild(g);
    return exportSvg;
}

function applyCompactLayout(nodes) {
    const threshold = appSettings.compactLayoutThreshold || 20;
    nodes.forEach(node => {
        const childCount = (node.children || node._children || []).length;

        if (childCount >= threshold) {
            node.data.hasCompactChildren = true; 
            
            const children = node.children;
            if (children) {
                const columns = Math.ceil(Math.sqrt(children.length));
                
                const horizontalSpacing = nodeWidth + 40;
                const verticalSpacing = levelHeight + 20;

                const totalWidth = (columns - 1) * horizontalSpacing;

                children.forEach((child, i) => {
                    child.data.isCompact = true;
                    const col = i % columns;
                    const row = Math.floor(i / columns);

                    if (currentLayout === 'vertical') {
                        child.y = node.y + (row + 1) * verticalSpacing;
                        child.x = node.x - totalWidth / 2 + col * horizontalSpacing;
                    } else { // horizontal
                        child.x = node.x + (row + 1) * verticalSpacing;
                        child.y = node.y - totalWidth / 2 + col * horizontalSpacing;
                    }
                });
            }
        } else if (node.data.hasCompactChildren) {
            delete node.data.hasCompactChildren;
            const allChildren = (node.children || []).concat(node._children || []);
            allChildren.forEach(child => {
                delete child.data.isCompact;
            });
        }
    });
}

// Employee detail functions
function showEmployeeDetailById(employeeId) {
    if (!employeeId) return;
    const employee = employeeById.get(employeeId);
    if (employee) {
        showEmployeeDetail(employee);
    }
}

function initializeAvatarFallbacks(container) {
    if (!container) return;
    container.querySelectorAll('[data-role="avatar-image"]').forEach(img => {
        const fallback = img.nextElementSibling && img.nextElementSibling.matches('[data-role="avatar-fallback"]')
            ? img.nextElementSibling
            : null;
        if (!fallback) return;

        const showFallback = () => {
            fallback.hidden = false;
            img.style.display = 'none';
        };

        img.addEventListener('error', showFallback, { once: true });
        if (img.complete && img.naturalWidth === 0) {
            showFallback();
        }
    });
}

function showEmployeeDetail(employee) {
    if (!employee) return;

    const detailPanel = document.getElementById('employeeDetail');
    const headerContent = document.getElementById('employeeDetailContent');
    const infoContent = document.getElementById('employeeInfo');

    const initials = (employee.name || '')
        .split(' ')
        .map(n => n[0])
        .join('')
        .substring(0, 2)
        .toUpperCase();

    const employeeAvatar = renderAvatar({
        imageUrl: employee.photoUrl && employee.photoUrl.includes('/api/photo/') ? employee.photoUrl : '',
        name: employee.name,
        initials,
        imageClass: 'employee-avatar-image',
        fallbackClass: 'employee-avatar-fallback'
    });

    headerContent.innerHTML = `
        <div class="employee-avatar-container">
            ${employeeAvatar}
        </div>
        <div class="employee-name">
            <h2>${escapeHtml(employee.name)}</h2>
        </div>
        <div class="employee-title">${escapeHtml(employee.title)}</div>
    `;

    let infoHTML = `
        <div class="info-item">
            <div class="info-label">Department</div>
            <div class="info-value">${escapeHtml(employee.department || 'Not specified')}</div>
        </div>
        <div class="info-item">
            <div class="info-label">Email</div>
            <div class="info-value">
                ${employee.email ? `<a href="mailto:${escapeHtml(employee.email)}">${escapeHtml(employee.email)}</a>` : 'Not available'}
            </div>
        </div>
        <div class="info-item">
            <div class="info-label">Phone</div>
            <div class="info-value">${escapeHtml(employee.phone || 'Not available')}</div>
        </div>
        ${employee.hireDate ? `
        <div class="info-item">
            <div class="info-label">Hire Date</div>
            <div class="info-value">${escapeHtml(formatHireDate(employee.hireDate))}</div>
        </div>
        ` : ''}
        ${employee.location ? `
        <div class="info-item">
            <div class="info-label">Office</div>
            <div class="info-value">${escapeHtml(employee.location)}</div>
        </div>
        ` : ''}
        ${employee.city || employee.state || employee.country ? `
        <div class="info-item">
            <div class="info-label">Location</div>
            <div class="info-value">${[employee.city, employee.state, employee.country].filter(Boolean).map(escapeHtml).join(', ')}</div>
        </div>
        ` : ''}
    `;

    if (employee.managerId && window.currentOrgData) {
        const manager = findManagerById(window.currentOrgData, employee.managerId);
        if (manager) {
            const managerInitials = (manager.name || '')
                .split(' ')
                .map(n => n[0])
                .join('')
                .substring(0, 2)
                .toUpperCase();

            const managerAvatar = renderAvatar({
                imageUrl: manager.photoUrl && manager.photoUrl.includes('/api/photo/') ? manager.photoUrl : '',
                name: manager.name,
                initials: managerInitials,
                imageClass: 'manager-avatar-image',
                fallbackClass: 'manager-avatar-fallback'
            });

            infoHTML += `
                <div class="manager-section">
                    <h3>Manager</h3>
                    <div class="manager-item" data-employee-id="${escapeHtml(manager.id)}">
                        <div class="manager-avatar-container">
                            ${managerAvatar}
                        </div>
                        <div class="manager-details">
                            <div class="manager-name">${escapeHtml(manager.name)}</div>
                            <div class="manager-title">${escapeHtml(manager.title)}</div>
                        </div>
                    </div>
                </div>
            `;
        }
    }

    const directReports = employee.children || [];
    if (directReports.length > 0) {
        infoHTML += `
            <div class="direct-reports">
                <h3>Direct Reports (${directReports.length})</h3>
                ${directReports.map(report => {
                    const reportInitials = (report.name || '')
                        .split(' ')
                        .map(n => n[0])
                        .join('')
                        .substring(0, 2)
                        .toUpperCase();

                    const reportAvatar = renderAvatar({
                        imageUrl: report.photoUrl && report.photoUrl.includes('/api/photo/') ? report.photoUrl : '',
                        name: report.name,
                        initials: reportInitials,
                        imageClass: 'report-avatar-image',
                        fallbackClass: 'report-avatar-fallback'
                    });

                    return `
                        <div class="report-item" data-employee-id="${escapeHtml(report.id)}">
                            <div class="report-avatar-container">
                                ${reportAvatar}
                            </div>
                            <div class="report-details">
                                <div class="report-name">${escapeHtml(report.name)}</div>
                                <div class="report-title">${escapeHtml(report.title)}</div>
                            </div>
                        </div>
                    `;
                }).join('')}
            </div>
        `;
    }

    infoContent.innerHTML = infoHTML;
    initializeAvatarFallbacks(detailPanel);
    detailPanel.classList.add('active');
}

function closeEmployeeDetail() {
    document.getElementById('employeeDetail').classList.remove('active');
}

function findNodeById(node, targetId) {
    if (node.data.id === targetId) {
        return node;
    }
    if (node.children || node._children) {
        const children = node.children || node._children;
        for (let child of children) {
            const result = findNodeById(child, targetId);
            if (result) return result;
        }
    }
    return null;
}

function findManagerById(rootData, managerId) {
    function searchNode(node) {
        if (node.id === managerId) {
            return node;
        }
        if (node.children) {
            for (let child of node.children) {
                const result = searchNode(child);
                if (result) return result;
            }
        }
        return null;
    }
    return searchNode(rootData);
}

function highlightNode(nodeId, highlight = true) {
    if (appSettings.searchHighlight !== false) {
        g.selectAll('.node-rect').each(function(d) {
            if (d.data.id === nodeId) {
                d3.select(this).classed('search-highlight', highlight);
            }
        });
    }
}

function clearHighlights() {
    g.selectAll('.node-rect').classed('search-highlight', false);
}

let searchTimeout;
const searchInput = document.getElementById('searchInput');
const searchResults = document.getElementById('searchResults');

searchInput.addEventListener('input', function(e) {
    clearTimeout(searchTimeout);
    const query = e.target.value.trim();
    
    clearHighlights();
    
    if (query.length < 2) {
        searchResults.classList.remove('active');
        return;
    }
    
    searchTimeout = setTimeout(() => {
        performSearch(query);
    }, 300);
});

searchInput.addEventListener('focus', function(e) {
    if (e.target.value.length >= 2) {
        performSearch(e.target.value);
    }
});

document.addEventListener('click', function(e) {
    if (!e.target.closest('.search-wrapper')) {
        searchResults.classList.remove('active');
    }
});

async function performSearch(query) {
    try {
        const response = await fetch(`${API_BASE_URL}/api/search?q=${encodeURIComponent(query)}`);
        const results = await response.json();
        
        if (results.length > 0) {
            displaySearchResults(results);
        } else {
            searchResults.innerHTML = '<div class="search-result-item">No results found</div>';
            searchResults.classList.add('active');
        }
    } catch (error) {
        console.error('Search error:', error);
    }
}

function displaySearchResults(results) {
    searchResults.innerHTML = '';

    results.forEach(emp => {
        const item = document.createElement('div');
        item.className = 'search-result-item';
        item.dataset.employeeId = emp.id;

        const name = document.createElement('div');
        name.className = 'search-result-name';
        name.textContent = emp.name || '';

        const title = document.createElement('div');
        title.className = 'search-result-title';
        const departmentText = emp.department ? ` – ${emp.department}` : '';
        title.textContent = `${emp.title || 'No Title'}${departmentText}`;

        item.appendChild(name);
        item.appendChild(title);
        searchResults.appendChild(item);
    });
    searchResults.classList.add('active');
}

function selectSearchResult(employeeId) {
    const employee = employeeById.get(employeeId);
    if (employee) {
        showEmployeeDetail(employee);
        searchResults.classList.remove('active');
        searchInput.value = '';
        
        expandToEmployee(employeeId);
    }
}

function expandToEmployee(employeeId) {
    if (appSettings.searchAutoExpand === false) {
        const targetNode = findNodeById(root, employeeId);
        if (targetNode) {
            highlightNode(employeeId);
            showEmployeeDetail(targetNode.data);
        }
        return;
    }
    
    const path = [];
    
    function findPath(node, targetId, currentPath) {
        currentPath.push(node);
        
        if (node.data.id === targetId) {
            path.push(...currentPath);
            return true;
        }
        
        if (node.children || node._children) {
            const children = node.children || node._children;
            for (let child of children) {
                if (findPath(child, targetId, [...currentPath])) {
                    return true;
                }
            }
        }
        
        return false;
    }
    
    findPath(root, employeeId, []);
    
    path.forEach(node => {
        if (node._children) {
            node.children = node._children;
            node._children = null;
        }
    });
    
    update(root);
    
    const targetNode = path[path.length - 1];
    if (targetNode) {
        setTimeout(() => {
            highlightNode(employeeId);
        }, 600);
        
        const container = document.getElementById('orgChart');
        const width = container.clientWidth;
        const height = container.clientHeight;
        
        svg.transition()
            .duration(750)
            .call(zoom.transform, 
                d3.zoomIdentity
                    .translate(width/2, height/2)
                    .scale(1)
                    .translate(-targetNode.x, -targetNode.y)
            );
    }
}

// Export to XLSX function
async function exportToXLSX() {
    try {
        const response = await fetch(`${API_BASE_URL}/api/export-xlsx`);
        
        if (response.ok) {
            // Get the blob
            const blob = await response.blob();
            
            // Create download link
            const url = window.URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            
            // Get filename from Content-Disposition header or use default
            const contentDisposition = response.headers.get('Content-Disposition');
            let filename = `org-chart-${new Date().toISOString().split('T')[0]}.xlsx`;
            if (contentDisposition && contentDisposition.includes('filename=')) {
                filename = contentDisposition.split('filename=')[1].replace(/"/g, '');
            }
            
            link.download = filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            // Clean up
            window.URL.revokeObjectURL(url);
        } else {
            const errorData = await response.json();
            alert(`Export failed: ${errorData.error || 'Unknown error'}`);
        }
    } catch (error) {
        console.error('Export error:', error);
        alert('Failed to export XLSX file');
    }
}

// Logout function
async function logout() {
    try {
        const response = await fetch('/logout', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            }
        });
        
        if (response.ok) {
            window.location.href = '/';
        }
    } catch (error) {
        console.error('Logout error:', error);
        // Force redirect even if request fails
        window.location.href = '/';
    }
}

function registerEventHandlers() {
    setupStaticEventListeners();
}

window.addEventListener('resize', () => {
    updateSvgSize();
    if (userAdjustedZoom) return;
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => {
        fitToScreen({ duration: 300, resetUser: true });
    }, RESIZE_DEBOUNCE_MS);
});

document.addEventListener('DOMContentLoaded', () => {
    registerEventHandlers();
    init();
});

document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
        closeEmployeeDetail();
        clearHighlights();
    }
});