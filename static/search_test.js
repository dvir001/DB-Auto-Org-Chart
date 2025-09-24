const API_BASE = window.location.origin;

(async function setDynamicTitle() {
    try {
        const res = await fetch(`${API_BASE}/api/settings`);
        if (res.ok) {
            const settings = await res.json();
            if (settings && settings.chartTitle) {
                document.title = `${settings.chartTitle} - Search Test`;
            }
        }
    } catch (error) {
        // ignore title errors
    }
})();

async function checkDataFile() {
    const resultsDiv = document.getElementById('dataFileResults');
    resultsDiv.innerHTML = '<span class="info">Checking data file...</span>';

    try {
        const response = await fetch(`${API_BASE}/api/employees`);
        const data = await response.json();

        if (data && data.name) {
            resultsDiv.innerHTML = `
                <span class="success">✓ Data file exists and is valid</span><br>
                <strong>Root Employee:</strong> ${data.name}<br>
                <strong>Title:</strong> ${data.title || 'N/A'}<br>
                <strong>Children:</strong> ${data.children ? data.children.length : 0}
            `;
        } else {
            resultsDiv.innerHTML = '<span class="error">✗ Data file exists but structure is invalid</span>';
        }
    } catch (error) {
        resultsDiv.innerHTML = `<span class="error">✗ Error checking data: ${error.message}</span>`;
    }
}

async function debugSearch() {
    const resultsDiv = document.getElementById('dataFileResults');
    resultsDiv.innerHTML = '<span class="info">Running debug...</span>';

    try {
        const response = await fetch(`${API_BASE}/api/debug-search`);
        const data = await response.json();

        resultsDiv.innerHTML = `<pre>${JSON.stringify(data, null, 2)}</pre>`;
    } catch (error) {
        resultsDiv.innerHTML = `<span class="error">✗ Debug failed: ${error.message}</span>`;
    }
}

async function forceUpdate() {
    const resultsDiv = document.getElementById('updateResults');
    resultsDiv.innerHTML = '<span class="info">Forcing update... This may take a moment...</span>';

    try {
        const response = await fetch(`${API_BASE}/api/force-update`, {
            method: 'POST'
        });
        const data = await response.json();

        if (data.success) {
            resultsDiv.innerHTML = `<span class="success">✓ ${data.message}</span>`;
        } else {
            resultsDiv.innerHTML = `<span class="error">✗ ${data.message || data.error}</span>`;
            if (data.traceback) {
                resultsDiv.innerHTML += `<pre>${data.traceback}</pre>`;
            }
        }
    } catch (error) {
        resultsDiv.innerHTML = `<span class="error">✗ Update failed: ${error.message}</span>`;
    }
}

async function testSearch() {
    const query = document.getElementById('searchQuery').value;
    const resultsDiv = document.getElementById('searchResults');

    if (query.length < 2) {
        resultsDiv.innerHTML = '<span class="error">Please enter at least 2 characters</span>';
        return;
    }

    resultsDiv.innerHTML = '<span class="info">Searching...</span>';

    try {
        const response = await fetch(`${API_BASE}/api/search?q=${encodeURIComponent(query)}`);

        if (!response.ok) {
            const error = await response.text();
            resultsDiv.innerHTML = `<span class="error">✗ Search returned error: ${response.status}</span><pre>${error}</pre>`;
            return;
        }

        const data = await response.json();

        if (data && data.length > 0) {
            const fragment = document.createDocumentFragment();
            const summary = document.createElement('span');
            summary.className = 'success';
            summary.textContent = `✓ Found ${data.length} results:`;
            fragment.appendChild(summary);
            fragment.appendChild(document.createElement('br'));
            fragment.appendChild(document.createElement('br'));

            data.forEach(emp => {
                const item = document.createElement('div');
                item.className = 'employee-item';
                item.innerHTML = `
                    <strong>${emp.name}</strong><br>
                    ${emp.title || 'No title'}<br>
                    ${emp.department || 'No department'}
                `;
                fragment.appendChild(item);
            });

            resultsDiv.innerHTML = '';
            resultsDiv.appendChild(fragment);
        } else {
            resultsDiv.innerHTML = '<span class="error">No results found</span>';
        }
    } catch (error) {
        resultsDiv.innerHTML = `<span class="error">✗ Search error: ${error.message}</span>`;
    }
}

async function viewAllEmployees() {
    const resultsDiv = document.getElementById('allEmployees');
    resultsDiv.innerHTML = '<span class="info">Loading all employees...</span>';

    try {
        const response = await fetch(`${API_BASE}/api/employees`);
        const data = await response.json();

        resultsDiv.innerHTML = '';
        resultsDiv.appendChild(buildEmployeeList(data));
    } catch (error) {
        resultsDiv.innerHTML = `<span class="error">✗ Error loading employees: ${error.message}</span>`;
    }
}

function buildEmployeeList(node, level = 0) {
    const container = document.createElement('div');
    const indentLevel = Math.min(level, 12);
    container.className = `employee-item indent-${indentLevel}`;

    const nameEl = document.createElement('strong');
    nameEl.textContent = node.name;
    container.appendChild(nameEl);
    container.appendChild(document.createElement('br'));

    const titleText = document.createTextNode(`${node.title || 'No title'}`);
    container.appendChild(titleText);

    if (node.department) {
        container.appendChild(document.createElement('br'));
        container.appendChild(document.createTextNode(node.department));
    }

    if (node.children && node.children.length > 0) {
        node.children.forEach(child => {
            container.appendChild(buildEmployeeList(child, level + 1));
        });
    }

    return container;
}

function registerTestActions() {
    const actionMap = {
        'check-data': checkDataFile,
        'debug-search': debugSearch,
        'force-update': forceUpdate,
        'test-search': testSearch,
        'view-all': viewAllEmployees
    };

    document.querySelectorAll('[data-test-action]').forEach(button => {
        const handler = actionMap[button.dataset.testAction];
        if (typeof handler === 'function') {
            button.addEventListener('click', handler);
        }
    });
}

document.addEventListener('DOMContentLoaded', () => {
    registerTestActions();
    checkDataFile();
});
