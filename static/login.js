document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('loginForm');
    const passwordInput = document.getElementById('password');
    const errorEl = document.getElementById('error');

    if (!form || !passwordInput || !errorEl) {
        return;
    }

    const buildLoginUrl = (nextPage) => {
        if (!nextPage) {
            return '/login';
        }
        return `/login?next=${encodeURIComponent(nextPage)}`;
    };

    form.addEventListener('submit', async (event) => {
        event.preventDefault();
        errorEl.textContent = '';

        const password = passwordInput.value;
        const nextPage = form.dataset.next || '';
        const loginUrl = buildLoginUrl(nextPage);

        try {
            const response = await fetch(loginUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ password })
            });

            const result = await response.json().catch(() => ({}));

            if (response.ok) {
                const redirectTarget = result.next || 'configure';
                window.location.href = `/${redirectTarget}`;
                return;
            }

            errorEl.textContent = result.error || 'Invalid password';
        } catch (err) {
            console.error('Login request failed', err);
            errorEl.textContent = 'Login failed';
        }
    });
});
