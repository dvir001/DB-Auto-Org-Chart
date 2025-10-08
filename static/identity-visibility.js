(function(global) {
    function normalize(value) {
        return value !== false;
    }

    function enforceIdentityVisibility(state) {
        const current = {
            names: normalize(state && state.names),
            titles: normalize(state && state.titles),
            departments: normalize(state && state.departments)
        };

        if (!current.names && !current.titles && !current.departments) {
            current.names = true;
        }

        return current;
    }

    const api = {
        enforceIdentityVisibility
    };

    if (global) {
        if (global.identityVisibility) {
            Object.assign(global.identityVisibility, api);
        } else {
            global.identityVisibility = api;
        }
    }

    if (typeof module !== 'undefined' && module.exports) {
        module.exports = api;
    }
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
