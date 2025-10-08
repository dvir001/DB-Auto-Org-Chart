## What is DB-AutoOrgChart?

> **Note:** This repository is a maintained fork of [jaffster595/DB-Auto-Org-Chart](https://github.com/jaffster595/DB-Auto-Org-Chart).

**Key additions in this fork**
- Hardened security posture: stricter Content-Security-Policy (`script-src 'self'`), dedicated login template + JS, sanitized redirect handling, and protections against placeholder secrets.
- Frontend refactor: all inline scripts/styles removed, `configure` and `search_test` now load modular JS/CSS bundles with shared CSS variables and safer DOM updates.
- Export & print tooling: one-click exports for the visible chart in SVG/PNG/PDF, a server-backed `/api/export-xlsx` endpoint wired to `Export to XLSX` UI controls, and a print-optimised window.
- Discovery & filtering perks: hide/show subtrees per user with persistent local storage, quick reset of hidden teams, enriched search/top-level selection helpers, configurable filters to drop disabled or guest accounts from the dataset by default, and the Compact Teams toggle available to every viewer.
- Admin insight upgrades: refreshed reporting dashboard with missing-manager, disabled-but-licensed, and filter-hidden licensed user summaries plus one-click XLSX exports.
- Operational polish: unused static assets trimmed, scheduler + data directories validated on startup, and logging made more actionable for Azure Graph interactions.

DB-AutoOrgChart is an application which connects to your Azure AD/Entra via Graph API, retrieves the appropriate information (employee name, title, department, 'reports to' etc.) then builds an interactive Organisation Chart based upon that information. It can be run as an App Service in Azure / Google Cloud etc or you can run it locally. NOTE: You will need the appropriate permissions in Azure to set up Graph API which is a requirement for this application to function. You only need to do this once, so someone with those permissions can set it up for you then provide the environment variables to you.

In short, these are the main features of DB-AutoOrgChart:

- Automatic organization hierarchy generation from Azure AD
- Real-time employee search functionality (employee directory)
- Completely configurable with zero coding knowledge via configure page
- Interactive D3.js-based org chart with zoom and pan
- Detailed employee information panel
- Print-friendly org chart export
- Admin reporting dashboard with refreshable XLSX exports (missing managers, disabled accounts holding licenses, filtered licensed users)
- Automatic daily updates at 20:00
- Color-coded hierarchy levels
- Responsive design for mobile and desktop

  <img width="1640" height="527" alt="image" src="https://github.com/user-attachments/assets/f33719e6-cc03-40bc-89fc-72d9e0f58674" />

It makes one API call per day at 20:00 and saves the acquired data within employee_data.json, which sits securely within the app service. When someone visits the org chart, it displays the Org data based upon the contents of employee_data.json rather than making constant API calls via Graph API. This way it makes a single API request, once per day, rather than making constant requests each time someone visit the page. This not only reduces the amount of traffic caused by this application, but also makes it faster and more responsive.

## Prerequisites
1. Docker Desktop (or Docker Engine with Compose plugin)
2. An Azure AD tenant with appropriate permissions to:
  a. Register an App in Azure for Graph API
  b. Create an App Service in Azure (optional but recommended)
3. Azure AD App Registration with User.Read.All permissions and admin consent (see Azure AD setup below)

## Azure AD setup

1. **Create an App Registration in Azure AD**
  - Go to Azure Portal ➜ Azure Active Directory ➜ App registrations
  - Click **New registration**
  - Name your app (e.g., “DB AutoOrgChart”)
  - No redirect URL is needed

2. **Configure API permissions**
  - In your app registration, open **API permissions**
  - Add a permission ➜ **Microsoft Graph** ➜ **Application permissions**
  - Select **User.Read.All**
  - Select **LicenseAssignment.Read.All** *(required for the reporting dashboard to list licensed users)*
  - Click **Grant admin consent** (requires admin privileges)

3. **Create a client secret**
  - Open **Certificates & secrets**
  - Click **New client secret** and choose an expiration period

4. **Capture the required IDs**
  - From the overview page, copy the **Application (client) ID** → `AZURE_CLIENT_ID`
  - Copy the **Directory (tenant) ID** → `AZURE_TENANT_ID`

## Environment variables

Open `.env.template`, save a new copy named `.env`, and populate it with your values. The application will refuse to start unless the required entries are set to strong, non-default values.

**Required values**

- `AZURE_TENANT_ID`: Directory (tenant) ID of your Azure AD tenant.
- `AZURE_CLIENT_ID`: Application (client) ID of your App Registration.
- `AZURE_CLIENT_SECRET`: Client secret created for the App Registration.
- `TOP_LEVEL_USER_EMAIL`: Email address of the most senior user (root of the org chart).
- `ADMIN_PASSWORD`: Strong password used to secure the `/configure` route.
- `SECRET_KEY`: 64+ character random string for Flask session protection. Generate one with:

  ```bash
  python -c "import secrets; print(secrets.token_hex(32))"
  ```

**Optional values**

- `TOP_LEVEL_USER_ID`: Explicit Graph object ID of the top-level user (if you know it).
- `CORS_ALLOWED_ORIGINS`: Comma-separated list of additional origins permitted to call the API (leave blank to restrict to same-origin requests).

For the Azure credentials, follow the steps in the "Azure AD setup" section. To obtain the top-level user details, locate the chosen user in Azure AD/Entra and copy the "Mail" and "Object (user) ID" fields from their profile.

## Docker / container deployment

The project ships with a Compose stack that pulls the published image `ghcr.io/dvir001/db-auto-org-chart:latest`, mounts a persistent volume for application data, and loads configuration from `.env`.

1. Copy the template and fill in your secrets:
  ```bash
  cp .env.template .env
  # edit .env with your tenant/client IDs, secrets, and admin password
  ```

2. Start (or update) the service:
  ```bash
  docker compose pull
  docker compose up -d
  ```

  The container exposes port `5000` by default—adjust the `ports` in `docker-compose.yml` if you need a different host port.

3. Persistent data lives in the named volume `orgchart_data`, which retains cached employee hierarchies and settings between restarts. Remove that volume if you want a clean slate.

To deploy on a new host, drop `docker-compose.yml` and your populated `.env` into a directory, then run the commands above, the stack will pull the image, provision the named volume, and start the app with no other files required.
