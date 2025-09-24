## What is DB-AutoOrgChart?

> **Note:** This repository is a maintained fork of [jaffster595/DB-Auto-Org-Chart](https://github.com/jaffster595/DB-Auto-Org-Chart).

**Key additions in this fork**
- Hardened security posture: stricter Content-Security-Policy (`script-src 'self'`), dedicated login template + JS, sanitized redirect handling, and protections against placeholder secrets.
- Frontend refactor: all inline scripts/styles removed, `configureme` and `search_test` now load modular JS/CSS bundles with shared CSS variables and safer DOM updates.
- Export & print tooling: one-click exports for the visible chart in SVG/PNG/PDF, a server-backed `/api/export-xlsx` endpoint wired to `Export to XLSX` UI controls, and a print-optimised window.
- Discovery & filtering perks: hide/show subtrees per user with persistent local storage, quick reset of hidden teams, enriched search/top-level selection helpers, configurable filters to drop disabled or guest accounts from the dataset by default, and the Compact Large Teams toggle available to every viewer.
- Operational polish: unused static assets trimmed, scheduler + data directories validated on startup, and logging made more actionable for Azure Graph interactions.

DB-AutoOrgChart is an application which connects to your Azure AD/Entra via Graph API, retrieves the appropriate information (employee name, title, department, 'reports to' etc.) then builds an interactive Organisation Chart based upon that information. It can be run as an App Service in Azure / Google Cloud etc or you can run it locally. NOTE: You will need the appropriate permissions in Azure to set up Graph API which is a requirement for this application to function. You only need to do this once, so someone with those permissions can set it up for you then provide the environment variables to you.

In short, these are the main features of DB-AutoOrgChart:

- Automatic organization hierarchy generation from Azure AD
- Real-time employee search functionality (employee directory)
- Completely configurable with zero coding knowledge via configureme page
- Interactive D3.js-based org chart with zoom and pan
- Detailed employee information panel
- Print-friendly org chart export
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

The supplied `Dockerfile` builds a production image that runs as an unprivileged `app` user. Runtime dependencies are installed during the image build so no `pip install` happens on container start.

```bash
docker compose up --build -d
```
