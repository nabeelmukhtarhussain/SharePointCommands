# SharePoint Document Library Assessment Tool

A single PowerShell script that scans a Microsoft 365 tenant and reports **every SharePoint document library** — its address, size, the parent site's visibility and sharing status — and exports a clean, ready-to-use **Excel report**.

Built for one job: sizing a **SharePoint document-library migration** without scanning the whole tenant.

---

## What it gives you

An Excel workbook (`LibraryReport.xlsx`) with these sheets:

| Sheet | What's in it |
|-------|--------------|
| **All Libraries** | Site, URL, Visibility, External Sharing, Library, Address, Item count, Size (GB), Category |
| **Under 50GB** | Libraries below 50 GB |
| **50-100GB** | Libraries between 50 and 100 GB |
| **Over 100GB** | Libraries above 100 GB |
| **Summary** | Totals and counts at a glance |

---

## Prerequisites

- **A SharePoint Administrator account** (to read every site).
- **Windows** with PowerShell. PowerShell 7 is required, but you don't need to install it yourself — the script installs it automatically if it's missing.
- **Internet access** (to install the modules and PowerShell 7 the first time).

> Everything else — PowerShell 7, the `PnP.PowerShell` and `ImportExcel` modules, and the sign-in app — is set up automatically by the script.

---

## How to use it (3 easy steps)

**1. Open PowerShell as Administrator**
Start menu → type `PowerShell` → right-click → **Run as administrator**.
*(Admin rights are only needed so it can install PowerShell 7 the first time.)*

**2. Paste the whole script and press Enter**
Copy the entire contents of `Get-SPODocLibraryReport.ps1` and paste it into the window.
The script takes over from here — it will switch to PowerShell 7, install what it needs, and walk you through sign-in.

**3. Sign in when asked**
- Enter your admin email (e.g. `admin@contoso.onmicrosoft.com`), then press **Enter** to accept the detected tenant.
- Sign in (and, on the very first run for a tenant, **consent** to the app).

That's it. When it finishes, you'll find **`LibraryReport.xlsx`** in the folder you ran it from.

---

## Things to keep in mind

- **First run = one consent.** The first time you run it against a tenant, you approve a sign-in app once. This is a Microsoft security step. A **Global Administrator** may need to approve it. After that, it's remembered (cached per tenant) and never asked again.
- **It grants you Site Collection Admin on every site.** This is what removes the usual "no access" gaps so no library is skipped. It's standard practice for migration reporting, but worth knowing.
- **Sizes come from SharePoint's storage metrics** — fast and accurate enough for migration planning, though they can lag real-time by a short while.
- **Visibility (Public/Private)** needs `Group.Read.All` consent. Without it, that column is left blank — **External Sharing still works**.
- **Close the report before re-running.** If `LibraryReport.xlsx` is open in Excel, the script automatically saves to a new timestamped file instead of failing.
- **Large tenants take longer.** It connects to each site one by one — expect a few minutes on big environments.

---

## FAQ

**Do I need to type the tenant name?**
No — it's detected from your admin email. You only override it for custom/vanity domains.

**Will it move or change my data?**
No. It only **reads** library sizes and **adds you as site admin** for access. It does not migrate, delete, or modify content.

**Can I run it on multiple tenants?**
Yes. Each tenant is cached separately, so you can run it across source and destination without conflicts.

---

*Made for IT admins and migration consultants. Use it, share it, improve it.*
