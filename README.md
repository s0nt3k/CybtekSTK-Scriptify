# Scriptify (ODT Utility)
**Cybtek STK – Office Deployment Tool Front‑End (PowerShell 5.1, WPF)**

> A friendly, click‑driven way to download, configure, and install Microsoft Office using the official Office Deployment Tool (ODT). Built for small businesses—and especially real‑estate teams—who need predictable, low‑friction Office installs.

---

## Why Scriptify?
Small businesses don’t have time to memorize XML options or dig through setup switches. Scriptify wraps Microsoft’s **Office Deployment Tool** in a clean Windows interface so you can:
- **Choose products/editions** (Office 2016/2019/2021/2024 and Microsoft 365), plus toggle **Project** and **Visio**.
- **Pick an update channel** (Current, Monthly Enterprise, Semi‑Annual, previews, and Perpetual VL tracks).
- **Select languages** (with a simple picker).
- **Include/exclude apps** like Word, Excel, Outlook (classic), PowerPoint, OneNote, OneDrive variants, etc.
- **Generate the ODT configuration XML** for you—no hand‑editing needed.
- **Download** Office media or **Install** directly.
- **Review logs** and save default settings for repeatable deployments.
- **Run common repairs** (Quick/Full via OfficeClickToRun) and kick off **PST repair** (SCANPST).

> Designed with non‑technical owners and managers in mind: fewer choices, clearer defaults, and safer guardrails.

---

## Features at a Glance
- ✅ PowerShell **5.1** compatible (Windows) with WPF UI
- ✅ Auto‑download/extract of ODT to `%AppData%\Microsoft\ODT`
- ✅ Configuration XML generator (channel, version, languages, app selections)
- ✅ **Download** or **Install** actions via `setup.exe`
- ✅ Optional **system restore point** before install (toggle in Settings)
- ✅ **Remove previous MSI/Click‑to‑Run** Office installs (guided prompt)
- ✅ **Logs** to `.\logs\download.json` and `.\logs\install.json`
- ✅ **Defaults** saved to `.\odt-defaults.json`
- ✅ **View Script** preview of the generated XML
- ✅ Licensing helper window (technician/company info saved into the script’s license block)

---

## Prerequisites
- Windows 10/11 (or Windows Server with desktop experience)
- PowerShell **5.1**
- .NET/WPF components available (standard on supported Windows SKUs)
- Internet access to download the Office Deployment Tool & (optionally) Office media

> **Execution Policy:** If your machine blocks script execution, you can run Scriptify in the current session without changing system‑wide policy:
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
```

> **Unblock the file** if downloaded from the internet:
```powershell
Unblock-File .\scriptify.ps1
```

---

## Getting Started
1. **Clone or download** this repository.
2. Open **Windows PowerShell** (5.1).
3. Run the utility:
   ```powershell
   .\scriptify.ps1
   ```
   The UI will download/extract the **Office Deployment Tool** on first launch to `%AppData%\Microsoft\ODT` and create working folders:
   - `.\generated` – generated XML configs
   - `.\logs` – JSON logs for download/install

---

## Quick‑Start (Typical Install)
1. **Office Product & Edition:** Choose the product (e.g., *MS Office 2021*) and an edition (e.g., *Professional Plus (Retail)*).  
   - Use **Settings → Show Volume Editions** if you need VL SKUs.
2. **Servicing Channel:** Pick **Current**, **Monthly Enterprise**, **Semi‑Annual**, or a preview/perpetual track.
3. **Version:** Leave **Latest Available** or switch to **Specific Version** and enter a build.
4. **Languages:** **Settings → Add Languages…** and select one or more locales (e.g., *en‑us, es‑es*).
5. **Applications:** **Settings → Select Applications…** to turn apps on/off (unchecked apps are excluded in XML).
6. Optional toggles under **Settings**:
   - Install 32‑Bit, Display Level (None), Force App Close, Disable Restore Point, Disable Remove Office
   - Install Microsoft Project, Install Microsoft Visio
7. Click **Download** to fetch media to a folder, or **Install** to run the deployment now.

> Use **File → View Script** to preview the exact XML that will be handed to ODT.

---

## Update Channels (Plain‑English)
- **Current** – Fastest retail track; features as soon as they’re ready.
- **MonthlyEnterprise** – Predictable once‑per‑month release; popular for most small businesses.
- **SemiAnnual** – Slower, stability‑focused (feature updates twice a year).
- **CurrentPreview / SemiAnnualPreview** – Early previews of each track for testing.
- **BetaChannel** – Early access builds; for testing labs only.
- **PerpetualVL2019 / PerpetualVL2021** – Tracks for perpetual Volume License Office.

---

## What the Generated XML Looks Like
Scriptify builds a standards‑compliant ODT configuration. Example (simplified):
```xml
<Configuration>
  <Add OfficeClientEdition="64" Channel="MonthlyEnterprise" SourcePath="D:\OfficeMedia">
    <Product ID="ProPlus2021Retail">
      <Language ID="en-us" />
      <Language ID="es-es" />
    </Product>
  </Add>

  <Display Level="Full" AcceptEULA="TRUE" />
  <Updates Enabled="TRUE" Channel="MonthlyEnterprise" />

  <!-- Excluded apps appear as ExcludeApp elements -->
  <ExcludeApp ID="OneDrive" />
</Configuration>
```

---

## Logs & Defaults
- **Logs**:  
  - `.\logs\download.json` and `.\logs\install.json` capture timestamp, host and selection metadata.  
  - Open via **View → Download Log / Install Log** and double‑click any row for details.
- **Defaults**:
  - Save your preferred configuration via **Tools → Set Defaults** (saved to `.\odt-defaults.json`).  
  - Clear via **Tools → Clear Defaults**.

---

## Repairs & Maintenance
- **Repair**: **Repair → Quick Repair / Full Repair** launches `OfficeClickToRun.exe` with suitable parameters.
- **PST Repair**: **Repair → PST Repair…** animates progress then starts **SCANPST.EXE** if available.
- **Remove Previous Installs**: **Tools → Remove Office** creates a removal XML and runs ODT to remove MSI and Click‑to‑Run Office.

> Scriptify will attempt to create a **system restore point** before install (unless disabled in Settings).

---

## Security, Privacy & Data
- No telemetry is sent by Scriptify. Logs are local JSON files under `.\logs`.
- The **License** helper window stores technician/company info in a small, embedded block **inside the script file** (for receipts/records). You can clear it with **Help → License** and entering `RESETLIC`.
- Be aware that running installs will **close Office apps** if needed (you can toggle “Force App Close” in Settings). Always save work first.

---

## Compliance & Licensing (Important)
Scriptify uses Microsoft’s **Office Deployment Tool** and **OfficeClickToRun**. You are responsible for:
- **Using proper licenses** (Retail vs. Volume).  
- **Assigning Microsoft 365 licenses** to users before installing apps.  
- **Not mixing incompatible editions** on the same device.

**Potential consequences of non‑compliance** (typical outcomes, not legal advice):
- Microsoft or reseller **audit/true‑up** resulting in back‑licensing charges.  
- **Penalties/fees** specified in your Volume Licensing agreement.  
- **Termination or suspension** of licensing services in severe cases.  
- For organizations handling consumer data (e.g., brokerages), deploying unlicensed/unsupported software can increase liability exposure in disputes or insurance claims.

> When in doubt, verify your SKU entitlements and edition/channel rules with your Microsoft agreement or partner.

---

## Troubleshooting
- **“ODT.exe is missing”** – First launch should download/extract ODT to `%AppData%\Microsoft\ODT`. Check connectivity or rerun Scriptify.
- **Execution Policy** – Use the Process‑scoped bypass shown above, or sign the script per your policy.
- **Checkpoint‑Computer not found** – Some SKUs or disabled services will skip restore point creation (Scriptify handles this gracefully).
- **OfficeClickToRun.exe not found** – Use an elevated “Microsoft Office” command prompt or ensure Office Click‑to‑Run is installed.
- **SCANPST.EXE not found** – Installed with Outlook. The PST repair window will still guide you; install Outlook tools if needed.
- **Access denied / UAC prompts** – For installs/removals, run PowerShell **as Administrator**.

---

## Contributing
Issues and PRs are welcome. Please include:
- Repro steps, OS/PowerShell version, and any log excerpts (sanitize sensitive info).  
- For UI changes, attach a short screen recording or screenshots if possible.

---

## Roadmap / Ideas
- Optional **silent mode** wrapper for zero‑touch installs
- Configuration **import/export** from existing XML
- **Channel/version discovery** helper
- Signed releases

---

## License (MIT)

MIT License

Copyright (c) 2025 Sonny Gibson

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

## Acknowledgements
- Microsoft **Office Deployment Tool** and **OfficeClickToRun**  
- Community feedback from small business owners and real‑estate teams who shaped this UI

---

**Made with ❤️ for busy teams who just want Office to install the right way the first time.**
