# hsdes_bug_status_transition
Python script to export HSD-ES API article data and bug status transition analysis to Excel.

---

## Prerequisites

### 1. Python 3.x
Download and install from https://python.org
During install, check **"Add Python to PATH"**

### 2. Intel Network / VPN
Must be connected to Intel network (on-site or via Intel VPN) to reach `https://hsdes-api.intel.com/rest/doc/`

### 3. Install required Python packages
Open Command Prompt or PowerShell:
```powershell
pip install requests requests-kerberos openpyxl --proxy http://proxy-iind.intel.com:911
```

### 4. Kerberos Authentication
The script uses Kerberos SSO for authentication.
On a **domain-joined Windows machine** (logged in with Intel credentials)

For more details, refer to the official Intel wiki: [HSD-ES API Authentication](https://wiki.ith.intel.com/pages/viewpage.action?pageId=954011548&spaceKey=HSDESWIKI&title=HSD-ES%2BAPI)

---

## SSL Certificate Setup (Required on fresh Windows systems)

Python uses its own certificate bundle (`certifi`) which does not include Intel's internal CA by default.
Without this fix, you will get:
```
SSL: CERTIFICATE_VERIFY_FAILED] certificate verify failed: self-signed certificate in certificate chain
```

### Fix — Add Intel CA cert to Python's certifi bundle

**Step 1 — Get the Intel CA certificate file:**
Download or locate `IntelSHA256RootCA-Base64.crt` from the [HSD-ES API](https://wiki.ith.intel.com/pages/viewpage.action?pageId=954011548&spaceKey=HSDESWIKI&title=HSD-ES%2BAPI) wiki page.

**Step 2 — Find certifi's cacert.pem location:**
```powershell
python -c "import certifi; print(certifi.where())"
```
Example output:
```
C:\Users\username\AppData\Local\Programs\Python\Python3xx\site-packages\certifi\cacert.pem
```

**Step 3 — Append the Intel cert to certifi's bundle:**
```powershell
Get-Content "C:\path\to\IntelSHA256RootCA-Base64.crt" | Add-Content "C:\path\to\certifi\cacert.pem"
```
Replace both paths with your actual paths from Steps 1 and 2.

**Step 4 — Run the script** — SSL errors should be resolved.

> **Note:** If you upgrade/reinstall Python or the `certifi` package, repeat Step 3 as `cacert.pem` gets overwritten.

---

## Running the Script

```powershell
python.exe fetch_hsdes-api_data.py
```

**Prompts**

Enter the path of **\<Platform Linux bugs.xlsx\>** Downloaded from
[PTL](https://hsdes.intel.com/appstore/generalapps/#/pages/community/101846185?queryId=16026642247&articleId=16027538722), [WCL](https://hsdes.intel.com/appstore/generalapps/#/pages/community/14014596121?queryId=16027752441&…), [NVL](https://hsdes.intel.com/appstore/generalapps/#/pages/community/101846185?queryId=16028339521&articleId=16027880920), [NVL-Hx](https://hsdes.intel.com/appstore/generalapps/#/pages/community/101846185?queryId=16029579904&articleId=16027880920), [ARL-S Ref](https://hsdes.intel.com/appstore/generalapps/#/pages/community/101846185?queryId=16029467971), [ARL-Hx Ref](https://hsdes.intel.com/appstore/generalapps/#/pages/community/101846185?queryId=16029284349) Linux Bug Query dashboard.

Example:
```
Enter the Excel file path containing article IDs (with 'id' column)
Excel file path: <"C:\Users\users\Downloads\NVL-S Linux bugs.xlsx">
```
Enter the platform name (e.g., `WCL`, `PTL-H`, `NVL-S`, `NVL-Hx`, `ARL-S-Ref`, `ARL-Hx-Ref`).
```
Enter platform name: <NVL>
```
Enter the output file name or leave it for default
```
Output filename [nvl_hsdes_export_20260303_174946.xlsx]:
```

The script will generate an Excel file in the same directory with the following sheets:
- `{platform}_data` — Raw article data
- `{platform}_Bugs_transition_graph` — Transition analysis tables (8 sets)
- `{platform}_state_transition_summary` — Consolidated SLA summary table
- `{platform}_status_summary` — Article status distribution

---

## Troubleshooting

| Error | Cause | Fix |
|---|---|---|
| `SSL: CERTIFICATE_VERIFY_FAILED` | Intel CA cert not in Python's bundle | Add cert to certifi (see SSL section above) |
| `ERROR: 401` | Kerberos auth failed | Ensure machine is domain-joined and on Intel network |
| `Skipping article - no data` | API returned empty | Check SSL + Kerberos, verify article ID exists |
| `Network is unreachable` (pip) | No internet / proxy needed | Use `--proxy http://proxy-iind.intel.com:911` with pip |