# PowerShell Reporting & Automation Framework

A modular, config-driven PowerShell framework that runs scheduled reporting
and automation jobs — via Azure DevOps Pipelines — against Azure/Entra ID,
Microsoft 365, network/security platforms, and business systems, with
results uploaded automatically to Azure Blob Storage.

## 🔧 Features

- Modular script structure — one script per report/job
- Config-driven orchestrator (`run_all.ps1` + `scripts/config.json`)
- Wide library of ready-made reports covering:
  - Azure / Entra ID (users, groups, dynamic groups, Conditional Access,
    Defender for Cloud alerts, applications, subscriptions, VNets, DNS,
    public IPs, cost-center tagging, hybrid AD-to-Entra ID group sync)
  - Microsoft 365 (license assignment matrix, Teams PSTN users)
  - Network & security platforms (Palo Alto Panorama, Meraki, DigiCert
    certificate monitoring)
  - Business systems (Workday, Azure DevOps repos & work items)
  - Password/access governance (LastPass Enterprise shared folder audit)
- Upload results (CSV) to Azure Blob Storage, with folder nesting and overwrite support
- Scheduled via Azure DevOps Pipelines (cron-based)

## 📁 Project Structure

```
/
├── azure-pipelines.yml          # Azure DevOps Pipeline (scheduled trigger)
├── run_all.ps1                  # Main orchestrator script
├── scripts/
│   ├── *.ps1                    # Individual report/automation scripts
│   ├── config.json              # Job schedule/parameters (sanitized example — see below)
│   └── config_template.json     # Minimal starter template
└── utils/
    └── upload.ps1                # Uploads a file to Azure Blob Storage
```

## 🚀 How It Works

1. Each report script accepts parameters (typically via environment variables
   for secrets) and writes a CSV to an output path.
2. `run_all.ps1` reads `scripts/config.json`, logs in to Azure with a service
   principal, and runs each configured job in turn.
3. The result of each job is uploaded to the configured Blob Storage folder
   via `utils/upload.ps1`.

## ⚠️ Before you use this repo

This repo is a **template**. `scripts/config.json` ships with placeholder
values (`yourcompany.com`, `YourOrg`, zeroed-out GUIDs, generic user emails)
so you can see the intended shape of a real configuration — replace them with
your own tenant IDs, organization names, and target users before running
anything.

**No credentials, tokens, or API keys are included anywhere in this repo.**
Every script authenticates using environment variables — see below.

## 🔑 Required Environment Variables

Depending on which report scripts you use, set the following as pipeline
secrets (Azure DevOps variable group) or local environment variables:

| Variable | Used for |
|---|---|
| `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`, `AZURE_TENANT_ID` | Azure AD app / service principal auth (Graph API + Az PowerShell) |
| `AZURE_SUBSCRIPTION_ID` | Optional — scope Azure Resource Graph queries to one subscription |
| `AZURE_STORAGE_ACCOUNT`, `AZURE_STORAGE_CONTAINER` | Blob Storage upload target |
| `PA_API_KEY` | Palo Alto / Panorama API key |
| `Dcert_API_KEY` | DigiCert API key |
| `MerakiApiKey` | Cisco Meraki Dashboard API key |
| `WORKDAY_USER`, `WORKDAY_PASS` | Workday API credentials |
| `PAT` | Azure DevOps Personal Access Token |
| `AzureOpenAIKey` | Azure OpenAI key (used by `devops_repos_info.ps1` for AI-generated repo summaries) |

None of these are set in the repo — set them as pipeline secrets (Azure
DevOps → Pipelines → Library → Variable groups) or as local environment
variables for testing.

## 🔧 Azure Setup

- Create a Storage Account and container for report output.
- Register an Azure AD App / service principal with the permissions each
  script needs (mostly Microsoft Graph delegated/application permissions —
  see comments at the top of each script).
- Store all secrets above as pipeline secrets and set up a Service Connection
  in Azure DevOps.

## 📅 Schedule Configuration (`scripts/config.json`)

```json
[
  {
    "name": "Example report",
    "script": "scripts/report_template.ps1",
    "outputName": "example_report_$(Date).csv",
    "parameters": {
      "StartDate": "2023-01-01",
      "EndDate": "2023-01-31"
    },
    "targetFolder": "Reports/Example"
  }
]
```

## 📝 Notes

- The blob upload supports nested folder creation and file overwrite.
- Use `$(Date)` in `outputName` to auto-generate filenames.
- Extend with additional report scripts as needed — each one just needs to
  accept parameters and write a CSV to `-OutputPath`.

---

© 2026 by Vasyl Volyk — MIT License
