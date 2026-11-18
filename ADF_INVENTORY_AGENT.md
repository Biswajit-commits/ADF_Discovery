# Azure Data Factory Inventory Agent

A Python-based agent that connects to an Azure Data Factory instance via Service Principal authentication (OAuth 2.0 client credentials), retrieves a comprehensive metadata inventory of all ADF resources, classifies each resource by complexity (SIMPLE / MEDIUM / COMPLEX), and exports the results to a professionally styled, multi-sheet Excel workbook.

> **No data is accessed** — the agent reads only resource configurations and definitions via the Azure Management REST API.

---

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Authentication](#authentication)
  - [How OAuth Works in This Agent](#how-oauth-works-in-this-agent)
  - [Setting Up a Service Principal](#setting-up-a-service-principal)
  - [Verifying the Token](#verifying-the-token)
- [Running the Agent](#running-the-agent)
- [Resource Inventory](#resource-inventory)
  - [Factory Info](#factory-info)
  - [Pipelines](#pipelines)
  - [Activities](#activities)
  - [Data Flows](#data-flows)
  - [Flowlets](#flowlets)
  - [Datasets](#datasets)
  - [Linked Services](#linked-services)
  - [Integration Runtimes](#integration-runtimes)
  - [Triggers](#triggers)
  - [Parameters](#parameters)
  - [Variables](#variables)
- [Complexity Classification](#complexity-classification)
  - [Classification Standards](#classification-standards)
  - [Per-Resource Classification Rules](#per-resource-classification-rules)
  - [Classification in Excel Output](#classification-in-excel-output)
- [Excel Output](#excel-output)
- [Security Notes](#security-notes)
- [Troubleshooting](#troubleshooting)
- [File Structure](#file-structure)
- [API Reference](#api-reference)

---

## Overview

The agent fetches **11 resource types** from a single ADF factory, classifies 7 of them by complexity, and writes everything into a dedicated colour-coded Excel workbook with an Azure-blue themed Summary sheet and a master Classification Overview tab.

```
adf_inventory_agent.py
        │
        ├── OAuth 2.0 Client Credentials  (Service Principal)
        │         POST /oauth2/v2.0/token
        │         scope: https://management.azure.com/.default
        │
        ├── Azure Management REST API  (api-version: 2018-06-01)
        │
        ├── Classification Engine  (7 resource classifiers)
        │         SIMPLE  ──► green  (#70AD47)
        │         MEDIUM  ──► amber  (#FFB900)
        │         COMPLEX ──► red    (#C00000)
        │
        └── adf_inventory.xlsx
              ├── Summary                    ← count + Simple/Medium/Complex breakdown
              ├── Factory Info
              ├── Pipelines
              ├── Activities
              ├── Data Flows
              ├── Flowlets
              ├── Datasets
              ├── Linked Services
              ├── Integration Runtimes
              ├── Triggers
              ├── Parameters
              ├── Variables
              └── Classification Overview    ← flat list of all classified resources
```

---

## Prerequisites

| Requirement | Version | Notes |
|---|---|---|
| Python | 3.9 or higher | `python --version` to check |
| pip | Any recent | Bundled with Python |
| VS Code | Any | Recommended IDE |
| Git Bash | Any | Needed for curl token tests on Windows |
| Azure Subscription | Any | With an existing ADF factory |
| Azure Service Principal | Any | With Reader role on the ADF resource |

---

## Installation

**1. Place the script** in a project folder:
```
C:\Projects\adf-inventory\
    adf_inventory_agent.py
```

**2. Open in VS Code:**
```
File → Open Folder → select the folder
```

**3. Create and activate a virtual environment** (`Ctrl + Backtick` to open terminal):

```bash
# Create
python -m venv .venv

# Activate — Windows
.venv\Scripts\activate

# Activate — Mac / Linux
source .venv/bin/activate
```

**4. Install dependencies:**
```bash
pip install requests openpyxl
```

> No Azure SDK (`azure-mgmt-datafactory`) is used. The agent uses only `requests` and `openpyxl` to avoid dependency conflicts and Rust-based DLL failures on Windows.

---

## Configuration

Open `adf_inventory_agent.py` and update the 7 constants at the top of the file:

```python
# ── Configuration ─────────────────────────────────────────────────────────────
TENANT_ID         = "YOUR_TENANT_ID"        # Azure AD tenant ID
CLIENT_ID         = "YOUR_CLIENT_ID"        # Service Principal app (client) ID
CLIENT_SECRET     = "YOUR_CLIENT_SECRET"    # Service Principal secret
SUBSCRIPTION_ID   = "YOUR_SUBSCRIPTION_ID" # Azure subscription ID
RESOURCE_GROUP    = "YOUR_RESOURCE_GROUP"  # Resource group containing the factory
FACTORY_NAME      = "YOUR_ADF_NAME"        # ADF factory name
OUTPUT_FILE       = "adf_inventory.xlsx"   # Output filename
# ─────────────────────────────────────────────────────────────────────────────
```

| Variable | Where to find it |
|---|---|
| `TENANT_ID` | Azure Portal → Azure Active Directory → Overview → Tenant ID |
| `CLIENT_ID` | Azure Portal → App Registrations → your SP → Application (client) ID |
| `CLIENT_SECRET` | Azure Portal → App Registrations → your SP → Certificates & secrets |
| `SUBSCRIPTION_ID` | Azure Portal → Subscriptions → your subscription → Subscription ID |
| `RESOURCE_GROUP` | Azure Portal → Resource groups → name of the RG containing the factory |
| `FACTORY_NAME` | Azure Portal → Data factories → factory name |
| `OUTPUT_FILE` | Filename for the Excel output (saved to the same folder as the script) |

---

## Authentication

The agent uses **OAuth 2.0 Client Credentials** (Service Principal / Machine-to-Machine) flow against the Azure AD token endpoint. The scope is `https://management.azure.com/.default`, which grants access to the Azure Management REST API.

### How OAuth Works in This Agent

```
CLIENT_ID + CLIENT_SECRET + TENANT_ID
               │
               ▼
POST https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
     scope = https://management.azure.com/.default
               │
               ▼
         access_token  ──► cached in memory (up to 1 hour)
               │            auto-refreshed 60s before expiry
               ▼
Authorization: Bearer <token>  ──► every Azure REST API call
```

The token is cached in the `_token_cache` dict and refreshed automatically — no manual intervention is needed during a run.

### Setting Up a Service Principal

**Step 1 — Register an Application**

In the Azure Portal:
- Go to **Azure Active Directory** → **App registrations** → **New registration**
- Name it (e.g. `adf-inventory-agent`)
- Leave redirect URI blank → click **Register**
- Copy the **Application (client) ID** — this is your `CLIENT_ID`
- Copy the **Directory (tenant) ID** — this is your `TENANT_ID`

**Step 2 — Generate a Client Secret**

- Inside the App registration → **Certificates & secrets** → **Client secrets** → **New client secret**
- Set a description and expiry → click **Add**
- **Copy the Value immediately** — it is shown only once

**Step 3 — Assign the Role to the Factory**

Navigate to your ADF factory in the Azure Portal:
- **Access control (IAM)** → **Add role assignment**
- Role: **Reader** (minimum required for read-only inventory)
- Members: select your App registration by name
- Click **Review + assign**

> For complete metadata including all pipeline activity details, the **Data Factory Contributor** role may be needed if `Reader` returns incomplete results.

**Step 4 — Note the IDs**

Collect all four IDs before updating the config:

| ID | Location in Azure Portal |
|---|---|
| Tenant ID | Azure AD → Overview |
| Client ID | App Registrations → your app → Overview |
| Client Secret | App Registrations → your app → Certificates & secrets |
| Subscription ID | Subscriptions → your subscription |

### Verifying the Token

Test in Git Bash before running the agent:

```bash
TENANT_ID="your-tenant-id"
CLIENT_ID="your-client-id"
CLIENT_SECRET="your-client-secret"

curl -s -X POST \
  "https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token" \
  -d "grant_type=client_credentials" \
  -d "client_id=${CLIENT_ID}" \
  -d "client_secret=${CLIENT_SECRET}" \
  -d "scope=https://management.azure.com/.default" | python -m json.tool
```

A successful response:
```json
{
    "access_token": "eyJ0eXAiOiJKV1Qi...",
    "token_type": "Bearer",
    "expires_in": 3599
}
```

Then verify ADF access:
```bash
TOKEN="paste-access-token-here"
SUBSCRIPTION_ID="your-subscription-id"
RESOURCE_GROUP="your-resource-group"
FACTORY_NAME="your-factory-name"

curl -s -H "Authorization: Bearer $TOKEN" \
  "https://management.azure.com/subscriptions/${SUBSCRIPTION_ID}/resourceGroups/${RESOURCE_GROUP}/providers/Microsoft.DataFactory/factories/${FACTORY_NAME}?api-version=2018-06-01" \
  | python -m json.tool | head -30
```

A `200` response with factory properties confirms the agent will work.

---

## Running the Agent

**From the VS Code terminal (with `.venv` activated):**

```bash
python adf_inventory_agent.py
```

**Expected output:**

```
==============================================================
  Azure Data Factory Inventory Agent
  Factory : MyADFFactory
  RG      : my-resource-group
  Time    : 2024-05-08 10:30:00
==============================================================
  [AUTH] Token obtained (expires in 3599s)

▶  Fetching Factory Info ...
   ✓  1 record(s) found

▶  Fetching Pipelines ...
   ✓  12 record(s) found

▶  Fetching Activities ...
   ✓  87 record(s) found

▶  Fetching Data Flows ...
   ✓  5 record(s) found

▶  Fetching Flowlets ...
   ✓  2 record(s) found

▶  Fetching Datasets ...
   ✓  24 record(s) found

▶  Fetching Linked Services ...
   ✓  8 record(s) found

▶  Fetching Integration Runtimes ...
   ✓  3 record(s) found

▶  Fetching Triggers ...
   ✓  6 record(s) found

▶  Fetching Parameters ...
   ✓  31 record(s) found

▶  Fetching Variables ...
   ✓  14 record(s) found

==============================================================
  ✅  Inventory saved → adf_inventory.xlsx
==============================================================
```

The Excel file is created in the same folder as the script.

---

## Resource Inventory

### Factory Info
API: `GET /factories/{factoryName}`

Single-row overview of the ADF factory itself. No complexity classification applied.

| Column | Description |
|---|---|
| Factory Name | ADF factory name |
| Resource Group | Azure resource group |
| Subscription ID | Azure subscription identifier |
| Location | Azure region (e.g. eastus) |
| State | Provisioning state (Succeeded, etc.) |
| Create Time | Factory creation timestamp |
| Purview Account | Linked Microsoft Purview resource ID if configured |
| Git Config Enabled | Yes / No — whether Git integration is active |
| Repo Type | FactoryGitHubConfiguration or FactoryVSTSConfiguration |
| Repo Name | Git repository name |
| Branch | Collaboration branch name |
| Global Parameters | Count of factory-level global parameters |
| Public Network Access | Enabled or Disabled |
| Tags | Azure resource tags as key-value dictionary |

---

### Pipelines
API: `GET /factories/{name}/pipelines` — **Classified**

All pipelines in the factory with structural metrics.

| Column | Description |
|---|---|
| Complexity | SIMPLE / MEDIUM / COMPLEX — colour-coded cell |
| Complexity Reason | Human-readable explanation |
| Pipeline Name | Pipeline display name |
| Description | Pipeline description |
| Folder | Folder path (Root if ungrouped) |
| Activity Count | Total number of activities in the pipeline |
| Activity Types | Deduplicated set of activity type names |
| Parameter Count | Number of pipeline-level parameters |
| Variable Count | Number of pipeline-level variables |
| Concurrency | Max concurrent runs if configured |
| Annotations | Pipeline annotation tags |

---

### Activities
API: Flattened from `GET /factories/{name}/pipelines` — **Classified**

Every individual activity from every pipeline, one row per activity. Activities are flattened across all pipelines.

| Column | Description |
|---|---|
| Complexity | SIMPLE / MEDIUM / COMPLEX — colour-coded cell |
| Complexity Reason | Human-readable explanation |
| Pipeline Name | Parent pipeline name |
| Activity Name | Activity display name |
| Activity Type | ADF activity type (Copy, ExecuteDataFlow, Lookup, etc.) |
| Description | Activity description |
| Depends On | Comma-separated upstream activity names |
| Input Datasets | Referenced input dataset names |
| Output Datasets | Referenced output dataset names |
| Linked Service | Directly referenced linked service |
| Notes | Extracted type-specific hints (URL, called pipeline, dataflow name, wait time) |
| Policy Timeout | Activity-level timeout override |
| Policy Retries | Activity-level retry count |

---

### Data Flows
API: `GET /factories/{name}/dataflows` (type = MappingDataFlow) — **Classified**

All Mapping Data Flows with structural metrics. Flowlets (type = Flowlet) are separated into their own sheet.

| Column | Description |
|---|---|
| Complexity | SIMPLE / MEDIUM / COMPLEX — colour-coded cell |
| Complexity Reason | Human-readable explanation |
| DataFlow Name | Data flow display name |
| Type | MappingDataFlow |
| Description | Data flow description |
| Folder | Folder path |
| Source Count | Number of source transformations |
| Sources | Comma-separated source names |
| Sink Count | Number of sink transformations |
| Sinks | Comma-separated sink names |
| Transformation Count | Number of intermediate transformations |
| Transformations | Comma-separated transformation names |
| Script Lines | Line count of the data flow script |
| Annotations | Data flow annotation tags |

---

### Flowlets
API: `GET /factories/{name}/dataflows` (type = Flowlet)

Reusable data flow components (Flowlets). No complexity classification applied — Flowlets are always simple by design as they are extracted sub-flows.

| Column | Description |
|---|---|
| Flowlet Name | Flowlet display name |
| Description | Flowlet description |
| Folder | Folder path |
| Source Count | Number of source parameters |
| Sources | Comma-separated source parameter names |
| Sink Count | Number of sink parameters |
| Sinks | Comma-separated sink parameter names |
| Script Lines | Line count of the flowlet script |
| Annotations | Annotation tags |

---

### Datasets
API: `GET /factories/{name}/datasets` — **Classified**

All datasets with their linked service reference and structural details.

| Column | Description |
|---|---|
| Complexity | SIMPLE / MEDIUM / COMPLEX — colour-coded cell |
| Complexity Reason | Human-readable explanation |
| Dataset Name | Dataset display name |
| Dataset Type | ADF dataset type (DelimitedText, Parquet, AzureSqlTable, etc.) |
| Description | Dataset description |
| Folder | Folder path |
| Linked Service | Referenced linked service name |
| Linked Service Type | Reference type (LinkedServiceReference) |
| Schema Column Count | Number of explicitly defined schema columns |
| Parameter Count | Number of dataset-level parameters |
| Location Type | Storage location type (AzureBlobFSLocation, SftpLocation, etc.) |
| Location Hints | Extracted non-secret location hints (container, fileName, tableName) |
| Annotations | Dataset annotation tags |

> **Security**: Connection strings, account keys, SAS tokens, and all secret-bearing fields are never read or written. Only non-sensitive location metadata is captured.

---

### Linked Services
API: `GET /factories/{name}/linkedservices` — **Classified**

All linked service connection definitions. Secrets are automatically redacted.

| Column | Description |
|---|---|
| Complexity | SIMPLE / MEDIUM / COMPLEX — colour-coded cell |
| Complexity Reason | Human-readable explanation |
| Linked Service Name | Linked service display name |
| Type | ADF linked service type (AzureBlobFS, Sftp, AzureSqlDatabase, etc.) |
| Description | Linked service description |
| Integration Runtime | IR used for this connection |
| Parameter Count | Number of linked service parameters |
| Connection Properties | Up to 6 non-secret configuration fields; secrets shown as `[REDACTED]` |
| Annotations | Annotation tags |

---

### Integration Runtimes
API: `GET /factories/{name}/integrationruntimes` — **Classified**

All compute environments (auto-resolve, self-hosted, Azure SSIS).

| Column | Description |
|---|---|
| Complexity | SIMPLE / MEDIUM / COMPLEX — colour-coded cell |
| Complexity Reason | Human-readable explanation |
| IR Name | Integration runtime name |
| Type | Managed (cloud), SelfHosted, or LinkedIntegrationRuntime |
| Description | IR description |
| Location | Azure region or AutoResolve |
| Node Size | VM size for Azure SSIS IR |
| Node Count | Number of nodes |
| Max Parallel Exec/Node | Parallel execution slots per node |
| Compute Type | DataFlowRuntime, etc. |
| Data Flow Compute | General or MemoryOptimized |
| VNet Enabled | Yes / No — managed VNet integration |
| VNet Subnet | Subnet resource ID if VNet-managed |
| Edition | Standard or Enterprise (SSIS IR) |
| License Type | BasePrice or LicenseIncluded (SSIS IR) |
| Catalog DB | SSISDB server endpoint if SSIS IR |
| Auto Update | Auto-update setting for self-hosted IR |

---

### Triggers
API: `GET /factories/{name}/triggers` — **Classified**

All trigger definitions including schedule, tumbling window, and event-based triggers.

| Column | Description |
|---|---|
| Complexity | SIMPLE / MEDIUM / COMPLEX — colour-coded cell |
| Complexity Reason | Human-readable explanation |
| Trigger Name | Trigger display name |
| Type | ScheduleTrigger, TumblingWindowTrigger, BlobEventsTrigger, CustomEventsTrigger |
| Description | Trigger description |
| Runtime State | Started or Stopped |
| Schedule | Human-readable schedule description |
| Start Time | Schedule start datetime |
| End Time | Schedule end datetime (N/A if open-ended) |
| Timezone | Schedule timezone |
| Pipeline Count | Number of pipelines this trigger fires |
| Pipelines | Comma-separated pipeline names |
| Annotations | Annotation tags |

---

### Parameters
API: Factory global params from `GET /factories/{name}` + pipeline params from `GET /factories/{name}/pipelines`

All parameters at both factory (global) and pipeline scope. No complexity classification applied — parameters are configuration primitives.

| Column | Description |
|---|---|
| Scope | Factory (Global) or Pipeline |
| Pipeline Name | Factory name for global params; pipeline name for pipeline-level params |
| Parameter Name | Parameter identifier |
| Type | String, Int, Bool, Array, Object, Float |
| Default Value | Default value (empty if none) |
| Description | Parameter description if defined |

---

### Variables
API: Flattened from `GET /factories/{name}/pipelines`

All pipeline-level variables. ADF does not have factory-level variables. No complexity classification applied — variables are runtime state holders.

| Column | Description |
|---|---|
| Pipeline Name | Parent pipeline name |
| Variable Name | Variable identifier |
| Type | String, Integer, Boolean, Array |
| Default Value | Initial/default value |

---

## Complexity Classification

### Classification Standards

Every classified resource receives two additional columns prepended as the first two fields in its sheet:

| Column | Value | Cell Colour | Hex |
|---|---|---|---|
| **Complexity** | `SIMPLE` | Green | `#70AD47` |
| **Complexity** | `MEDIUM` | Amber | `#FFB900` |
| **Complexity** | `COMPLEX` | Red (white text) | `#C00000` |
| **Complexity Reason** | Free-text explanation | Inherits row shading | — |

Classification is applied by `apply_classification()` after each resource is fetched, before it is written to Excel. The highest matching rule wins — a single COMPLEX condition overrides any SIMPLE signals in the same row.

Resources with no meaningful complexity distinction (Factory Info, Flowlets, Parameters, Variables) do not receive classification columns.

---

### Per-Resource Classification Rules

#### Pipelines

| Level | Triggering Conditions |
|---|---|
| **SIMPLE** | < 5 activities, < 5 parameters, < 5 variables, no concurrency limit |
| **MEDIUM** | 5–14 activities · OR · ≥ 5 parameters · OR · ≥ 5 variables · OR · concurrency set |
| **COMPLEX** | ≥ 15 activities |

---

#### Activities

| Level | Triggering Conditions |
|---|---|
| **SIMPLE** | Standard Copy / basic data activity, no upstream dependencies |
| **MEDIUM** | Lookup / GetMetadata / Filter / ForEach / IfCondition / Switch / Until / Wait · OR · 1–2 upstream dependencies · OR · custom retries configured · OR · non-default timeout set |
| **COMPLEX** | ExecutePipeline / ExecuteDataFlow / HDInsight / Databricks / Custom / AzureFunction / WebActivity / Script activity type · OR · ≥ 3 upstream dependencies |

---

#### Data Flows

| Level | Triggering Conditions |
|---|---|
| **SIMPLE** | 1 source, 1 sink, ≤ 3 transformations, < 200 script lines |
| **MEDIUM** | 4–9 transformations · OR · 3–4 total sources + sinks combined |
| **COMPLEX** | ≥ 10 transformations · OR · ≥ 200 script lines · OR · ≥ 5 total sources + sinks combined |

---

#### Datasets

| Level | Triggering Conditions |
|---|---|
| **SIMPLE** | Simple flat format (CSV, TSV, etc.), < 10 schema columns, < 2 parameters |
| **MEDIUM** | JSON / Excel / Snowflake / Google Sheets dataset type · OR · 10–29 schema columns · OR · ≥ 2 parameters |
| **COMPLEX** | Delta Lake / ORC / Avro / Parquet / Binary / XML dataset type · OR · ≥ 30 schema columns |

---

#### Linked Services

| Level | Triggering Conditions |
|---|---|
| **SIMPLE** | Simple cloud storage (Blob, ADLS Gen1, S3, FTP) |
| **MEDIUM** | Azure SQL / SQL Server / MySQL / PostgreSQL / REST / Web / OData / Azure Data Lake / Azure Databricks / Azure File Share |
| **COMPLEX** | Snowflake / Databricks (non-Azure) / HDInsight / Google BigQuery / Amazon Redshift / SAP / Oracle / Teradata / CosmosDB / Dynamics / Salesforce · OR · linked service using a **Self-Hosted IR** |

---

#### Integration Runtimes

| Level | Triggering Conditions |
|---|---|
| **SIMPLE** | Auto-resolve managed IR, single node, no VNet |
| **MEDIUM** | Managed IR with VNet integration · OR · > 1 node configured |
| **COMPLEX** | Self-Hosted IR · OR · Azure SSIS IR with SSISDB catalog configured |

---

#### Triggers

| Level | Triggering Conditions |
|---|---|
| **SIMPLE** | Schedule trigger, started, fires 1 pipeline |
| **MEDIUM** | Tumbling window trigger · OR · fires 2 pipelines · OR · trigger is in Stopped state |
| **COMPLEX** | Blob Events trigger (BlobEventsTrigger) · OR · Custom Events trigger · OR · fires ≥ 3 pipelines |

---

### Classification in Excel Output

**Per-resource sheets**

Classified sheets have `Complexity` and `Complexity Reason` as the first two columns. The `Complexity` cell is filled with the level colour; COMPLEX uses white text for contrast:

```
┌───────────┬─────────────────────────────────────────┬──────────────────────┬ ...
│ Complexity│ Complexity Reason                       │ Pipeline Name        │
├───────────┼─────────────────────────────────────────┼──────────────────────┤
│  SIMPLE   │ simple linear pipeline                  │ pl_copy_reference    │
│  MEDIUM   │ 7 activities (5-14); concurrency set    │ pl_ingest_sales      │
│  COMPLEX  │ 18 activities (>=15)                    │ pl_load_warehouse    │
└───────────┴─────────────────────────────────────────┴──────────────────────┘
   🟢 green                                              🔴 red
```

**Summary sheet**

Includes three colour-coded count columns — Simple, Medium, Complex — with a SUM total row:

```
┌──────────────────────┬───────┬────────┬────────┬─────────┬────────────────────┐
│ Resource Type        │ Count │ Simple │ Medium │ Complex │ Sheet              │
├──────────────────────┼───────┼────────┼────────┼─────────┼────────────────────┤
│ Pipelines            │    12 │      6 │      4 │       2 │ Pipelines          │
│ Activities           │    87 │     41 │     33 │      13 │ Activities         │
│ Data Flows           │     5 │      2 │      2 │       1 │ Data Flows         │
│ ...                  │   ... │    ... │    ... │     ... │ ...                │
├──────────────────────┼───────┼────────┼────────┼─────────┼────────────────────┤
│ TOTAL                │   198 │     98 │     72 │      28 │                    │
└──────────────────────┴───────┴────────┴────────┴─────────┴────────────────────┘
                                 🟢 green  🟡 amber  🔴 red
```

**Classification Overview sheet**

A flat master sheet at the end of the workbook aggregating every classified resource across all 7 types. Four columns: `Resource Type`, `Complexity`, `Complexity Reason`, `Name / ID`. Independently filterable and sortable — useful for producing a factory-wide governance or migration complexity report.

---

## Excel Output

The output file `adf_inventory.xlsx` contains:

**Summary sheet (first tab)** — Azure-blue themed title block with factory name and generation timestamp. Per-resource-type table with Count, Simple, Medium, Complex columns (colour-coded headers), and a TOTAL SUM row.

**Per-resource sheets** — one sheet per resource type:
- `Complexity` and `Complexity Reason` as the first two columns (classified resources only)
- Colour-coded `Complexity` cell — green / amber / red
- Frozen header row
- Azure-blue column headers with white bold text
- Alternating row shading (light blue `#EBF3FB`)
- Auto-fitted column widths (max 60 characters)
- Excel Table with sort and filter dropdowns on every column
- Colour-coded tab per sheet

**Classification Overview sheet (last tab)** — flat list of all classified resources across all 7 classifiable types.

**Theme colours:**

| Element | Hex | Used For |
|---|---|---|
| Header background | `#0078D4` | Column headers, summary title (Azure blue) |
| Header text | `#FFFFFF` | Column header font |
| Accent | `#005A9E` | Summary sub-header, border accents |
| Alternate rows | `#EBF3FB` | Even data rows |
| SIMPLE | `#70AD47` | Complexity cell fill — green |
| MEDIUM | `#FFB900` | Complexity cell fill — amber |
| COMPLEX | `#C00000` | Complexity cell fill — red (white text) |

**Sheet tab colours:**

| Sheet | Tab Colour |
|---|---|
| Summary / Factory Info / Classification Overview | Azure Blue `#0078D4` |
| Pipelines | Dark Blue `#005A9E` |
| Activities | Navy `#003F6E` |
| Data Flows | Red `#E81123` |
| Flowlets | Orange `#FF8C00` |
| Datasets | Green `#107C10` |
| Linked Services | Purple `#B4009E` |
| Integration Runtimes | Dark Azure `#004E8C` |
| Triggers | Deep Orange `#D83B01` |
| Parameters | Olive `#498205` |
| Variables | Teal `#038387` |

---

## Security Notes

The agent is designed to be safe for use in enterprise environments:

- **No data access** — only Azure Management API endpoints are called. No data plane APIs (storage, SQL, SFTP, etc.) are ever contacted.
- **Secret redaction** — the `get_linked_services()` function scans every connection property key against a blocklist of secret field names (`password`, `servicePrincipalKey`, `accountKey`, `sasToken`, `connectionString`, `clientSecret`, `credential`, `encryptedCredential`) and replaces their values with `[REDACTED]`. Raw secrets never appear in the Excel output.
- **Minimum required role** — the Service Principal needs only the **Reader** role on the ADF factory resource. No write permissions are required.
- **SSL warnings suppressed** — `verify=False` is used to handle corporate proxy certificates. In production, set `verify=True` or provide the CA bundle path via `verify="/path/to/ca-bundle.crt"`.
- **Token in memory only** — the OAuth access token is cached in a Python dict and never written to disk.

---

## Troubleshooting

**`Auth failed [400]: AADSTS700016`**
The Client ID does not exist in the specified tenant. Verify `TENANT_ID` and `CLIENT_ID` match the same App Registration.

**`Auth failed [401]: AADSTS7000215`**
The client secret is invalid or expired. Regenerate it in Azure Portal → App Registrations → Certificates & secrets.

**`HTTP 403: AuthorizationFailed`**
The Service Principal exists and can authenticate, but lacks the Reader role on the factory or subscription. Go to the ADF factory → Access control (IAM) → Add role assignment → Reader.

**`HTTP 404: ResourceNotFound`**
One of `SUBSCRIPTION_ID`, `RESOURCE_GROUP`, or `FACTORY_NAME` is incorrect. Verify all three in the Azure Portal.

**Activities sheet has fewer rows than expected**
Some pipelines may return empty activity lists if the SP has Reader but not Data Factory Reader role at the pipeline level. Grant **Data Factory Contributor** for full access.

**Flowlets sheet shows "No Flowlets found"**
This is normal if the factory does not use reusable Flowlet components. The sheet is included for completeness.

**`ModuleNotFoundError: No module named 'requests'`**
Virtual environment is not activated. Run `.venv\Scripts\activate` (Windows) or `source .venv/bin/activate` (Mac/Linux).

**SSL certificate errors in corporate environments**
Replace `verify=False` with `verify="C:/path/to/company-ca.crt"` in all `requests.get` and `requests.post` calls.

**`000` HTTP status / connection refused**
A firewall is blocking outbound HTTPS to `login.microsoftonline.com` or `management.azure.com`. Run the script from a machine with direct internet access or configure proxy settings.

---

## File Structure

```
adf-inventory/
├── adf_inventory_agent.py   ← main agent script
├── adf_inventory.xlsx       ← generated output (after first run)
└── .venv/                   ← Python virtual environment
```

---

## API Reference

All calls use Azure Management REST API version `2018-06-01`.

Base URL pattern:
```
https://management.azure.com
  /subscriptions/{subscriptionId}
  /resourceGroups/{resourceGroupName}
  /providers/Microsoft.DataFactory
  /factories/{factoryName}
  /{resource}
  ?api-version=2018-06-01
```

| Resource | Method | Path |
|---|---|---|
| OAuth Token | POST | `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token` |
| Factory Info | GET | `/factories/{factoryName}` |
| Pipelines | GET | `/factories/{factoryName}/pipelines` |
| Activities | GET | (flattened from Pipelines response) |
| Data Flows | GET | `/factories/{factoryName}/dataflows` |
| Flowlets | GET | `/factories/{factoryName}/dataflows` (filtered by type=Flowlet) |
| Datasets | GET | `/factories/{factoryName}/datasets` |
| Linked Services | GET | `/factories/{factoryName}/linkedservices` |
| Integration Runtimes | GET | `/factories/{factoryName}/integrationruntimes` |
| Triggers | GET | `/factories/{factoryName}/triggers` |
| Parameters (global) | GET | `/factories/{factoryName}` (from globalParameters) |
| Parameters (pipeline) | GET | `/factories/{factoryName}/pipelines` (from properties.parameters) |
| Variables | GET | `/factories/{factoryName}/pipelines` (from properties.variables) |

All list endpoints support pagination via `nextLink` in the response — the agent follows all pages automatically until the full result set is collected.
