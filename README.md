# psToDo — Calendar Reminder (On-premises edition)

![logo](img/pstodologo.png)

The **on-prem edition** of [psToDo](https://github.com/fardinbarashi/psToDo). Same job — watch expiry dates in a JSON file and alert a team by mail before things lapse, plus a self-contained HTML dashboard — but built for environments that run **their own Exchange server** instead of Microsoft 365.

The **main script and the HTML report** only talk to **your on-prem Exchange server**, over SMTP. No Microsoft Graph, no app registration, no Teams webhook.

The **optional cloud importers** (Entra app registrations and M365 Message Center) are included from the cloud edition for hybrid organisations. They are the only parts that contact Microsoft Graph, the main script never calls them, and they can be left unused on an isolated network.

Based on psToDo 1.3. Runs on **Windows PowerShell 5.1** and **PowerShell 7.x**.

---

## Table of Contents
- [Setup from start to finish](#setup-from-start-to-finish) (step 11: [edit in the browser](#11-optional-edit-the-data-in-the-browser))
- [Maintenance](#maintenance)
- [Troubleshooting](#troubleshooting)
- [What is different from the cloud edition](#what-is-different-from-the-cloud-edition)
- [Components](#components)
- [Requirements](#requirements)
- [SMTP authentication modes](#smtp-authentication-modes)
- [How alerting works](#how-alerting-works)
- [Adding an object to monitor](#adding-an-object-to-monitor)
- [Object field reference](#object-field-reference)
- [State file & completed archive](#state-file--completed-archive)
- [Plugins (importers)](#plugins-importers)
- [HTML dashboard](#html-dashboard)
- [Repository layout](#repository-layout)

---

## Setup from start to finish

The examples use the domain `CONTOSO`, the mail domain `contoso.com`, the Exchange server `ex01.contoso.local` and the install folder `D:\Automation\psToDo`. Replace them with your own values.

### 0. Decide the names

Write these down first; every later step uses them.

| Item | Example | Notes |
|------|---------|-------|
| Server that runs psToDo | `AUTO01` | Domain-joined Windows Server, Windows PowerShell 5.1 is enough |
| Install folder | `D:\Automation\psToDo` | Not under a web root |
| Service account | `CONTOSO\svc-pstodo` | Runs every scheduled task |
| Sender mailbox | `svc-pstodo@contoso.com` | The service account's own mailbox (simplest), or a shared mailbox it has Send-As on |
| SMTP endpoint | `mail.contoso.com`, port 587 | The client submission connector |
| Exchange PowerShell | `http://ex01.contoso.local/PowerShell/` | Only for the Exchange certificate plugin |
| IIS site | `psToDo` | Points at the report folder |
| Viewer group | `CONTOSO\psToDo-Viewers` | AD **security** group allowed to open the dashboard |

### 1. Create the service account and its mailbox

Create a normal AD user (for example `svc-pstodo`) with a strong password. Then give it a mailbox in the Exchange Management Shell:

```powershell
Enable-Mailbox -Identity "CONTOSO\svc-pstodo"
Get-Mailbox svc-pstodo | Format-List PrimarySmtpAddress
```

Note the **PrimarySmtpAddress**. It is the only address the account can send as, and it goes into `mailSender` later. A common mistake is to use the server name (`svc-pstodo@mail.contoso.com`) instead.

To send from a shared mailbox instead, grant Send-As and use that mailbox's address:

```powershell
Add-ADPermission -Identity "pstodo-noreply" -User "CONTOSO\svc-pstodo" -ExtendedRights "Send As"
```

### 2. Copy the files

1. Unpack the zip to the install folder, for example `D:\Automation\psToDo`.
2. Remove the "downloaded from the internet" mark so the execution policy does not block the scripts:

    ```powershell
    Get-ChildItem "D:\Automation\psToDo" -Recurse -Filter *.ps1 | Unblock-File
    ```

3. Give the service account **Modify** on the install folder. The scripts write logs, state, backups, the report and `monitorobjects.json`:

    ```powershell
    icacls "D:\Automation\psToDo" /grant "CONTOSO\svc-pstodo:(OI)(CI)M"
    ```

### 3. Configure the SMTP connection

Edit `Settings\Config\ExchangeOnpremSettings.json`. The recommended setup signs in as the service account on port 587. It needs no change on the Exchange side and can deliver to external recipients:

```json
"SmtpServer": "mail.contoso.com",
"SmtpPort": 587,
"UseTls": true,
"TlsAllowUntrustedCertificate": false,
"AuthMode": "Credential",
"CredentialXmlPath": "Settings\\Config\\smtp-credential.xml",
"ExchangePowerShellUri": "http://ex01.contoso.local/PowerShell/",
"ExchangeAuthentication": "Kerberos"
```

`SmtpServer` must match a name on the Exchange TLS certificate, otherwise STARTTLS fails. For the other modes (anonymous relay, Windows sign-in) see [SMTP authentication modes](#smtp-authentication-modes).

`ExchangePowerShellUri` is only used by the Exchange certificate plugin. Use a server's **internal** name, not the load-balanced mail name, because Kerberos normally fails against a load-balanced name. Find it in the Exchange Management Shell:

```powershell
Get-PowerShellVirtualDirectory | Format-List Server, InternalUrl
```

### 4. Store the service account's password

The password is stored in `Settings\Config\smtp-credential.xml`, encrypted with Windows DPAPI. Only the account that created the file can read it, and only on the machine where it was created. So create it **as the service account, on the psToDo server**:

```powershell
runas /user:CONTOSO\svc-pstodo powershell.exe
# in the new window:
cd "D:\Automation\psToDo"
.\Install-help\New-SmtpCredentialFile.ps1 -UserName CONTOSO\svc-pstodo
```

The script prompts for the password and checks that the file can be read back.

### 5. Test the mail connection

Still in the window running as the service account:

```powershell
.\Install-help\Test-OnpremSmtp.ps1 -To you@contoso.com -From svc-pstodo@contoso.com
```

You should receive a test mail. The output should say `auth: Credential`. If it fails, see [Troubleshooting](#troubleshooting).

### 6. Add what to monitor

Edit `Files\db\monitorobjects.json`. The shipped file holds sample objects from a lab. Replace them with your own (see [Adding an object to monitor](#adding-an-object-to-monitor)). In every object:

- `mailSender` = the address from step 1.
- `mailRecipients` = who gets the alert.
- `notifyMethodbyMail` = `true`.

### 7. Optional: set up the plugins

The plugins fill `monitorobjects.json` automatically. In each plugin you use, set `mailSender` and `mailRecipients` in its defaults file (`Settings\plugin\<name>\config\*-defaults.json`). Objects imported with the placeholder `REPLACE_ME_WITH_SENDER_MAIL_ADDRESS` fail to send.

| Plugin | Extra setup |
|--------|-------------|
| Exchange certificates | The service account needs the Exchange role *View-Only Configuration*: `New-ManagementRoleAssignment -Role "View-Only Configuration" -User "svc-pstodo"` |
| AD account expiry | RSAT on the psToDo server: `Install-WindowsFeature RSAT-AD-PowerShell`. Set the OUs and groups in `adaccount-defaults.json` |
| Entra / Message Center | Internet access, an app registration with a certificate and `MsGraphSettings.json`, as in the [cloud edition](https://github.com/fardinbarashi/psToDo#app-registration--graph-permissions) |

Always run a plugin with `-WhatIf` the first time:

```powershell
.\Settings\plugin\Import-ExchangeCertificates-Onprem-v1.0.ps1 -WhatIf
```

### 8. Dry run

```powershell
.\psToDo-Onprem-v1.0.ps1 -WhatIf
.\psToDo-HTML-Report-Onprem-v1.0.ps1
```

`-WhatIf` lists what would be sent without sending or recording anything. Run the scripts as files (F5 or `.\script.ps1`), not as a selection (F8).

### 9. Create the scheduled tasks

Create the tasks on the psToDo server, all running as the service account with **Run whether user is logged on or not**. Saving the task normally grants the account *Log on as a batch job*, unless a GPO controls that right.

| Task | Suggested trigger | Script |
|------|-------------------|--------|
| psToDo – Import | Weekly, before the alert run | The plugins you use, one action each |
| psToDo – Alert | Daily | `psToDo-Onprem-v1.0.ps1` |
| psToDo – Report | Daily, after the alert run | `psToDo-HTML-Report-Onprem-v1.0.ps1` |

Action for each task (quote the path if it contains spaces):

- Program: `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`
- Arguments: `-NoLogo -NoProfile -File "D:\Automation\psToDo\psToDo-Onprem-v1.0.ps1"`
- Start in: `D:\Automation\psToDo`

Run the alert task once by hand and check the newest file in `Logs\`. It should end with `Mail: n of n sent`.

### 10. Publish the dashboard in IIS

The dashboard is a static page, `Files\report\index.html`. Only members of one AD group may see it.

1. Create the viewer group and add the people who should see the page:

    ```powershell
    New-ADGroup -Name "psToDo-Viewers" -GroupScope Global -GroupCategory Security
    Add-ADGroupMember -Identity "psToDo-Viewers" -Members alice, bob
    ```

2. On the IIS server, as Administrator, run the helper. It installs Windows Authentication and URL Authorization, writes `web.config` next to `index.html` with your group, sets Windows Authentication on and Anonymous off for the site, and gives the group and the app pool Read on the folder:

    ```powershell
    .\Install-help\Setup-IisReportSecurity.ps1 -SiteName psToDo -Group "CONTOSO\psToDo-Viewers" -CreateSite -Port 8080 -WhatIf
    .\Install-help\Setup-IisReportSecurity.ps1 -SiteName psToDo -Group "CONTOSO\psToDo-Viewers" -CreateSite -Port 8080
    ```

    Leave out `-CreateSite -Port` if the site already exists. If the site points at another folder, add `-PhysicalPath` and run the report with `-websitePath "<that folder>\index.html"`.

3. Check the NTFS rights the helper set on the report folder (`Files\report`). With Windows Authentication, IIS reads the page **as the signed-in user** and reads `web.config` as the site's app pool, so both need Read. With `-CreateSite` the site gets its own app pool, `IIS AppPool\psToDo`. If you skipped this with `-SkipNtfs`, set it by hand:

    ```powershell
    icacls "D:\Automation\psToDo\Files\report" /grant "CONTOSO\psToDo-Viewers:(OI)(CI)R"
    icacls "D:\Automation\psToDo\Files\report" /grant "IIS AppPool\psToDo:(OI)(CI)R"
    ```

4. Add an **https** binding with a certificate. Sign-in is protected over http too, but the page content is not.

5. Test: a group member sees the page, anyone else gets HTTP 401. A user who was just added to the group must sign out and in again first.

### 11. Optional: edit the data in the browser

`edit.html` shows the same objects as a table people can change directly: name, expire date, status, triggers, mail on/off, and in a details panel the server, description, sender, recipients, subject and body. Objects can be added, duplicated and deleted. Members of an **editors** group can save; everyone else who can open the site sees the same page read-only.

Saving goes through `psToDo-api.ashx`, a small ASP.NET handler in the same IIS folder. It writes `Files\db\monitorobjects.json` directly.

1. Create the editors group. Keep it small: editors decide who gets alert mail and what it says.

    ```powershell
    New-ADGroup -Name "psToDo-Editors" -GroupScope Global -GroupCategory Security
    Add-ADGroupMember -Identity "psToDo-Editors" -Members alice
    ```

2. On the IIS server, as Administrator, after step 10:

    ```powershell
    .\Install-help\Setup-IisReportEditor.ps1 -SiteName psToDo -EditorsGroup "CONTOSO\psToDo-Editors" -WhatIf
    .\Install-help\Setup-IisReportEditor.ps1 -SiteName psToDo -EditorsGroup "CONTOSO\psToDo-Editors"
    ```

    It installs ASP.NET 4.x for IIS, sets the app pool to .NET CLR v4.0, lets the editors group and the `.ashx` extension through `web.config`, writes the paths and the group to `appSettings`, and gives the app pool **Modify** on `Files\db`, `Files\backup\psToDo` and `Logs` (the handler runs as the app pool, not as the signed-in user).

3. Generate the page, and add the same line to the report task from step 9:

    ```powershell
    .\psToDo-HTML-Report-Editable-Onprem-v1.0.ps1
    ```

4. Open `https://<server>/edit.html` as an editor: the page says **Editing**. A viewer sees **Read-only**.

What protects the data:

- Every save is checked against the version the page loaded. If the file changed in the meantime (another editor, a plugin run), the save is refused and the page offers to download your version. Nothing is overwritten.
- A backup of `monitorobjects.json` is taken before every save (`Files\backup\psToDo\monitorobjects-<time>-web-<user>.json`), and every save is logged with the user name in `Logs\psToDo-web-edits.log`.
- The page validates before saving (name required, valid dates, whole-number triggers, e-mail addresses), and the handler checks again.
- Fields the page does not show, such as the plugin keys, are kept unchanged.
- Saving requires a custom header and a JSON body, which a browser cannot send from another site, so a malicious page cannot save on an editor's behalf.

Alerts and `index.html` pick up a change on their next scheduled run.

### 12. Final checklist

- [ ] Test mail from step 5 arrived.
- [ ] `-WhatIf` run lists the expected objects.
- [ ] The alert task ran on schedule and the log says `auth: Credential` and `Mail: n of n sent`.
- [ ] The dashboard opens for a group member and returns 401 for others.
- [ ] If you use the editor: an editor can save in `edit.html`, a viewer sees it read-only.
- [ ] `monitorobjects.json` and `*-credential.xml` are not under any web root and not in git.

---

## Maintenance

**Service account password changed:** update the password on every scheduled task, then recreate `smtp-credential.xml` as in step 4. Until then every mail fails with `5.7.57`.

**Moving to another server:** the credential file does not move with the files. Recreate it on the new server as the service account.

**Upgrading psToDo:** keep `Files\db\monitorobjects.json`, `Files\state\`, `Settings\Config\ExchangeOnpremSettings.json`, `Settings\Config\smtp-credential.xml`, your plugin defaults and `Files\report\web.config`. Replace everything else.

**Disk space:** nothing is cleaned up automatically. `Logs\` and `Files\backup\` grow by a few files per run.

**Send an alert again:** each trigger window sends once. `.\psToDo-Onprem-v1.0.ps1 -ResetState` forgets what was sent.

---

## Troubleshooting

Start with the newest transcript in `Logs\`.

| Error in the log | Cause | Fix |
|------------------|-------|-----|
| `5.7.57 Client was not authenticated to send anonymous mail` | No sign-in reached Exchange. Often `AuthMode` is not what you think, or `CurrentUser` fails because Kerberos does not work against a load-balanced name | Check that Section 3 says `auth: Credential`. Use `Credential` on port 587 |
| `5.7.60 Client does not have permissions to send as this sender` | `mailSender` is not the account's own address | Set `mailSender` to the PrimarySmtpAddress from step 1, or grant Send-As |
| `5.7.54 Unable to relay` | Anonymous relay to an external recipient | Use `Credential`, or allow external relay on the connector |
| `Could not read the SMTP credential file` | The file was made by another account or on another machine | Recreate it as the service account on this server (step 4) |
| `Cannot reach the on-prem Exchange SMTP endpoint` | DNS, firewall or wrong port | `Test-NetConnection mail.contoso.com -Port 587` |
| `The remote certificate is invalid` | `SmtpServer` is not on the Exchange certificate, or the CA is not trusted | Use the name on the certificate |
| `expired 400 days ago, not a reminder day` | Not an error: expired objects remind every 7 days | — |
| HTTP 401 for a group member | Not signed out and in since joining, the group is a distribution group, or NTFS read is missing | Sign out and in; use a security group; check step 10.3 |
| HTTP 500.19 | Authentication settings put in `web.config`, or URL Authorization not installed | Run `Setup-IisReportSecurity.ps1` again |
| `edit.html`: "Live data could not be loaded ... showing the snapshot" | The page cannot reach `psToDo-api.ashx`: ASP.NET not installed, `.ashx` blocked by request filtering, or the handler is missing | Run `Setup-IisReportEditor.ps1`, then the editable report again |
| `edit.html`: "Read-only: ... is not a member of ..." | The user is not in the editors group, or has not signed in again since being added | Add to the group; sign out and in |
| `edit.html`: "changed by someone else" | Another editor or a plugin wrote the file after the page loaded it | Download your version if needed, then Reload and redo the change |
| `edit.html`: "Not saved: UnauthorizedAccessException ..." | The app pool cannot write `Files\db`, the backup folder or `Logs` | Run `Setup-IisReportEditor.ps1` again (NTFS step) |

---

## What is different from the cloud edition

| Area | Cloud edition (psToDo 1.3) | On-prem edition |
|------|----------------------------|-----------------|
| Mail transport | Microsoft Graph `sendMail`, app registration + certificate | **SMTP to on-prem Exchange** (STARTTLS), anonymous relay connector, Windows auth or stored credential |
| Settings file | `Settings\Config\MsGraphSettings.json` | `Settings\Config\ExchangeOnpremSettings.json` |
| Teams | Workflows webhook | **Not available** (cloud service). `notifyMethodbyTeams` is read but ignored. Mail is the only channel. |
| Modules | `Microsoft.Graph.*`, auto-installed from PSGallery | **None** for the main script and the report. Plugins need the Exchange remote PowerShell endpoint / the RSAT `ActiveDirectory` module — never downloaded automatically. |
| Importer plugins | Entra app registrations, M365 Message Center | The same two cloud importers (optional, need Graph) **plus** on-prem importers for **Exchange certificates** and **AD account / service-account password expiry** |
| Dashboard | Monitored / Completed / No channel / No credentials / Not monitored | Same tabs. Teams shown as "n/a on-prem". Understands the new `source` values. |
| Install helpers | Shared mailbox + Application Access Policy, self-signed cert | DPAPI credential file, SMTP test, IIS access setup, relay Receive connector, and the self-signed cert helper for the cloud importers |

Everything else — trigger windows, state file, backups, Kanban status, completed archive, JSON layout — is unchanged, so an existing `monitorobjects.json` from the cloud edition works as-is (Entra/Message Center rows keep working as ordinary manual rows).

---

## Components

| Script | Job |
|--------|-----|
| `psToDo-Onprem-v1.0.ps1` | Reads `Files\db\monitorobjects.json`, decides what is due, and sends the alerts **through your Exchange server**. |
| `psToDo-HTML-Report-Onprem-v1.0.ps1` | Reads `Files\db\monitorobjects.json` and writes the HTML status page. |
| `psToDo-HTML-Report-Editable-Onprem-v1.0.ps1` | Writes `edit.html`, the page where editors change the data in the browser (built from `Settings\Web\edit-template.html`), and keeps `psToDo-api.ashx` next to it up to date. |
| `Settings\plugin\Import-ExchangeCertificates-Onprem-v1.0.ps1` | Imports every certificate on every Exchange server. |
| `Settings\plugin\Import-ADAccountExpiry-Onprem-v1.0.ps1` | Imports AD account expiry and service-account password expiry. |
| `Settings\plugin\Import-EntraAppRegistrations-v1.3.ps1` (+ `-Delegated`) | Cloud importer: Entra app registration secrets and certificates. Needs Graph. |
| `Settings\plugin\Import-M365MessageCenter-v1.1.3.ps1` (+ `-Delegated-v1.3`) | Cloud importer: Message Center advisories with a deadline. Needs Graph. |
| `Install-help\*` | One-time setup and test helpers (see below). |

Function files live in `Settings\Functions\psToDo`:

| File | Role |
|------|------|
| `Get-OnpremExchangeSettings.ps1` | Parses and validates `ExchangeOnpremSettings.json`. |
| `Connect-OnpremExchangeSmtp.ps1` | Checks the SMTP port answers, builds the `SmtpClient` (TLS + auth mode). Also `Disconnect-OnpremExchangeSmtp`. |
| `Send-AlertMail.ps1` | Builds the HTML mail and submits it over SMTP. |
| `Send-AlertNotification.ps1` | Mail-only dispatcher. |
| `Test-RequiredModule.ps1` | Checks a module is present; never installs anything. |
| `Format-AlertMailBody.ps1`, `Get-AlertPresentation.ps1`, `Get-SentState.ps1`, `Save-SentState.ps1` | Unchanged from the cloud edition. |

The Graph helpers used by the cloud importers (`Connect-CalenderReminderGraph.ps1`, `Initialize-RequiredModules.ps1`) live in a separate folder, `Settings\Functions\cloud-plugins`, so the main script never loads anything Graph-related.

---

## Requirements

| Requirement | Detail |
|-------------|--------|
| PowerShell | Windows PowerShell **5.1** or PowerShell **7.x** |
| OS | Windows Server with Task Scheduler, domain-joined |
| Exchange | On-prem Exchange 2016, 2019 or Subscription Edition |
| Network | TCP to the Exchange SMTP port (587 or 25). Certificate plugin: HTTP to `/PowerShell/` on one Exchange server. AD plugin: LDAP to a domain controller |
| Modules | Main script and report: **none**. AD plugin: `ActiveDirectory` (RSAT). Exchange plugin: none (remote session). Cloud importers: `Microsoft.Graph.*`, installed from PSGallery on first run |
| IIS | Only for the dashboard: Windows Authentication and URL Authorization role services |

---

## SMTP authentication modes

Set in `Settings\Config\ExchangeOnpremSettings.json`. Pick **one**.

| `AuthMode` | Port | When to use | Exchange side |
|------------|------|-------------|---------------|
| `Credential` (recommended) | 587 | Always works when the account has a mailbox; external recipients allowed | Default *Client Frontend* connector. Password stored with `Install-help\New-SmtpCredentialFile.ps1`, run as the task account |
| `Anonymous` | 25 | The psToDo host has a fixed IP and only internal recipients | A relay connector scoped to that IP: `Install-help\Setup-ExchangeRelayReceiveConnector.ps1 -Server ex01 -RemoteIPRanges 10.10.5.20` (add `-AllowExternalRelay` for external recipients). Any sender in an accepted domain |
| `CurrentUser` | 587 | The task account has the mailbox and Kerberos/NTLM works against `SmtpServer` | Default *Client Frontend* connector. Often fails with `5.7.57` against a load-balanced name; use `Credential` then |

`UseTls` turns on STARTTLS; keep it on. `TlsAllowUntrustedCertificate` is for labs only. The `Exchange*` keys are only used by the Exchange certificate plugin.

---

## How alerting works

Identical to the cloud edition. Each `dateTrigger` is a number of **days remaining until expiry** and marks a moment an alert goes out. The script sorts them and treats the **smallest as the most urgent**. Every object has its own triggers and expiry date.

| Window | Status | Meaning |
|--------|--------|---------|
| largest trigger | ok / notice | First, gentle heads-up |
| middle trigger | warning | Second alert point |
| smallest trigger | critical | Last call |
| expiry day | critical | Sent on the day itself |
| below 0 | expired | Reminder every 7 days (`$expiredNotifyDays`) |

Each window fires **once**, recorded in `Files\state\sent-state.json`. A failed send is not recorded, so it is retried on the next run. `-WhatIf` sends nothing and records nothing; `-ResetState` wipes the state file.

Mail subject gets a `[NOTICE]` / `[WARNING]` / `[CRITICAL]` / `[EXPIRED]` prefix; critical and expired mails are flagged high priority.

---

## Adding an object to monitor

`Files\db\monitorobjects.json` is an array of objects you add to manually (or via the plugins).

```json
{
  "id": "21",
  "name": "Wildcard cert lab.local",
  "expireDate": "2027-03-01",
  "template": "WebServerV2",
  "servername": "web-prod-01.lab.local",
  "environment": "prod.local",
  "status": "Backlog",
  "description": "Renew via internal CA and rebind in IIS. See KB-1101.",
  "1dateTrigger": 30,
  "2dateTrigger": 15,
  "3dateTrigger": 7,
  "notifyMethodbyMail": true,
  "notifyMethodbyTeams": false,
  "mail": {
    "mailSender": "pstodo-noreply@lab.local",
    "mailSubject": "Wildcard cert lab.local expires soon",
    "mailBody": "The certificate for web-prod-01.lab.local is approaching its expiry date.",
    "mailRecipients": [ "it-ops@lab.local" ]
  },
  "teams": { "teamSubject": "", "teamBody": "", "teamWebhookUrl": "" }
}
```

`mailSender` must be the sending account's own address or one it has Send-As on (see [step 1](#1-create-the-service-account-and-its-mailbox)). The `teams` block can stay empty; it is kept only so the file is interchangeable with the cloud edition.

---

## Object field reference

| Field | Required | Notes |
|-------|----------|-------|
| `id` | yes | Unique numeric string. Plugins allocate the next free number. |
| `name` | yes | Shown everywhere. |
| `expireDate` | yes | `yyyy-MM-dd`. Invalid or missing → listed under **Not monitored**. |
| `template`, `servername`, `environment` | no | Free text, shown in mail and dashboard. |
| `status` | no | `Backlog` / `Active` / `Completed` (Kanban). `Completed` rows never alert and are archived. |
| `description` | no | "Action required" text in the mail. |
| `1dateTrigger`..`3dateTrigger` | ≥1 | Days before expiry. Order does not matter. |
| `notifyMethodbyMail` | yes | Real boolean. `false` → the object never alerts (shown under **No channel**). |
| `notifyMethodbyTeams` | no | Ignored in the on-prem edition. |
| `mail.mailSender`, `mail.mailRecipients` | yes for mail | Sender address and array of recipients. |
| `mail.mailSubject`, `mail.mailBody` | no | Subject gets the urgency prefix; body is placed above the fact table. |
| `source`, `exchangeCertKey`, `adAccountKey`, … | plugin-written | Identity fields for idempotent re-imports. Do not edit. |

---

## State file & completed archive

Unchanged. **`Files\state\sent-state.json`** records which windows already alerted (`id_expireDate_trigger`). Renewing an object changes `expireDate`, which retires the old keys and re-arms every window.

**`Files\config\MonitorObjectComplete.json`** is written on every report run with a copy of each `Completed` object and a preserved `copiedDate`. Matching uses `exchangeCertKey`, then `adAccountKey`, then `entraKeyId`, `messageId`, `appId`, otherwise `name`+`servername`+`expireDate`.

Backups of `monitorobjects.json`, `sent-state.json` and `ExchangeOnpremSettings.json` are taken into `Files\backup\psToDo\` on every main-script run.

---

## Plugins (importers)

Six importers live in `Settings\plugin`: two on-prem importers and the cloud edition's two importers, each with a delegated variant. All take a timestamped backup of `monitorobjects.json` first, never touch manual objects, do not create duplicates on re-runs, and support `-UseMockData`, `-NoDateRefresh` and `-WhatIf`. Each has its own defaults file (triggers, sender, recipients) and version file under `Settings\plugin\<name>\config\`.

| Script | Source | Talks to | Identity key |
|--------|--------|----------|--------------|
| `Import-ExchangeCertificates-Onprem-v1.0.ps1` | `Get-ExchangeCertificate` on every Exchange server (non-Edge) | Exchange remote PowerShell (`ExchangePowerShellUri`) | `exchangeCertKey` = `<server FQDN>\|<thumbprint>` |
| `Import-ADAccountExpiry-Onprem-v1.0.ps1` | `accountExpires` and `msDS-UserPasswordExpiryTimeComputed` | A domain controller | `adAccountKey` = `<objectGUID>\|account` or `\|password` |
| `Import-EntraAppRegistrations-v1.3.ps1` (+ `-Delegated-v1.3`) | Entra app registration secrets and certificates | Microsoft Graph | `entraKeyId` |
| `Import-M365MessageCenter-v1.1.3.ps1` (+ `-Delegated-v1.3`) | Message Center advisories with an action-required date | Microsoft Graph | `messageId` |

Always run with `-WhatIf` first:

```powershell
.\Settings\plugin\Import-ExchangeCertificates-Onprem-v1.0.ps1 -UseMockData -WhatIf   # no Exchange needed
.\Settings\plugin\Import-ExchangeCertificates-Onprem-v1.0.ps1 -WhatIf                # live
```

### Cloud importers — specifics
Copied from psToDo 1.3 and behave the same way, with two differences: they load their Graph helpers from `Settings\Functions\cloud-plugins`, and the Entra importers have a fix for a crash on re-run (`Cannot convert null to type "System.DateTime"`) when an app registration without credentials was already imported. Configuration: `Settings\Config\MsGraphSettings.json` and `Settings\plugin\entra\config` / `messagecenter\config`, exactly as in the cloud edition. Apps without secrets or certificates appear in the dashboard's **No credentials** tab. See the [cloud README](https://github.com/fardinbarashi/psToDo#plugins-importers) for permissions and the delegated sign-in options.

### Exchange certificate importer — specifics
Imports **only** `Get-ExchangeServer` and `Get-ExchangeCertificate` from the remote session. Self-signed certificates and certificates bound to no service are skipped by default (`includeSelfSigned` / `includeServicesNone` in `exchangecert\config\exchangecert-defaults.json`). Use `-Server mail-01, mail-02` to limit the scope. Tokens: `{{name}}`, `{{expireDate}}`, `{{servername}}`, `{{environment}}`, `{{thumbprint}}`, `{{subject}}`, `{{services}}`. Environment tag: `Exchange on-prem - <domain>`. Extra fields written: `thumbprint`, `subject`, `services`.
**Rights:** Exchange RBAC *View-Only Configuration* for the account running it.

### AD account importer — specifics
Two kinds of rows: **account expiry** (every enabled user under `searchBases` with an `AccountExpirationDate`) and **password expiry** (only users under `passwordExpirySearchBases` or in `passwordExpiryGroups` — meant for service accounts, not the whole user population). `PasswordNeverExpires` accounts and disabled accounts are skipped (`excludeDisabled`). Tokens: `{{name}}`, `{{expireDate}}`, `{{servername}}` (= domain), `{{environment}}`, `{{samAccountName}}`, `{{distinguishedName}}`, `{{expiryType}}`. Use `-DomainController dc-01` to pin a DC.
**Requires:** RSAT `ActiveDirectory` module (`Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0`). Any domain user can read the attributes.

---

## HTML dashboard

```powershell
.\psToDo-HTML-Report-Onprem-v1.0.ps1
.\psToDo-HTML-Report-Onprem-v1.0.ps1 -websitePath 'C:\inetpub\pstodo\index.html'
```

Writes `Files\report\index.html` and a timestamped copy in `Files\backup\psToDo-HTML-Report\`. Tabs: **Monitored**, **Completed**, **No channel** (mail disabled), **No credentials** (Entra apps without secrets or certificates), **Not monitored** (bad date). The page has no external assets, so it works on an isolated network.

Access is limited by `Files\report\web.config` (see [step 10](#10-publish-the-dashboard-in-iis)). It uses IIS **URL Authorization** (`system.webServer/security/authorization`), which also protects static files. A `system.web/authorization` rule would not: it only applies to ASP.NET requests, so a static `index.html` would stay open. Windows Authentication on and Anonymous off cannot be set in `web.config` (that gives HTTP 500.19); the helper sets them on the site.

---

## Repository layout

```
PsToDoOnprem\
├── psToDo-Onprem-v1.0.ps1
├── psToDo-HTML-Report-Onprem-v1.0.ps1
├── psToDo-HTML-Report-Editable-Onprem-v1.0.ps1
├── Files\
│   ├── db\monitorobjects.json            (sample on-prem data)
│   └── report\web.config                 (IIS: one AD group only)
├── Install-help\
│   ├── Setup-ExchangeRelayReceiveConnector.ps1
│   ├── New-SmtpCredentialFile.ps1
│   ├── Test-OnpremSmtp.ps1
│   ├── Setup-IisReportSecurity.ps1        (IIS access for one AD group)
│   ├── Setup-IisReportEditor.ps1          (saving from edit.html for an editors group)
│   └── Create a self-signed cert to app-reg.ps1   (cloud importers only)
└── Settings\
    ├── Config\
    │   ├── ExchangeOnpremSettings.json
    │   ├── MsGraphSettings.json             (cloud importers only)
    │   └── version.json
    ├── Web\
    │   ├── edit-template.html             (the editable page)
    │   └── psToDo-api.ashx                (save handler, copied next to edit.html)
    ├── Functions\
    │   ├── psToDo\                       (main-script functions, no Graph)
    │   ├── psToDo-HtmlReport\
    │   └── cloud-plugins\                (Graph helpers for the cloud importers)
    └── plugin\
        ├── Import-ExchangeCertificates-Onprem-v1.0.ps1
        ├── Import-ADAccountExpiry-Onprem-v1.0.ps1
        ├── Import-EntraAppRegistrations-v1.3.ps1 (+ -Delegated-v1.3)
        ├── Import-M365MessageCenter-v1.1.3.ps1  (+ -Delegated-v1.3)
        ├── exchangecert\config\{exchangecert-defaults.json, version.json}
        ├── adaccount\config\{adaccount-defaults.json, version.json}
        ├── entra\config\{entra-app-defaults.json, version.json}
        └── messagecenter\config\{messagecenter-defaults.json, version.json}
```

Runtime folders (`Logs\`, `Files\state\`, `Files\report\`, `Files\backup\`, `Files\config\`) are created on first run. Keep `monitorobjects.json` and any `*-credential.xml` out of git and out of any web root.

---

Author: Fardin Barashi — [LinkedIn](https://www.linkedin.com/in/fardin-barashi-a56310a2/) · [GitHub](https://github.com/fardinbarashi/psToDo)
