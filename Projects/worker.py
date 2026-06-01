# worker.py
# -----------------------------------------------------------------------------
# Active Directory User Creation Backend Worker
#
# This file is the backend worker process launched by the Flask web app in
# app.py. It is normally not started directly by the end user. The Flask app
# saves the uploaded spreadsheet, creates a job JSON file, and then starts this
# worker as a subprocess with that job file as input.
#
# The worker reads the spreadsheet, connects to Active Directory, creates user
# accounts, fills in attributes, resolves supervisors, adds users to default
# groups, runs PowerShell post-processing, and writes output files such as
# summary.xlsx, run_log.txt, created_users.json, post_process.ps1, and
# AssignOffice365E1.ps1.
#
# SECURITY / PUBLIC REPO NOTE:
# This copy keeps the same code structure and workflow, but uses dummy placeholder
# values for company-specific information such as AD domains, email domains,
# group distinguished names, OU paths, company names, temporary passwords, and
# licensing examples. Replace placeholders only in a private/internal deployment
# environment.
#
# Input:
# - A job.json file created by app.py
# - A spreadsheet containing required user columns
#
# Output:
# - summary.xlsx for processed user accounts
# - run_log.txt for job logs
# - ERROR.txt if the job fails
# - created_users.json for newly created user DNs
# - post_process.ps1 for AD post-processing
# - AssignOffice365E1.ps1 for follow-up license assignment
# -----------------------------------------------------------------------------

import os
import io
import sys
import json
import gc
import traceback
import subprocess
import re  # for splitting hyphens/spaces in last names (used only for sAM calc)

import pandas as pd
from ldap3 import Server, Connection, ALL, SUBTREE, BASE
from ldap3.utils.conv import escape_filter_chars
from pyad import aduser, adcontainer, adgroup, pyad
import pythoncom  # COM for ADSI (pyad)

# --- Force UTF-8 console output on Windows (prevents 'charmap' errors) ---
if os.name == "nt":
    os.environ.setdefault("PYTHONUTF8", "1")
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# ---------------------------
# LDAP helpers
# ---------------------------
# Creates and returns an authenticated LDAP connection.
def connect_ldap(ldap_server, full_username, password):
    srv = Server(ldap_server, get_info=ALL)
    return Connection(srv, full_username, password, auto_bind=True)

# Builds a fallback base DN from the LDAP server hostname.
def derive_dn_from_server(ldap_server: str, fallback: str) -> str:
    host = (ldap_server or "").split(":")[0]
    parts = host.split(".")
    if len(parts) >= 2:
        domain = parts[-2:]
        return ",".join(f"DC={p}" for p in domain)
    return fallback

# Finds the AD base DN using configured fallback values or LDAP root data.
def get_base_dn(conn: Connection, ldap_server: str, fallback: str) -> str:
    if fallback and fallback.strip().lower().startswith("dc="):
        return fallback
    try:
        other = getattr(conn.server.info, "other", None)
        if other:
            dnc = other.get("defaultNamingContext", [])
            if dnc:
                return dnc[0]
    except Exception:
        pass
    try:
        conn.search("", "(objectClass=*)", search_scope=BASE, attributes=["defaultNamingContext"])
        if conn.entries:
            val = conn.entries[0].entry_attributes_as_dict.get("defaultNamingContext")
            if val:
                return val[0]
    except Exception:
        pass
    try:
        conn.search("", "(objectClass=*)", search_scope=BASE, attributes=["namingContexts"])
        if conn.entries:
            ncs = conn.entries[0].entry_attributes_as_dict.get("namingContexts", [])
            for nc in ncs:
                if str(nc).lower().startswith("dc="):
                    return str(nc)
    except Exception:
        pass
    try:
        return derive_dn_from_server(ldap_server, fallback)
    except Exception:
        pass
    return fallback or "DC=example,DC=local"

# Finds a supervisor DN from DOMAIN\user, UPN, or sAMAccountName.
def resolve_user_dn_by_account(account, ldap_server, full_username, password, base_fallback, logger):
    """Resolve DN from DOMAIN\\user, user@domain, or sAMAccountName."""
    if not account:
        return None
    try:
        conn = connect_ldap(ldap_server, full_username, password)
    except Exception as e:
        logger(f"⚠ LDAP bind failed for supervisor username lookup '{account}': {e}")
        return None
    try:
        base = get_base_dn(conn, ldap_server, base_fallback)
        raw = account.strip()
        sams, upns = set(), set()

        if "\\" in raw:      # DOMAIN\user
            sams.add(raw.split("\\", 1)[1])
            upns.add(raw)    # harmless extra try
        elif "@" in raw:     # UPN
            upns.add(raw)
            sams.add(raw.split("@", 1)[0])
        else:                # plain sAM
            sams.add(raw)

        ors = []
        for s in sams: ors.append(f"(sAMAccountName={escape_filter_chars(s)})")
        for u in upns: ors.append(f"(userPrincipalName={escape_filter_chars(u)})")
        if not ors:
            return None

        flt = f"(& (objectClass=user) (|{''.join(ors)}))"
        conn.search(search_base=base, search_filter=flt, search_scope=SUBTREE, attributes=["distinguishedName"])
        if conn.entries:
            dn = conn.entries[0].distinguishedName.value
            logger(f"[DEBUG] Supervisor Username '{account}' -> {dn}")
            return dn
        logger(f"⚠ Supervisor Username '{account}' not found.")
        return None
    except Exception as e:
        logger(f"⚠ Error during supervisor username lookup '{account}': {e}")
        return None
    finally:
        try:
            conn.unbind()
        except Exception:
            pass

# Finds a supervisor DN by guessing common sAMAccountName formats from First Last.
def get_supervisor_dn(supervisor, ldap_server, full_username, password, base_fallback, logger):
    """Fallback: infer sAM from 'First Last' -> FirstLast / FLast."""
    if not supervisor:
        return None
    try:
        conn = connect_ldap(ldap_server, full_username, password)
    except Exception as e:
        logger(f"⚠ LDAP bind failed for supervisor lookup '{supervisor}': {e}")
        return None
    try:
        base = get_base_dn(conn, ldap_server, base_fallback)
        parts = supervisor.strip().split()
        if len(parts) < 2:
            logger(f"⚠ Supervisor '{supervisor}' invalid format (need First Last)")
            return None
        sup_fn = parts[0].replace(" ", "").replace("-", "")
        sup_ln = parts[-1].replace(" ", "").replace("-", "")
        sam1 = f"{sup_fn}{sup_ln}"
        sam2 = f"{sup_fn[0]}{sup_ln}"
        logger(f"[DEBUG] Supervisor search sAMAccountName in {{'{sam1}', '{sam2}'}}")

        for candidate in (sam1, sam2):
            flt = f"(&(objectClass=user)(sAMAccountName={escape_filter_chars(candidate)}))"
            conn.search(search_base=base,
                        search_filter=flt, search_scope=SUBTREE, attributes=["distinguishedName"])
            if conn.entries:
                dn = conn.entries[0].distinguishedName.value
                logger(f"[DEBUG] Supervisor '{supervisor}' -> {dn}")
                return dn
        logger(f"⚠ Supervisor '{supervisor}' not found.")
        return None
    except Exception as e:
        logger(f"⚠ Error while searching supervisor '{supervisor}': {e}")
        return None
    finally:
        try:
            conn.unbind()
        except Exception:
            pass

# Reads the uploaded spreadsheet or CSV into a pandas DataFrame.
def read_input_dataframe(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return pd.read_csv(path, dtype=str).fillna("")
    return pd.read_excel(path, dtype=str).fillna("")

# ---------------------------
# Post-processing via PowerShell
# ---------------------------
# Runs generated PowerShell post-processing for newly created AD users.
def run_post_powershell(created_dns, full_username, password, dest_ou_dn, out_dir, logger, ldap_server):
    """
    Post-steps for ONLY newly created users (mirror original standalone script):
      - proxyAddresses primary = "SMTP:" + userPrincipalName
      - add alias "smtp:" + GivenName + "." + Surname + "@example.org"
      - mail + networkid = userPrincipalName
      - company = "ExampleCompany"
      - Move-ADObject to dest_ou_dn
    Always target the same DC used by creation to avoid replication issues.
    """
    if not created_dns:
        logger("[PS] No newly created users to post-process.")
        return

    logger(f"[PS] Post-processing {len(created_dns)} newly created users...")

    users_json_path = os.path.join(out_dir, "created_users.json")
    with open(users_json_path, "w", encoding="utf-8") as uf:
        json.dump([{"DN": dn} for dn in created_dns], uf)

    ps_code = r"""
param(
  [Parameter(Mandatory=$true)] [string]$UsersJsonPath,
  [Parameter(Mandatory=$true)] [string]$CredUser,
  [Parameter(Mandatory=$true)] [string]$CredPass,
  [Parameter(Mandatory=$true)] [string]$DestOU,
  [Parameter(Mandatory=$true)] [string]$Server
)

$ErrorActionPreference = "Stop"
[Environment]::SetEnvironmentVariable('ADPS_LoadDefaultDrive','0','Process')
Import-Module ActiveDirectory -ErrorAction Stop

$sec  = ConvertTo-SecureString $CredPass -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($CredUser, $sec)

$raw   = Get-Content -LiteralPath $UsersJsonPath -Raw
$users = $raw | ConvertFrom-Json
Write-Output "[PS] Loaded $($users.Count) users for post-processing"
Write-Output "[PS] Using AD Server: $Server"

foreach($u in $users){
  try{
    $dn = $u.DN

    # Retry up to ~12 seconds for visibility on the specified DC
    $usr = $null
    for($i=0; $i -lt 6 -and -not $usr; $i++){
      try{
        $usr = Get-ADUser -Server $Server -Identity $dn -Properties userPrincipalName,GivenName,Surname,DistinguishedName -Credential $cred
      } catch {
        $usr = $null
      }
      if(-not $usr){ Start-Sleep -Seconds 2 }
    }
    if(-not $usr){ Write-Output "[PS] ERROR: user not found on $($Server): $dn"; continue }

    # Primary SMTP proxy = UPN (exactly like original)
    $primary = "SMTP:" + $usr.userPrincipalName
    Set-ADUser -Server $Server -Identity $usr -Replace @{ proxyAddresses = @($primary) } -Credential $cred

    # Alias smtp proxy = GivenName.Surname@example.org (exactly like original)
    $alias = "smtp:" + $usr.GivenName + "." + $usr.Surname + "@example.org"
    try{
      Set-ADUser -Server $Server -Identity $usr -Add @{ proxyAddresses = @($alias) } -Credential $cred
    } catch {
      Write-Output "[PS] Note: add proxy failed for ${dn} (may already exist): $($_.Exception.Message)"
    }

    # mail and networkid set to UPN
    Set-ADObject -Server $Server -Identity $usr.DistinguishedName -Replace @{ mail=$($usr.userPrincipalName); networkid=$($usr.userPrincipalName) } -Credential $cred

    # company = 'ExampleCompany'
    Set-ADObject -Server $Server -Identity $usr.DistinguishedName -Replace @{ company="ExampleCompany" } -Credential $cred

    # move to destination OU
    Move-ADObject -Server $Server -Identity $usr.DistinguishedName -TargetPath $DestOU -Credential $cred

    Write-Output "[PS] OK: updated and moved ${dn}"
  }
  catch{
    Write-Output "[PS] ERROR: ${dn} -> $($_.Exception.Message)"
  }
}
Write-Output "[PS] Post-processing complete."
"""
    ps_path = os.path.join(out_dir, "post_process.ps1")
    with open(ps_path, "w", encoding="utf-8") as pf:
        pf.write(ps_code)

    try:
        proc = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-File", ps_path,
             "-UsersJsonPath", users_json_path,
             "-CredUser", full_username,
             "-CredPass", password,
             "-DestOU", dest_ou_dn,
             "-Server", ldap_server],
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=1800
        )
        if proc.stdout:
            for line in proc.stdout.splitlines():
                logger(line)
        if proc.returncode != 0:
            logger(f"[PS] PowerShell exited with code {proc.returncode}: {proc.stderr.strip() or 'See WORKER_STDERR.txt'}")
    except Exception as e:
        logger(f"[PS] Failed to run PowerShell post-steps: {e}")

# ---------------------------
# Assign Office 365 E1 script (same filename; updated SKU + text)
# ---------------------------
# Writes a PowerShell script that assigns Office 365 E1 licenses from summary.xlsx.
def write_assign_e1_script(out_dir):
    ps = r"""<#
.DESCRIPTION
PowerShell script to assign Office 365 E1 licenses to users listed in the 'Email' column of summary.xlsx
using Microsoft Graph. Logs who was licensed, who was skipped (already licensed), and who failed.
#>

param([string]$ExcelFilePath = "summary.xlsx")

if (-not [IO.Path]::IsPathRooted($ExcelFilePath)) {
  $ExcelFilePath = Join-Path $PSScriptRoot $ExcelFilePath
}

try {
  Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction Stop
} catch {
  Write-Warning "ExecPolicy: $($_.Exception.Message)"
}
$log = Join-Path $PSScriptRoot "AssignOffice365E1-Log.txt"
Start-Transcript -Path $log -Append

function PauseExit { Write-Host "`nPress Enter to exit..."; $null = Read-Host; Stop-Transcript; exit }

try {
  Write-Host "=== AssignOffice365E1 (Office 365 E1) ==="
  Write-Host "Start: $(Get-Date)"
  Write-Host "PSVersion: $($PSVersionTable.PSVersion)"
  Write-Host "ScriptRoot: $PSScriptRoot"
  Write-Host "Excel: $ExcelFilePath"

  $mods = 'Microsoft.Graph.Users','Microsoft.Graph.Identity.DirectoryManagement','ImportExcel'
  foreach($m in $mods){
    if(-not (Get-Module -ListAvailable -Name $m)){
      try{ Install-Module $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop; Write-Host "Installed: $m" }
      catch{ Write-Error "Install $m failed: $($_.Exception.Message)"; PauseExit }
    }
    try{ Import-Module $m -ErrorAction Stop; Write-Host "Imported: $m" }
    catch{ Write-Error "Import $m failed: $($_.Exception.Message)"; PauseExit }
  }

  if(-not (Test-Path $ExcelFilePath)){ Write-Error "Excel not found: $ExcelFilePath"; PauseExit }
  try{ $users = Import-Excel -Path $ExcelFilePath -WorksheetName "Sheet1" | ? { $_.Email } }
  catch{ Write-Error "Read Excel failed: $($_.Exception.Message)"; PauseExit }
  if(-not $users){ Write-Error "No emails found in Excel"; PauseExit }
  Write-Host "Users to process: $($users.Count)"

  Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.ReadWrite.All" -ErrorAction Stop -NoWelcome

  # Office 365 E1 (STANDARDPACK) SKU GUID (example placeholder)
  $targetSkuId = [Guid]'18181a46-0d4e-45cd-891e-60aabd171b4e'
  $sku = Get-MgSubscribedSku | Where-Object { $_.SkuId -eq $targetSkuId }
  if(-not $sku){ Write-Error "Office 365 E1 not found (SkuId 18181a46-0d4e-45cd-891e-60aabd171b4e)"; PauseExit }
  $available = $sku.PrepaidUnits.Enabled - $sku.ConsumedUnits
  Write-Host "Available Office 365 E1 licenses: $available"

  $success=@(); $already=@(); $failed=@()
  $defaultLoc="US"

  foreach($u in $users){
    $upn = ($u.Email.Trim())
    Write-Host "`nProcessing: $upn"

    $mg = Get-MgUser -UserId $upn -ErrorAction SilentlyContinue -Property "id,displayName,usageLocation"
    if(-not $mg){ Write-Warning "  Not found -> skipped"; $failed += $upn; continue }

    if(-not $mg.UsageLocation){
      Write-Host "  Setting usageLocation=$defaultLoc"
      try{ Update-MgUser -UserId $upn -UsageLocation $defaultLoc -ErrorAction Stop }
      catch{ Write-Warning "  Set usageLocation failed -> skipped"; $failed += $upn; continue }
    }

    $have = Get-MgUserLicenseDetail -UserId $upn -ErrorAction Stop | Where-Object { $_.SkuId -eq $sku.SkuId }
    if($have){ Write-Host "  Already has Office 365 E1 -> skipped"; $already += $upn; continue }

    $body = @{ addLicenses = @(@{ skuId=$sku.SkuId; disabledPlans=@() }); removeLicenses=@() } | ConvertTo-Json -Depth 10
    try{
      Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$upn/assignLicense" -Body $body -ContentType "application/json" -ErrorAction Stop
      Write-Host "  Licensed (Office 365 E1)"
      $success += $upn
    } catch {
      Write-Error "  License failed: $($_.Exception.Message)"
      $failed += $upn
    }
  }

  Write-Host "`n=== Summary ==="
  Write-Host ("  Success: {0}" -f $success.Count)
  Write-Host ("  Already licensed (skipped): {0}" -f $already.Count)
  Write-Host ("  Failed: {0}" -f $failed.Count)

  if($already.Count){
    Write-Host "  -- Already licensed:"
    $already | ForEach-Object { Write-Host ("    {0}" -f $_) }
  }
  if($failed.Count){
    Write-Host "  -- Failed:"
    $failed | ForEach-Object { Write-Host ("    {0}" -f $_) }
  }

} catch {
  Write-Error "Unhandled: $($_.Exception.Message)"
} finally {
  Disconnect-MgGraph -ErrorAction SilentlyContinue
  PauseExit
}
"""
    # Keep your existing filename for compatibility with app.py bundling
    with open(os.path.join(out_dir, "AssignOffice365E1.ps1"), "w", encoding="utf-8") as fh:
        fh.write(ps)

# ---------------------------
# Core processing
# ---------------------------
# Processes spreadsheet rows, creates AD users, assigns groups, and writes output data.
def process_users(input_path, target_ou_dn, ldap_server, full_username, password,
                  default_temp_pw, base_fallback, dest_ou_dn, out_dir, logger):
    """Returns: (summary_df, combined_log_text)"""
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    text_log = io.StringIO()
    # Writes a message to the saved log and live console output.
    def add_log(msg):
        text_log.write(msg + "\n")
        logger(msg)

    conn = None
    created_rows = []
    created_dns = []

    try:
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input not found: {input_path}")

        try:
            conn = connect_ldap(ldap_server, full_username, password)
        except Exception as e:
            add_log(f"⚠ LDAP bind failed for main run: {e}")
            raise

        _ = get_base_dn(conn, ldap_server, base_fallback)
        pyad.set_defaults(ldap_server=ldap_server, username=full_username, password=password)
        temp_ou = adcontainer.ADContainer.from_dn(target_ou_dn)

        df = read_input_dataframe(input_path)
        required = ["First Name", "Last Name", "Employee ID", "Job Title", "Department", "Location", "Supervisor"]
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"Missing required columns: {', '.join(missing)}")

        add_log(f"[START] Processing {len(df)} rows")

        for idx, row in df.iterrows():
            raw_fn = str(row.get("First Name", "")).strip()
            raw_ln = str(row.get("Last Name", "")).strip()
            if not raw_fn or not raw_ln:
                add_log(f"[!] Row {idx+2}: Missing First/Last Name -> SKIPPED")
                continue

            # Build account parts:
            # First name: remove spaces/hyphens for account name
            # Last name (for sAM): take only the part before first hyphen OR space
            fn_clean = raw_fn.replace(" ", "").replace("-", "")
            ln_primary = re.split(r"[-\s]", raw_ln, 1)[0].strip()
            ln_clean_for_sam = ln_primary.replace(" ", "").replace("-", "")

            eid        = str(row.get("Employee ID", "")).strip()
            job_title  = str(row.get("Job Title", "")).strip()
            department = str(row.get("Department", "")).strip()
            office     = str(row.get("Location", "")).strip()
            supervisor = str(row.get("Supervisor", "")).strip()
            sup_user   = str(row.get("Supervisor Username", "")).strip()

            display_name = f"{raw_fn} {raw_ln}"
            username     = fn_clean + ln_clean_for_sam
            upn          = f"{username}@example.org"
            temp_pw      = default_temp_pw

            # Check existence (within target OU)
            cn_esc  = escape_filter_chars(display_name)
            sam_esc = escape_filter_chars(username)
            upn_esc = escape_filter_chars(upn)
            search_filter = (
                f"(&(|(objectClass=user)(objectClass=person))"
                f"(|(cn={cn_esc})(sAMAccountName={sam_esc})(userPrincipalName={upn_esc})))"
            )
            conn.search(search_base=target_ou_dn, search_filter=search_filter, search_scope=SUBTREE,
                        attributes=["distinguishedName"])
            if conn.entries:
                add_log(f"[SKIP] {display_name} already exists (found in target OU) -> included in summary")
                created_rows.append({"First Name": raw_fn, "Last Name": raw_ln, "Email": upn, "Employee ID": eid})
                continue

            try:
                supervisor_dn = None
                if sup_user:
                    add_log(f"[DEBUG] Resolving Supervisor Username '{sup_user}'")
                    supervisor_dn = resolve_user_dn_by_account(
                        sup_user, ldap_server, full_username, password, base_fallback, add_log
                    )
                if not supervisor_dn and supervisor:
                    supervisor_dn = get_supervisor_dn(
                        supervisor, ldap_server, full_username, password, base_fallback, add_log
                    )

                attrs = {
                    "givenName": raw_fn,
                    "sn": raw_ln,
                    "employeeID": eid,
                    "sAMAccountName": username,
                    "userPrincipalName": upn,
                    "displayName": display_name,
                    "title": job_title,
                    "description": job_title,
                    "department": department,
                    "physicalDeliveryOfficeName": office
                }
                if supervisor_dn:
                    attrs["manager"] = supervisor_dn

                add_log(f"[CREATE] {display_name} ...")
                new_user = aduser.ADUser.create(display_name, temp_ou, password=temp_pw, optional_attributes=attrs)
                new_user.force_pwd_change_on_login()
                new_user.enable()
                new_user.update_attribute("userAccountControl", 512)
                add_log(f"[OK] Created {display_name} (ID:{eid} Title:{job_title} Dept:{department} Office:{office} Supervisor:{supervisor or sup_user})")

                try:
                    created_dns.append(new_user.dn)
                except Exception:
                    pass

                # Default groups
                for dn in [
                    "CN=Standard.Users,OU=Security Groups,OU=Groups,DC=example,DC=local",
                    "CN=SSPR_Group,CN=Users,DC=example,DC=local",
                    "CN=Security Training Group,OU=Security Groups,OU=Groups,DC=example,DC=local",
                    "CN=Browser Policy Group,OU=Security Groups,OU=Groups,DC=example,DC=local",  # <-- added/updated DN
                ]:
                    try:
                        grp = adgroup.ADGroup.from_dn(dn)
                        grp.add_members([new_user])
                        add_log(f"    [+] Added to group: {grp.get_attribute('cn')[0]}")
                    except Exception as ge:
                        add_log(f"    [!] Group add failed ({dn}): {ge}")

                try:
                    del new_user
                    del grp
                except Exception:
                    pass

                created_rows.append({"First Name": raw_fn, "Last Name": raw_ln, "Email": upn, "Employee ID": eid})

            except Exception as e:
                msg = str(e)
                if ("object already exists" in msg.lower()) or ("0x80071392" in msg) or ("-2147019886" in msg):
                    add_log(f"[SKIP] {display_name} already exists (caught on create) -> included in summary")
                    created_rows.append({"First Name": raw_fn, "Last Name": raw_ln, "Email": upn, "Employee ID": eid})
                else:
                    add_log(f"[ERROR] {display_name}: {e}")
                    add_log(traceback.format_exc(limit=1))

        # Post-steps for newly created only
        try:
            run_post_powershell(
                created_dns=created_dns,
                full_username=full_username,
                password=password,
                dest_ou_dn=dest_ou_dn,
                out_dir=out_dir,
                logger=add_log,
                ldap_server=ldap_server,   # <-- ensure same DC
            )
            add_log("[NOTICE] Microsoft 365 directory sync may take up to ~30 minutes to fully reflect changes.")
        except Exception as e:
            add_log(f"[PS] Post-processing step failed to start: {e}")

        add_log(f"[DONE] Completed. Summary rows: {len(created_rows)}")
        summary_df = pd.DataFrame(created_rows) if created_rows else pd.DataFrame(
            columns=["First Name", "Last Name", "Email", "Employee ID"]
        )
        return summary_df, text_log.getvalue()

    finally:
        try:
            if conn:
                conn.unbind()
        except Exception:
            pass
        gc.collect()
        pythoncom.CoUninitialize()

# ---------------------------
# Entrypoint
# ---------------------------
# Loads the job file, starts processing, writes outputs, and returns the final status.
def main():
    if len(sys.argv) < 2:
        print(json.dumps({"ok": False, "error": "missing job.json"}), flush=True)
        sys.exit(2)
    job_path = sys.argv[1]
    with open(job_path, "r", encoding="utf-8") as jf:
        job = json.load(jf)

    input_path      = job["input_path"]
    target_ou_dn    = job["target_ou_dn"]
    ldap_server     = job["ldap_server"]
    full_username   = job["full_username"]
    password        = job["password"]
    default_temp_pw = job["default_temp_pw"]
    out_dir         = job["out_dir"]
    base_fallback   = job.get("base_dn_fallback") or os.environ.get("BASE_DN", "DC=example,DC=local")
    dest_ou_dn      = job.get("dest_ou_dn") or os.environ.get("DEST_OU_DN", "OU=Active,OU=Users,DC=example,DC=local")

    # Prints worker log messages immediately for live streaming.
    def logger(msg):
        print(msg, flush=True)

    try:
        print("[JOB] Starting user creation job", flush=True)
        summary_df, log_text = process_users(
            input_path=input_path,
            target_ou_dn=target_ou_dn,
            ldap_server=ldap_server,
            full_username=full_username,
            password=password,
            default_temp_pw=default_temp_pw,
            base_fallback=base_fallback,
            dest_ou_dn=dest_ou_dn,
            out_dir=out_dir,
            logger=logger
        )

        os.makedirs(out_dir, exist_ok=True)
        with pd.ExcelWriter(os.path.join(out_dir, "summary.xlsx"), engine="openpyxl") as xw:
            summary_df.to_excel(xw, index=False)
        with open(os.path.join(out_dir, "run_log.txt"), "w", encoding="utf-8") as lf:
            lf.write(log_text)

        # Writes AssignM365BusinessBasic.ps1 (filename kept) but content assigns Office 365 E1
        write_assign_e1_script(out_dir)

        result = {"ok": True, "count": int(len(summary_df))}
        with open(os.path.join(out_dir, "result.json"), "w", encoding="utf-8") as rf:
            json.dump(result, rf)

        print(json.dumps(result), flush=True)
        sys.exit(0)

    except Exception as e:
        os.makedirs(out_dir, exist_ok=True)
        with open(os.path.join(out_dir, "ERROR.txt"), "w", encoding="utf-8") as ef:
            ef.write(f"{type(e).__name__}: {e}\n\n{traceback.format_exc()}")
        result = {"ok": False, "error": str(e)}
        with open(os.path.join(out_dir, "result.json"), "w", encoding="utf-8") as rf:
            json.dump(result, rf)
        print(json.dumps(result), flush=True)
        sys.exit(1)

if __name__ == "__main__":
    main()
