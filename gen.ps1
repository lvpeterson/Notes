#Requires -Version 5.1
<#
.SYNOPSIS
    Builds a pentest notebook in OneNote with headings, tables, and checkboxes.

.PARAMETER NotebookName
    Name for the new notebook. Default: "Pentest - <date>"

.PARAMETER AppNames
    Array of app names to create section groups for.
    Example: -AppNames "CustomerPortal","AdminPanel","API"

.PARAMETER NotebookPath
    Where to save the notebook. Default: Documents\OneNote Notebooks

.EXAMPLE
    .\New-PentestNotebook.ps1 -NotebookName "Acme - Web Assessment" -AppNames "CustomerPortal","API"
#>

[CmdletBinding()]
param(
    [string]$NotebookName = "Pentest - $(Get-Date -Format 'yyyy-MM-dd')",
    [string[]]$AppNames   = @("App 1"),
    [string]$NotebookPath = [System.IO.Path]::Combine(
        [Environment]::GetFolderPath("MyDocuments"),
        "OneNote Notebooks"
    )
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# XML HELPERS
# ---------------------------------------------------------------------------

function Get-Timestamp {
    return [DateTime]::UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
}

# Escape all XML-unsafe characters in a content string
function xe {
    param([string]$Text)
    $t = $Text -replace '&', '&amp;'
    $t = $t   -replace '<', '&lt;'
    $t = $t   -replace '>', '&gt;'
    $t = $t   -replace '"', '&quot;'
    return $t
}

function New-PageXml {
    param(
        [string]$PageId,
        [string]$Title,
        [string]$BodyXml
    )
    $ts = Get-Timestamp
    $safeTitle = xe $Title
    return ("<?xml version=""1.0""?>" +
        "<one:Page xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote""" +
        " ID=""" + $PageId + """ dateTime=""" + $ts + """ lastModifiedTime=""" + $ts + """>" +
        "<one:Title><one:OE><one:T>" + $safeTitle + "</one:T></one:OE></one:Title>" +
        $BodyXml +
        "</one:Page>")
}

function h1 {
    param([string]$Text)
    $t = xe $Text
    return "<one:OE style=""font-size:16.0pt;font-weight:bold;color:#1F3864""><one:T>" + $t + "</one:T></one:OE>"
}

function h2 {
    param([string]$Text)
    $t = xe $Text
    return "<one:OE style=""font-size:12.0pt;font-weight:bold;color:#2E75B6""><one:T>" + $t + "</one:T></one:OE>"
}

function p {
    param([string]$Text = "")
    $t = xe $Text
    return "<one:OE><one:T>" + $t + "</one:T></one:OE>"
}

function mono {
    param([string]$Text = "")
    $t = xe $Text
    return "<one:OE style=""font-family:Courier New;font-size:10.0pt""><one:T>" + $t + "</one:T></one:OE>"
}

function cb {
    param([string]$Text)
    $t = xe $Text
    return "<one:OE><one:Tag index=""0"" completed=""false"" /><one:T>" + $t + "</one:T></one:OE>"
}

function xtr {
    param([string[]]$Cells, [bool]$Header = $false)
    $style = ""
    if ($Header) { $style = " style=""font-weight:bold;background-color:#D9E1F2""" }
    $cellXml = ""
    foreach ($cell in $Cells) {
        $t = xe $cell
        $cellXml += "<one:Cell><one:OEChildren><one:OE" + $style + "><one:T>" + $t + "</one:T></one:OE></one:OEChildren></one:Cell>"
    }
    return "<one:Row>" + $cellXml + "</one:Row>"
}

function xtable {
    param([string[]]$Headers, [object[]]$Rows)
    $cols = ""
    foreach ($h in $Headers) { $cols += "<one:Column />" }
    $headerRow = xtr -Cells $Headers -Header $true
    $dataRows = ""
    foreach ($row in $Rows) {
        $dataRows += xtr -Cells $row
    }
    return "<one:Table bordersVisible=""true""><one:Columns>" + $cols + "</one:Columns>" + $headerRow + $dataRows + "</one:Table>"
}

function xbody {
    param([string[]]$Elements)
    $inner = $Elements -join ""
    return "<one:Outline><one:OEChildren>" + $inner + "</one:OEChildren></one:Outline>"
}

# ---------------------------------------------------------------------------
# PAGE BUILDERS
# ---------------------------------------------------------------------------

function Build-ScopeAndStack {
    $r3 = @( @("","",""), @("","","") )
    $r6 = @( @("","","","","",""), @("","","","","","") )
    return xbody @(
        h1 "Scope and Stack"
        p  "Fill at kickoff. Rarely touched after that."
        p  ""
        h2 "Engagement"
        p  "Client:"
        p  "Engagement Type:  [ ] Web App  [ ] API  [ ] Internal  [ ] Red Team  [ ] Cloud  [ ] WiFi"
        p  "Dates:   Start:                    End:"
        p  "Tester:"
        p  ""
        h2 "In Scope"
        xtable @("App / Target","URL / IP Range","Notes") $r3
        p  ""
        h2 "Out of Scope"
        p  ""
        h2 "Creds / Access Provided"
        xtable @("Account","Role","Notes") $r3
        p  ""
        h2 "Tech Stack"
        p  "Fill as discovered."
        xtable @("App","Framework","Language","Auth","WAF / CDN","Cloud") $r6
        p  ""
        h2 "Contacts"
        xtable @("Name","Role","Contact") $r3
    )
}

function Build-Findings {
    $rows = @(
        @("1","","","","","draft"),
        @("2","","","","","draft"),
        @("3","","","","","draft"),
        @("4","","","","","draft"),
        @("5","","","","","draft")
    )
    return xbody @(
        h1 "Findings"
        p  "Add a row the moment something is confirmed. One place, all apps."
        p  "Severity: Crit / High / Med / Low / Info"
        p  ""
        h2 "Summary"
        xtable @("#","Sev","App","Title","Vuln Class","Status") $rows
        p  ""
        h2 "Finding Detail"
        p  "Duplicate this block for each finding."
        p  ""
        h2 "Finding 1"
        p  "App:"
        p  "Severity:"
        p  "Vuln Class:"
        p  "Endpoint(s):"
        p  "Confirmed:   Yes / No     Date:"
        p  ""
        p  "What it is:"
        p  ""
        p  "How you got there:"
        p  ""
        p  "Impact:"
        p  ""
        p  "Evidence:"
        p  "------------------------------------"
        p  ""
        h2 "Finding 2"
        p  "App:"
        p  "Severity:"
        p  "Vuln Class:"
        p  "Endpoint(s):"
        p  "Confirmed:   Yes / No     Date:"
        p  ""
        p  "What it is:"
        p  ""
        p  "How you got there:"
        p  ""
        p  "Impact:"
        p  ""
        p  "Evidence:"
    )
}

function Build-Surface {
    param([string]$AppName)
    $rows = @(
        @("?","","","","",""),
        @("?","","","","",""),
        @("?","","","","",""),
        @("?","","","","",""),
        @("?","","","","",""),
        @("?","","","","",""),
        @("?","","","","",""),
        @("?","","","","","")
    )
    return xbody @(
        h1 ("Surface - " + $AppName)
        p  "Daily driver. Drop endpoints here as you find them. Tag vuln classes. Update status inline."
        p  "Create a vuln class page only when you have real indicators worth tracking."
        p  ""
        h2 "Status Key"
        p  "?  untested    >  active    v  confirmed (add to Findings)    -  nothing found    x  N/A"
        p  ""
        h2 "Vuln Class Tags"
        p  "sqli   xss   ssti   ssrf   xxe   idor   lfi   auth   csrf   rce   open-redirect   other"
        p  ""
        h2 "Endpoint Map"
        xtable @("Status","Method","Endpoint","Auth","Params of Interest","Vuln Classes") $rows
        p  ""
        h2 "Interesting Headers / Stack Clues"
        p  "Paste response headers, error messages, anything that reveals the stack."
        p  ""
        h2 "JS Files / Hidden Endpoints"
        p  "Endpoints from JS analysis, robots.txt, content discovery."
        p  ""
        h2 "Session / Auth Notes"
        p  "Token format, cookie flags, auth flow quirks - quick notes only."
        p  ""
        h2 "Daily Log"
        p  "One entry per session."
        p  ""
        p  "Date:"
        p  "Tested:"
        p  "Found:"
        p  "Next:"
        p  "------------------------------------"
        p  "Date:"
        p  "Tested:"
        p  "Found:"
        p  "Next:"
        p  "------------------------------------"
        p  "Date:"
        p  "Tested:"
        p  "Found:"
        p  "Next:"
    )
}

function Build-SQLi {
    param([string]$AppName)
    $rows = @( @("?","","","",""), @("?","","","",""), @("?","","","","") )
    return xbody @(
        h1 ("SQLi - " + $AppName)
        p  ("App: " + $AppName)
        p  "Created because:"
        p  ""
        h2 "Candidates"
        xtable @("Status","Method","Endpoint","Param(s)","Type Suspected") $rows
        p  "Types: error-based   blind-boolean   blind-time   union   OOB   second-order   NoSQL"
        p  ""
        h2 "What Tipped You Off"
        p  "Stack trace, timing delta, different response on quote, boolean diff - paste raw."
        p  ""
        h2 "DB Fingerprint"
        p  "Type:      MySQL / MSSQL / PostgreSQL / Oracle / SQLite / MongoDB / Unknown"
        p  "Version:"
        p  "Evidence:"
        p  ""
        h2 "Working Payloads"
        p  "Endpoint:"
        p  "Param:"
        mono "Payload:"
        p  "Result:"
        p  "------------------------------------"
        p  "Endpoint:"
        p  "Param:"
        mono "Payload:"
        p  "Result:"
        p  ""
        h2 "What You Can Reach"
        cb "Schema / DB names"
        cb "Table enumeration"
        cb "Credential tables"
        cb "File read"
        cb "File write / webshell"
        cb "OS command execution"
        cb "Linked servers / lateral movement"
        p  ""
        h2 "Tooling"
        cb "Manual (Burp Repeater)"
        cb "sqlmap   Flags:"
        cb "Other:"
        p  ""
        h2 "Impact"
        p  ""
        p  "-> Confirmed? Add to Findings page."
    )
}

function Build-XSS {
    param([string]$AppName)
    $rows = @( @("?","","","","",""), @("?","","","","","") )
    return xbody @(
        h1 ("XSS - " + $AppName)
        p  ("App: " + $AppName)
        p  "Created because:"
        p  ""
        h2 "Candidates"
        xtable @("Status","Method","Endpoint","Param / Sink","Type","Context") $rows
        p  "Type: reflected   stored   DOM"
        p  "Context: HTML body   HTML attr   JS string   JS block   URL   CSS   JSON response"
        p  ""
        h2 "What Tipped You Off"
        p  "Input:"
        p  "Output:"
        p  ""
        h2 "Filter / WAF Behavior"
        p  "What gets blocked, what slips through."
        p  ""
        h2 "Working Payloads"
        p  "Context:"
        mono "Payload:"
        p  "Result:"
        p  "------------------------------------"
        p  "Context:"
        mono "Payload:"
        p  "Result:"
        p  ""
        h2 "Execution Notes"
        cb "document.domain confirmed"
        cb "Cookies accessible (document.cookie)"
        cb "HttpOnly on session cookie - theft not viable"
        cb "Stored - roles that see the payload:"
        p  ""
        h2 "CSP"
        mono "CSP Header:"
        cb  "No CSP"
        cb  "Bypassable - bypass path:"
        cb  "Blocks exploitation"
        p  ""
        h2 "DOM XSS"
        xtable @("Source","Sink") @( @("location.search","innerHTML"), @("location.hash","eval()"), @("","") )
        p  ""
        h2 "Impact"
        p  ""
        p  "-> Confirmed? Add to Findings page."
    )
}

function Build-SSTI {
    param([string]$AppName)
    $rows = @( @("?","","","",""), @("?","","","","") )
    $probes = @(
        @("{{7*7}}", ""),
        @('${7*7}',  ""),
        @("#{7*7}",  ""),
        @("{{7*'7'}}", ""),
        @("<%= 7*7 %>", "")
    )
    return xbody @(
        h1 ("SSTI - " + $AppName)
        p  ("App: " + $AppName)
        p  "Created because:"
        p  ""
        h2 "Candidates"
        xtable @("Status","Method","Endpoint","Param(s)","Engine Suspected") $rows
        p  ""
        h2 "Engine Identification"
        xtable @("Probe","Result") $probes
        p  "Confirmed engine:"
        p  "Evidence:"
        p  ""
        h2 "Working Payloads"
        p  "Engine:"
        mono "Payload:"
        p  "Result:"
        p  ""
        h2 "RCE Reached"
        cb "OS command execution confirmed"
        p  "id / whoami output:"
        cb "File read viable"
        cb "Shell / beacon dropped"
        p  ""
        h2 "Impact"
        p  ""
        p  "-> Confirmed? Add to Findings page."
    )
}

function Build-SSRF {
    param([string]$AppName)
    $rows  = @( @("?","","","",""), @("?","","","","") )
    $reach = @( @("169.254.169.254","80","Cloud metadata",""), @("","","","") )
    return xbody @(
        h1 ("SSRF - " + $AppName)
        p  ("App: " + $AppName)
        p  "Created because:"
        p  ""
        h2 "Candidates"
        xtable @("Status","Method","Endpoint","Param(s)","Type") $rows
        p  "Type: full-response   blind   partial (error-based)"
        p  ""
        h2 "OOB Confirmation"
        p  "Collaborator / interactsh payload:"
        p  "Callback received:   Yes / No"
        p  "Protocol:   HTTP / DNS"
        p  "Source IP:"
        p  ""
        h2 "Internal Reach"
        xtable @("Host / IP","Port","Service","Result") $reach
        p  ""
        h2 "Cloud Metadata"
        cb "AWS IAM creds retrieved:"
        cb "GCP metadata reachable"
        cb "Azure metadata reachable"
        p  ""
        h2 "Protocols / Bypasses That Worked"
        cb "file:///etc/passwd"
        cb "dict://"
        cb "gopher://"
        cb "Decimal IP"
        cb "IPv6"
        cb "Open redirect chain"
        cb "Other:"
        p  ""
        h2 "Impact"
        p  ""
        p  "-> Confirmed? Add to Findings page."
    )
}

function Build-IDOR {
    param([string]$AppName)
    $rows     = @( @("?","","","",""), @("?","","","","") )
    $accounts = @( @("","standard",""), @("","standard",""), @("","admin","") )
    return xbody @(
        h1 ("IDOR / Broken Access Control - " + $AppName)
        p  ("App: " + $AppName)
        p  "Created because:"
        p  ""
        h2 "Candidates"
        xtable @("Status","Method","Endpoint","Object Ref / Param","AC Type") $rows
        p  "AC Type: IDOR-horizontal   IDOR-vertical   BOLA   BFLA   mass-assignment"
        p  ""
        h2 "Test Accounts"
        xtable @("Account","Role","User ID") $accounts
        p  ""
        h2 "Object Reference Pattern"
        p  "Sequential int / GUID / username / hash - and where it is exposed."
        p  ""
        h2 "Confirmed Bypasses"
        p  "Type:   Horizontal / Vertical / Mass assignment"
        mono "Request:"
        p  "Result:"
        p  "------------------------------------"
        p  "Type:   Horizontal / Vertical / Mass assignment"
        mono "Request:"
        p  "Result:"
        p  ""
        h2 "Scope of Exposed Data"
        cb "PII"
        cb "Financial"
        cb "Auth material"
        cb "Files"
        cb "Admin data"
        p  ""
        h2 "Impact"
        p  "Blast radius - how many records, read vs. modify/delete."
        p  ""
        p  "-> Confirmed? Add to Findings page."
    )
}

function Build-LFI {
    param([string]$AppName)
    $rows     = @( @("?","","","",""), @("?","","","","") )
    $traversal = @(
        @("../../../../etc/passwd",""),
        @("..%2F..%2F..%2Fetc%2Fpasswd",""),
        @("....//....//etc/passwd",""),
        @("%252e%252e%252fetc%252fpasswd","")
    )
    $reads = @(
        @("/etc/passwd",""),
        @("/proc/self/environ",""),
        @("App config:",""),
        @("Log file:","")
    )
    return xbody @(
        h1 ("Path Traversal / LFI - " + $AppName)
        p  ("App: " + $AppName)
        p  "Created because:"
        p  ""
        h2 "Candidates"
        xtable @("Status","Method","Endpoint","Param(s)","Type") $rows
        p  "Type: path-traversal   LFI   RFI   file-download   zip-slip"
        p  ""
        h2 "Traversal Behavior"
        xtable @("Probe","Result") $traversal
        p  "Base dir (inferred):"
        p  "Depth needed:"
        p  "Encoding bypass required:"
        p  ""
        h2 "Confirmed File Reads"
        xtable @("Path","Output") $reads
        p  ""
        h2 "RCE Path"
        cb "Log poisoning - log location:   injection point:   execution confirmed:"
        cb "PHP session include"
        cb "Upload + include"
        cb "PHP filter chain"
        p  ""
        h2 "Impact"
        p  ""
        p  "-> Confirmed? Add to Findings page."
    )
}

function Build-Auth {
    param([string]$AppName)
    $rows = @(
        @("?","POST","/login","Credential submit","Brute / lockout"),
        @("?","POST","/forgot-password","Password reset","Reset poisoning"),
        @("?","","","","")
    )
    return xbody @(
        h1 ("Auth / JWT / OAuth - " + $AppName)
        p  ("App: " + $AppName)
        p  "Created because:"
        p  ""
        h2 "Mechanism"
        p  "Type:   Session cookie / JWT / OAuth 2.0 / API key / Basic / Custom"
        p  "MFA:    Yes / No / Partial"
        p  "SSO:    Yes / No     Provider:"
        p  ""
        h2 "Candidates"
        xtable @("Status","Method","Endpoint","Surface","Issue Suspected") $rows
        p  ""
        h2 "JWT"
        p  "Token (header.payload only):"
        mono ""
        p  "Header:"
        p  "Payload:"
        p  ""
        cb "alg:none - Result:"
        cb "Weak secret cracked - Result:"
        cb "RS256 to HS256 confusion - Result:"
        cb "kid injection - Result:"
        cb "jwk/jku/x5u injection - Result:"
        cb "Expired token accepted - Result:"
        cb "Not invalidated on logout - Result:"
        p  ""
        h2 "OAuth"
        p  "Flow:   authorization_code / implicit / client_credentials"
        p  ""
        cb "redirect_uri manipulation - Result:"
        cb "state param absent / not validated - Result:"
        cb "Auth code reuse - Result:"
        cb "Token in Referer - Result:"
        cb "Account pre-hijacking - Result:"
        p  ""
        h2 "Password Reset"
        cb "Host header poisoning - Result:"
        cb "Token entropy sufficient"
        cb "Token expires / single-use"
        cb "Username enumeration possible"
        cb "Rate limiting absent"
        p  ""
        h2 "Session / Brute"
        p  "Cookie flags:"
        cb "HttpOnly"
        cb "Secure"
        cb "SameSite"
        p  ""
        cb "Lockout present"
        cb "Lockout bypassable - Method:"
        cb "MFA bypassable - Method:"
        p  ""
        h2 "Impact"
        p  ""
        p  "-> Confirmed? Add to Findings page."
    )
}

function Build-XXE {
    param([string]$AppName)
    $rows  = @( @("?","","","",""), @("?","","","","") )
    $reads = @( @("/etc/passwd",""), @("Other:","") )
    $xxePayload1 = "<?xml version=" + [char]34 + "1.0" + [char]34 + "?>"
    $xxePayload2 = "<!DOCTYPE foo [<!ENTITY xxe SYSTEM " + [char]34 + "file:///etc/passwd" + [char]34 + ">]>"
    $xxePayload3 = "<root><data>&xxe;</data></root>"
    return xbody @(
        h1 ("XXE - " + $AppName)
        p  ("App: " + $AppName)
        p  "Created because:"
        p  ""
        h2 "Candidates"
        xtable @("Status","Method","Endpoint","Content-Type","Type") $rows
        p  "Type: in-band   blind-OOB   XInclude   SVG upload   XLSX/DOCX upload"
        p  ""
        h2 "What Tipped You Off"
        p  "XML accepted, verbose parser errors, file upload processing XML internally."
        p  ""
        h2 "Exploitation"
        p  "In-band file read:"
        mono $xxePayload1
        mono $xxePayload2
        mono $xxePayload3
        p  "Result:"
        p  ""
        p  "SSRF via XXE:"
        p  "Target URL:"
        p  "Result:"
        p  ""
        p  "Blind OOB:"
        p  "Collaborator payload:"
        p  "Callback received:   Yes / No"
        p  ""
        p  "XInclude (no DOCTYPE control):"
        p  "Result:"
        p  ""
        h2 "Confirmed File Reads"
        xtable @("Path","Output") $reads
        p  ""
        h2 "Impact"
        p  ""
        p  "-> Confirmed? Add to Findings page."
    )
}

# ---------------------------------------------------------------------------
# ONENOTE COM WIRING
# ---------------------------------------------------------------------------

function Get-OneNoteApp {
    Write-Host "[*] Connecting to OneNote..." -ForegroundColor Cyan
    try {
        $app = New-Object -ComObject OneNote.Application
        Write-Host "[+] Connected to OneNote." -ForegroundColor Green
        return $app
    }
    catch {
        Write-Error "Could not connect to OneNote COM object. Make sure OneNote desktop is installed and open."
        exit 1
    }
}

function New-Notebook {
    param($App, [string]$Name, [string]$Path)
    $fullPath = [System.IO.Path]::Combine($Path, $Name)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
    Write-Host "[*] Creating notebook: $Name" -ForegroundColor Cyan
    $nbId = ""
    $App.OpenHierarchy($fullPath, "", [ref]$nbId,
        [Microsoft.Office.Interop.OneNote.CreateFileType]::cftNotebook)
    return $nbId
}

function New-SectionGroup {
    param($App, [string]$ParentId, [string]$Name)
    $sgId = ""
    $App.OpenHierarchy($Name, $ParentId, [ref]$sgId,
        [Microsoft.Office.Interop.OneNote.CreateFileType]::cftFolder)
    return $sgId
}

function New-Section {
    param($App, [string]$ParentId, [string]$Name)
    $sId = ""
    $App.OpenHierarchy($Name, $ParentId, [ref]$sId,
        [Microsoft.Office.Interop.OneNote.CreateFileType]::cftSection)
    return $sId
}

function New-Page {
    param($App, [string]$SectionId, [string]$Title)
    $pageId = ""
    $App.CreateNewPage($SectionId, [ref]$pageId,
        [Microsoft.Office.Interop.OneNote.NewPageStyle]::npsBlankPageWithTitle)
    return $pageId
}

function Set-PageContent {
    param($App, [string]$PageId, [string]$Title, [string]$BodyXml)
    $xml = New-PageXml -PageId $PageId -Title $Title -BodyXml $BodyXml
    $App.UpdatePageContent($xml, [DateTime]::MinValue,
        [Microsoft.Office.Interop.OneNote.XMLSchema]::xs2013, $false)
}

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "======================================================"  -ForegroundColor Yellow
Write-Host "  Pentest Notebook Builder"                              -ForegroundColor Yellow
Write-Host "======================================================"  -ForegroundColor Yellow
Write-Host ""
Write-Host "  Notebook : $NotebookName"
Write-Host "  Apps     : $($AppNames -join ', ')"
Write-Host "  Path     : $NotebookPath"
Write-Host ""

$onenote = Get-OneNoteApp

$nbId = New-Notebook -App $onenote -Name $NotebookName -Path $NotebookPath
Start-Sleep -Milliseconds 500

# Overview section
Write-Host "[*] Creating _Overview section..." -ForegroundColor Cyan
$overviewId = New-Section -App $onenote -ParentId $nbId -Name "_Overview"
Start-Sleep -Milliseconds 300

Write-Host "    [+] Scope and Stack" -ForegroundColor Green
$pageId = New-Page -App $onenote -SectionId $overviewId -Title "Scope and Stack"
Set-PageContent -App $onenote -PageId $pageId -Title "Scope and Stack" -BodyXml (Build-ScopeAndStack)
Start-Sleep -Milliseconds 200

Write-Host "    [+] Findings" -ForegroundColor Green
$pageId = New-Page -App $onenote -SectionId $overviewId -Title "Findings"
Set-PageContent -App $onenote -PageId $pageId -Title "Findings" -BodyXml (Build-Findings)
Start-Sleep -Milliseconds 200

# One section group per app
foreach ($appName in $AppNames) {
    Write-Host "[*] Creating section group: $appName" -ForegroundColor Cyan
    $sgId = New-SectionGroup -App $onenote -ParentId $nbId -Name $appName
    Start-Sleep -Milliseconds 300

    $surfaceId = New-Section -App $onenote -ParentId $sgId -Name "Surface"
    Start-Sleep -Milliseconds 200
    Write-Host "    [+] Surface" -ForegroundColor Green
    $pageId = New-Page -App $onenote -SectionId $surfaceId -Title ("Surface - " + $appName)
    Set-PageContent -App $onenote -PageId $pageId -Title ("Surface - " + $appName) -BodyXml (Build-Surface -AppName $appName)
    Start-Sleep -Milliseconds 200

    $vulnClasses = @("SQLi","XSS","SSTI","SSRF","XXE","IDOR","LFI","Auth")

    foreach ($vcName in $vulnClasses) {
        Write-Host "    [+] $vcName" -ForegroundColor Green
        $vcId = New-Section -App $onenote -ParentId $sgId -Name $vcName
        Start-Sleep -Milliseconds 200
        $pageId  = New-Page -App $onenote -SectionId $vcId -Title ($vcName + " - " + $appName)
        $pageTitle = $vcName + " - " + $appName
        switch ($vcName) {
            "SQLi" { $bodyXml = Build-SQLi -AppName $appName }
            "XSS"  { $bodyXml = Build-XSS  -AppName $appName }
            "SSTI" { $bodyXml = Build-SSTI -AppName $appName }
            "SSRF" { $bodyXml = Build-SSRF -AppName $appName }
            "XXE"  { $bodyXml = Build-XXE  -AppName $appName }
            "IDOR" { $bodyXml = Build-IDOR -AppName $appName }
            "LFI"  { $bodyXml = Build-LFI  -AppName $appName }
            "Auth" { $bodyXml = Build-Auth -AppName $appName }
        }
        Set-PageContent -App $onenote -PageId $pageId -Title $pageTitle -BodyXml $bodyXml
        Start-Sleep -Milliseconds 200
    }
}

Write-Host ""
Write-Host "======================================================"  -ForegroundColor Green
Write-Host "  Done. Notebook ready in OneNote."                      -ForegroundColor Green
Write-Host "======================================================"  -ForegroundColor Green
Write-Host ""
Write-Host "Structure created:"  -ForegroundColor Yellow
Write-Host "  $NotebookName"
Write-Host "  +-- _Overview"
Write-Host "  |   +-- Scope and Stack"
Write-Host "  |   +-- Findings"
foreach ($a in $AppNames) {
    Write-Host "  +-- $a"
    Write-Host "      +-- Surface"
    Write-Host "      +-- SQLi / XSS / SSTI / SSRF / XXE / IDOR / LFI / Auth"
}
Write-Host ""
