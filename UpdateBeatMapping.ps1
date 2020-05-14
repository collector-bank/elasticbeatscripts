
[CmdletBinding()]
Param(
    [string] $ElasticUrl,
    [string] $ElasticClusterName,
    [string] $ElasticUsername,
    [string] $ElasticPassword,
    [string] $SlackUrls,
    [string] $SendGridApiKey,
    [string] $EmailRecipients
)

Set-StrictMode -v latest
$ErrorActionPreference = "Stop"

function Main([string] $elasticurl, [string] $elasticclustername, [string] $elasticusername, [string] $elasticpassword, [string] $slackurls, [string] $sendgridapikey, [string] $emailrecipients) {
    Log "Starting..."

    if (!$elasticurl) {
        $elasticurl = Get-AutomationVariable "UpdateBeatMapping_ElasticUrl"
        if (!$elasticurl) {
            Write-Error "Variable UpdateBeatMapping_ElasticUrl not set."
            exit 1
        }
    }
    if (!$elasticclustername) {
        try {
            $elasticclustername = Get-AutomationVariable "UpdateBeatMapping_ElasticClusterName"
        }
        catch {
            $elasticclustername = $null
        }
    }
    if (!$elasticusername) {
        $elasticusername = Get-AutomationVariable "UpdateBeatMapping_ElasticUsername"
        if (!$elasticusername) {
            Write-Error "Variable UpdateBeatMapping_ElasticUsername not set."
            exit 1
        }
    }
    if (!$elasticpassword) {
        $elasticpassword = Get-AutomationVariable "UpdateBeatMapping_ElasticPassword"
        if (!$elasticpassword) {
            Write-Error "Variable UpdateBeatMapping_ElasticPassword not set."
            exit 1
        }
    }
    if (!$slackurls) {
        try {
            $slackurls = Get-AutomationVariable "UpdateBeatMapping_SlackUrls"
        }
        catch {
            $slackurls = $null
        }
    }
    if ($slackurls) {
        Log "Got slack urls."
    }
    if (!$sendgridapikey) {
        try {
            $sendgridapikey = Get-AutomationVariable "SendGridApiKey"
        }
        catch {
            $sendgridapikey = $null
        }
    }
    if ($sendgridapikey) {
        Log "Got send grid api key."
    }
    if (!$emailrecipients) {
        try {
            $emailrecipients = Get-AutomationVariable "UpdateBeatMapping_EmailRecipients"
        }
        catch {
            $emailrecipients = $null
        }
    }
    if ($emailrecipients) {
        Log "Got email recipients."
    }

    Get-Dependencies

    $watch = [Diagnostics.Stopwatch]::StartNew()

    $downloadUrls = @{ }
    [string[]] $beatnames = "filebeat", "heartbeat", "metricbeat", "winlogbeat"

    foreach ($beatname in $beatnames) {
        [string] $pageurl = "https://www.elastic.co/downloads/beats/$beatname"
        Log "Getting download url from page: '$pageurl'"
        [string] $downloadUrl = Get-DownloadUrl $pageurl $beatname
        if ($downloadUrl) {
            $downloadUrls[$beatname] = $downloadUrl
        }
    }

    [int] $count = 0
    foreach ($beatname in $downloadUrls.Keys) {
        $count++
    }
    Log ("Got $count download urls.")
    foreach ($downloadUrl in $downloadUrls.Values) {
        Log "'$downloadUrl'"
    }

    $mappings = Get-Mappings $elasticurl $elasticusername $elasticpassword

    [int] $count = 0
    foreach ($mapping in $mappings) {
        $count++
    }
    Log ("Got $count mappings.")
    foreach ($mapping in $mappings | Sort-Object "Name") {
        Log "'$($mapping.Name)'"
    }

    [string[]] $updates = @()
    foreach ($beatname in $beatnames) {
        [string] $downloadUrl = $downloadUrls[$beatname]

        if ($downloadUrl) {
            $mapping = $mappings | Where-Object { $downloadUrl.EndsWith("/$($_.Name)-windows-x86_64.zip") }

            if ($mapping) {
                Log "$beatname already updated: '$($mapping.Name)' '$downloadUrl'"
            }
            else {
                [string] $beatversion = [IO.Path]::GetFileNameWithoutExtension($downloadUrl)
                $updates += "$beatversion ($downloadUrl)"
                Setup-BeatMapping $beatname $downloadUrl $elasticurl $elasticusername $elasticpassword
            }
        }
        else {
            Log "$($beatname): Couldn't find download url."
        }
    }

    if ($updates.Count -eq 0) {
        Log "No updates applied, not sending any notifications."
    }
    else {
        if ($elasticclustername) {
            [string] $message = "Updated mapping to:`n$elasticclustername ($elasticurl)`nBeats:`n" + ($updates -join "`n")
        }
        else {
            [string] $message = "Updated mapping to:`n'$elasticurl'`nBeats:`n" + ($updates -join "`n")
        }
        if ($slackurls) {
            [string[]] $urls = $slackurls.Split("`n,;".ToCharArray(), [StringSplitOptions]::RemoveEmptyEntries)
            foreach ($slackurl in $urls) {
                Log "Sending slack message to: '$slackurl'"
                Invoke-WebRequest -UseBasicParsing $slackurl -Method Post -Headers @{ "Content-Type" = "application/json" } -Body "{ `"text`":`"$message`" }" | Out-Null
            }
        }
        if ($sendgridapikey -and $emailrecipients.Split("`n")) {
            [string[]] $recipients = $emailrecipients.Split("`n,;".ToCharArray(), [StringSplitOptions]::RemoveEmptyEntries)
            [string] $subject = "Updated beat mapping"
            foreach ($recipient in $recipients) {
                $client = New-Object SendGrid.SendGridClient -ArgumentList $sendgridapikey
                $from = New-Object SendGrid.Helpers.Mail.EmailAddress -ArgumentList "no_reply@collector.se"
                $to = New-Object SendGrid.Helpers.Mail.EmailAddress -ArgumentList $recipient
                $msg = [SendGrid.Helpers.Mail.MailHelper]::CreateSingleEmail($from, $to, $subject, $message, "")
                Log "Sending email to: '$recipient'"
                $response = $client.SendEmailAsync($msg)
                $response.GetAwaiter().GetResult() | Out-Null
            }
        }
    }

    Log "Done: $($watch.Elapsed)" Cyan
}

function Get-Dependencies() {
    Log "Current dir: '$((pwd).Path)'"

    Import-Nuget "https://globalcdn.nuget.org/packages/newtonsoft.json.9.0.1.nupkg" "5D96EE51B2AFF592039EEBC2ED203D9F55FDDF9C0882FB34D3F0E078374954A5"

    Import-Nuget "https://globalcdn.nuget.org/packages/sendgrid.9.12.0.nupkg" "E1B10B0C2A99C289227F0D91F5335D08CDA4C3203B492EBC1B0D312B748A3C04"
}

function Import-Nuget([string] $moduleurl, [string] $dllhash) {
    [string] $nugetfile = Split-Path -Leaf $moduleurl
    [int] $end = $nugetfile.IndexOf(".")
    if ($end -lt 0) {
        [string] $shortname = $nugetfile
    }
    else {
        [string] $shortname = $nugetfile.Substring(0, $end)
    }
    [string] $dllfile = "$($shortname).dll"

    if (Test-Path $dllfile) {
        [string] $hash = (Get-FileHash $dllfile).Hash
        if ($hash -eq $dllhash) {
            Log "Using binary that's already downloaded: '$dllfile'"

            [string] $dllpath = Join-Path (pwd).Path $dllfile
            Log "Importing dllfile: '$dllpath'"
            Import-Module $dllpath
            return
        }
        else {
            Log "Deleting binary: '$dllfile' with wrong hash: '$hash'"
            del $dllfile
        }
    }

    Log "Downloading nuget file: '$moduleurl' -> '$nugetfile'"
    Invoke-WebRequest -UseBasicParsing $moduleurl -OutFile $nugetfile

    [string] $zipfile = "$($shortname).zip"

    if (Test-Path $zipfile) {
        Log "Deleting old zipfile: '$zipfile'"
        del $zipfile
    }

    Log "Renaming: '$nugetfile' -> '$zipfile'"
    ren $nugetfile $zipfile

    if (Test-Path $shortname) {
        Log "Deleting old folder: '$shortname'"
        rd -Recurse -Force $shortname
    }

    Log "Extracting: '$zipfile' -> '$shortname'"
    Expand-Archive $zipfile $shortname

    [string] $path = Join-Path $shortname (Join-Path "lib" (Join-Path "netstandard*" "*.dll"))
    Log "Searching path: '$path'"
    if (Test-Path $path) {
        $dllpath = dir $path | Sort-Object FullName -Descending | Select-Object -First 1

        Log "Moving: '$dllpath' -> '$dllfile'"
        move $dllpath $dllfile
    }
    else {
        Write-Error "Didn't find any netstandard dllfile."
        exit 1
    }

    [string] $hash = (Get-FileHash $dllfile).Hash
    if ($hash -ne $dllhash) {
        Write-Error "Couldn't download, wrong hash: '$dllfile': '$hash'"
        exit 1
    }

    [string] $dllpath = Join-Path (pwd).Path $dllfile
    Log "Importing dllfile: '$dllpath'"
    Import-Module $dllpath
}

function Get-DownloadUrl([string] $pageurl, [string] $beatname) {
    [string] $page = Invoke-WebRequest -UseBasicParsing $pageurl

    [string] $regex = '\"https://artifacts.elastic.co/downloads/beats/' + $beatname + '/' + $beatname + '-[0-9\.]+-windows-x86_64.zip\"'
    $matches = @($page | Select-String $regex | Select-Object -First 1 -ExpandProperty "Matches")
    if ($matches.Count -lt 1) {
        return
    }

    [string] $downloadLink = $matches[0].Value

    $downloadLink = $downloadLink.Substring(1, $downloadLink.Length - 2)

    return $downloadLink
}

function Get-Mappings([string] $elasticurl, [string] $elasticusername, [string] $elasticpassword) {
    [string] $queryurl = "$($elasticurl)/_template"
    [string] $auth = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("$($elasticusername):$($elasticpassword)"))

    [string] $result = Invoke-WebRequest -UseBasicParsing $queryurl -Headers @{ "Authorization" = "Basic $auth"; "Content-Type" = "application/json" }

    return [Newtonsoft.Json.Linq.JObject]::Parse($result)
}

function Setup-BeatMapping([string] $beatname, [string] $downloadUrl, [string] $elasticurl, [string] $elasticusername, [string] $elasticpassword) {
    Log "Setuping: $beatname" Cyan

    [string] $zipfile = Split-Path -Leaf $downloadUrl
    if (Test-Path $zipfile) {
        Log "Deleting: '$zipfile'"
        del $zipfile
    }

    [int] $zipfileSize = 10mb

    Robust-Download $downloadUrl $zipfile $zipfileSize
    if (!(Test-Path $zipfile) -or (dir $zipfile).Length -lt $zipfileSize) {
        return
    }

    Update-BeatMapping $zipfile $elasticurl $elasticusername $elasticpassword

    Delete-OldFiles $beatname 3
}

function Robust-Download([string] $url, [string] $outfile, [int] $zipfileSize) {
    if (!$url.StartsWith("https://artifacts.elastic.co/downloads/beats/")) {
        Log "Invalid url: '$url'" Yellow
        return
    }

    for ([int] $tries = 1; !(Test-Path $outfile) -or (dir $outfile).Length -lt $zipfileSize; $tries++) {
        if (Test-Path $outfile) {
            Log "Deleting (try $tries): '$outfile'"
            del $outfile
        }

        Log "Downloading (try $tries): '$url' -> '$outfile'"
        try {
            Invoke-WebRequest -UseBasicParsing $url -OutFile $outfile
        }
        catch {
            Log "Couldn't download (try $tries): '$url' -> '$outfile'" Yellow
            Start-Sleep 5
        }

        if (!(Test-Path $outfile) -or (dir $outfile).Length -lt $zipfileSize) {
            if ($tries -le 10) {
                Log "Couldn't download (try $tries): '$url' -> '$outfile'" Yellow
            }
            else {
                Log "Couldn't download (try $tries): '$url' -> '$outfile'" Yellow
                return
            }
        }
    }

    Log "Downloaded: '$outfile'"
}

function Update-BeatMapping([string] $zipfile, [string] $elasticurl, [string] $elasticusername, [string] $elasticpassword) {
    [string] $folder = [IO.Path]::GetFileNameWithoutExtension($zipfile)

    if (Test-Path $folder) {
        [int] $retries = 1
        do {
            Log "Deleting folder (try $retries): '$folder'"
            rd -Recurse -Force $folder -ErrorAction SilentlyContinue
            if (Test-Path $folder) {
                Start-Sleep 2
                $retries++
            }
            else {
                $retries = 11
            }
        } while ($retries -le 10)
    }
    if (Test-Path $folder) {
        Log "Couldn't delete folder: '$folder'" Yellow
        return
    }

    Log "Extracting: '$zipfile'"
    Expand-Archive $zipfile

    [string] $searchpath = Join-Path (Join-Path $folder "*") "*.exe"
    $exefiles = @(dir $searchpath -File)
    if ($exefiles.Count -ne 1) {
        Log "Couldn't find unique exefile, found $($exefiles.Count) exe files using $($searchpath): '$($exefiles -join "', '")'" Yellow
        return
    }

    [string] $fullname = $exefiles[0].FullName
    pushd (Split-Path $fullname)
    [IO.Directory]::SetCurrentDirectory((pwd).Path) # pushd only sets powershell's fake current dir, not the current dir for the process, which apis as Diagnostics.Process.Start requuires.

    [string] $exefile = Split-Path -Leaf $fullname
    [string] $exeargs = "setup --index-management -E output.logstash.enabled=false -E output.elasticsearch.hosts=[`"$elasticurl`"] -E output.elasticsearch.username=`"$elasticusername`" -E output.elasticsearch.password=`"$elasticpassword`""
    Log "Running: >>>$exefile<<< >>>$($exeargs.Replace($elasticpassword, ("*" * $elasticpassword.Length)))<<<"

    $watch = [Diagnostics.Stopwatch]::StartNew()
    [Diagnostics.Process] $process = [Diagnostics.Process]::Start($exefile, $exeargs)
    $process.WaitForExit()
    popd

    Log "Waited: $($watch.Elapsed)"

    if (!$env:UpdateBeatMapping_DontDeleteFolders) {
        Log "Deleting folder: '$folder'"
        rd -Recurse -Force $folder -ErrorAction SilentlyContinue
    }
}

function Delete-OldFiles([string] $beatname, [int] $keep) {
    $files = @(dir "$($beatname)-*-windows-x86_64.zip" | Sort-Object "LastWriteTime" -Descending | Select-Object -Skip $keep)

    Log "Found $($files.Count) old $beatname files."
    foreach ($file in $files) {
        [string] $filename = $file.FullName
        Log "Deleting: '$filename'"
        del $filename
    }
}

function Log([string] $message, $color) {
    [DateTime] $date = [DateTime]::UtcNow
    [string] $logfile = Join-Path ([IO.Path]::GetTempPath()) "UpdateBeatMapping_$($date.ToString("yyyy_MM")).log"
    [string] $annotatedMessage = $date.ToString("yyyy-MM-dd HH:mm:ss") + ": $message"
    Add-Content $logfile $annotatedMessage

    Write-Output $message

    if ($color) {
        #Write-Host $message -f $color
    }
    else {
        #Write-Host $message -f Green
    }
}

Main $ElasticUrl $ElasticClusterName $ElasticUsername $ElasticPassword $SlackUrls $SendGridApiKey $EmailRecipients
