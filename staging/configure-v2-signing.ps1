# Windows PowerShell 5.1+ / PowerShell 7. Requires authenticated GitHub CLI (gh).
# Run locally, never in chat/CI. Only the dedicated V2 environment is changed.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$CertificatePath,
    [Parameter(Mandatory=$true)][string]$ProvisioningProfilePath
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$Repo = 'ionut290/hera-app'
$EnvName = 'ios-varga-v2-staging'
$Team = 'NJ64236Q54'
$Bundle = 'it.vargacantieri.staging'
$ProfileName = 'Varga Cantieri 2 Staging TestFlight'
$gh = (Get-Command gh -CommandType Application -ErrorAction Stop).Source
$login = & $gh api --hostname github.com user --jq '.login'
if ($LASTEXITCODE -ne 0 -or ($login | Out-String).Trim() -ne 'ionut290') {
    throw 'Sign in to GitHub CLI as the repository owner before proceeding.'
}
$environment = & $gh api --hostname github.com "repos/$Repo/environments/$EnvName" --jq '.name'
if ($LASTEXITCODE -ne 0 -or ($environment | Out-String).Trim() -ne $EnvName) {
    throw 'The protected V2 environment does not exist or is inaccessible.'
}
$certFile = Get-Item -LiteralPath $CertificatePath
$profileFile = Get-Item -LiteralPath $ProvisioningProfilePath
if ($certFile.PSIsContainer -or $certFile.Extension -ne '.p12' -or
    $profileFile.PSIsContainer -or $profileFile.Extension -ne '.mobileprovision') {
    throw 'Select the original .p12 certificate and the newly downloaded V2 .mobileprovision.'
}
$certBytes = [IO.File]::ReadAllBytes($certFile.FullName)
$profileBytes = [IO.File]::ReadAllBytes($profileFile.FullName)
# Decode the signed CMS profile locally. Never print its contents.
try { Add-Type -AssemblyName System.Security } catch {}
try { Add-Type -AssemblyName System.Security.Cryptography.Pkcs } catch {}
$cms = New-Object System.Security.Cryptography.Pkcs.SignedCms
$cms.Decode($profileBytes)
$settings = New-Object System.Xml.XmlReaderSettings
$settings.DtdProcessing = [System.Xml.DtdProcessing]::Ignore
$settings.XmlResolver = $null
$stringReader = New-Object System.IO.StringReader ([Text.Encoding]::UTF8.GetString($cms.ContentInfo.Content))
$xmlReader = [System.Xml.XmlReader]::Create($stringReader, $settings)
$xml = New-Object System.Xml.XmlDocument
$xml.XmlResolver = $null
try { $xml.Load($xmlReader) } finally { $xmlReader.Dispose(); $stringReader.Dispose() }
function Get-ValueNode($dict, [string]$key) {
    $node = $dict.SelectSingleNode("key[text()='$key']/following-sibling::*[1]")
    if ($null -eq $node) { throw "Required provisioning field missing: $key" }
    return $node
}
$root = $xml.SelectSingleNode('/plist/dict')
$ent = Get-ValueNode $root 'Entitlements'
$teams = Get-ValueNode $root 'TeamIdentifier'
if ((Get-ValueNode $root 'Name').InnerText -ne $ProfileName -or
    (Get-ValueNode $ent 'application-identifier').InnerText -ne "$Team.$Bundle" -or
    $null -eq $teams.SelectSingleNode("string[text()='$Team']") -or
    (Get-ValueNode $ent 'get-task-allow').Name -ne 'false' -or
    $null -ne $root.SelectSingleNode("key[text()='ProvisionedDevices' or text()='ProvisionsAllDevices']")) {
    throw 'Wrong bundle, team, profile name or distribution type. No secret changed.'
}
$expiry = [DateTimeOffset]::Parse((Get-ValueNode $root 'ExpirationDate').InnerText)
if ($expiry -le [DateTimeOffset]::UtcNow) { throw 'Profile expired.' }
$secure = Read-Host 'Password of the existing Apple Distribution p12' -AsSecureString
$ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
$password = $null
$certificates = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
function Set-ProtectedSecret([string]$name, [string]$value) {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $gh
    $psi.Arguments = "secret set $name --repo $Repo --env $EnvName"
    $psi.UseShellExecute = $false
    $psi.RedirectStandardInput = $true
    $process = [Diagnostics.Process]::Start($psi)
    try {
        $process.StandardInput.Write($value)
        $process.StandardInput.Close()
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) { throw "GitHub rejected secret $name; check environment permissions." }
        Write-Host "Configured $name"
    } finally { $process.Dispose() }
}
try {
    $password = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
    $certificates.Import($certBytes, $password, $flags)
    $allowed = @((Get-ValueNode $root 'DeveloperCertificates').SelectNodes('data') | ForEach-Object {
        [Convert]::ToBase64String([Convert]::FromBase64String($_.InnerText))
    })
    $matching = @($certificates | Where-Object {
        $_.HasPrivateKey -and $_.Subject -match 'Apple Distribution:' -and
        $_.NotAfter -gt [DateTime]::Now -and
        $allowed -contains [Convert]::ToBase64String($_.RawData)
    })
    if ($matching.Count -ne 1) { throw 'The p12 private key/certificate does not match the V2 profile. No secret changed.' }
    # Explicit local consent. This changes only the three V2 environment secrets.
    if ((Read-Host "Type CONFIGURE V2 to store three signing secrets in $EnvName") -cne 'CONFIGURE V2') {
        Write-Host 'Cancelled; no secret changed.'
        return
    }
    Set-ProtectedSecret 'V2_IOS_DISTRIBUTION_CERTIFICATE_BASE64' ([Convert]::ToBase64String($certBytes))
    Set-ProtectedSecret 'V2_IOS_DISTRIBUTION_CERTIFICATE_PASSWORD' $password
    Set-ProtectedSecret 'V2_IOS_PROVISIONING_PROFILE_BASE64' ([Convert]::ToBase64String($profileBytes))
    Write-Host 'V2 signing secrets configured. No workflow enabled, no upload or release.'
} finally {
    foreach ($cert in $certificates) { $cert.Dispose() }
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    $secure.Dispose()
    $password = $null
    [Array]::Clear($certBytes,0,$certBytes.Length)
}
