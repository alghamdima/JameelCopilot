<#
.SYNOPSIS
  Builds JameelCopilot-<version>-x64.msi.

.DESCRIPTION
  1. Builds the add-in in Release (tools\Build.ps1), stamped with the
     version and signed, and runs its tests.
  2. Builds and signs the installer's custom action.
  3. Builds the MSI with WiX 5 and signs it with the same certificate.
  Prints the MSI path and its SHA256.

.PARAMETER Version
  Three-part product version, for example 1.0.0. A higher version upgrades an
  installed lower one.

.PARAMETER CertThumbprint
  A code-signing certificate in Cert:\CurrentUser\My. It signs the add-in's
  DLL and manifests, the custom action and the MSI. Use the company
  certificate for a release; a self-signed one is only good for a pilot.

.PARAMETER TimestampServer
  Optional RFC 3161 timestamp server address, so the signatures stay valid
  after the certificate expires.

.EXAMPLE
  .\Installer\Build-Msi.ps1 -Version 1.0.0 -CertThumbprint <thumbprint>
#>
param(
    [Parameter(Mandatory = $true)][ValidatePattern('^\d+\.\d+\.\d+$')][string]$Version,
    [Parameter(Mandatory = $true)][string]$CertThumbprint,
    [string]$OutDir,
    [string]$TimestampServer,
    [switch]$SkipTests
)

# Not 'Stop': see tools\Build.ps1. Each step checks the tool's exit code.
$ErrorActionPreference = 'Continue'
$repo = Split-Path $PSScriptRoot -Parent
if (-not $OutDir) { $OutDir = Join-Path $repo 'out' }
$logs = Join-Path $OutDir 'logs'
New-Item -ItemType Directory -Force -Path $logs | Out-Null

function Stop-Build([string]$Message, [string]$Log) {
    Write-Host $Message
    if ($Log -and (Test-Path $Log)) {
        Select-String -Path $Log -Pattern 'error ' | Select-Object -First 25 |
            ForEach-Object { Write-Host ('  ' + ($_.Line.Trim() -replace '\s*\[[^\]]*\]\s*$', '')) }
    }
    exit 1
}

# 1. The add-in, from clean, so nothing stale is packaged.
foreach ($dir in 'VSTO2\JameelCopilot\bin\Release', 'VSTO2\JameelCopilot\obj\Release') {
    $path = Join-Path $repo $dir
    if (Test-Path $path) { Remove-Item -Recurse -Force $path }
}
$buildArgs = @{ CertThumbprint = $CertThumbprint; Configuration = 'Release'; LogDir = $logs; Version = $Version }
if ($TimestampServer) { $buildArgs.TimestampServer = $TimestampServer }
if (-not $SkipTests) { $buildArgs.Test = $true }
& (Join-Path $repo 'tools\Build.ps1') @buildArgs
if ($LASTEXITCODE -ne 0) { Stop-Build 'The add-in build or its tests failed; no MSI was built.' $null }

$addInDir = Join-Path $repo 'VSTO2\JameelCopilot\bin\Release'
# Files whose absence does not fail the build or the tests but breaks the
# installed add-in: the manifests, and two DLLs that only the add-in's own
# output must carry (the Excel export's packaging library, the WebView2 loader).
$mustShip = 'JameelCopilot.dll', 'JameelCopilot.vsto', 'JameelCopilot.dll.manifest', 'System.IO.Packaging.dll',
    'runtimes\win-x64\native\WebView2Loader.dll', 'runtimes\win-x86\native\WebView2Loader.dll'
foreach ($needed in $mustShip) {
    if (-not (Test-Path (Join-Path $addInDir $needed))) { Stop-Build "The add-in output has no $needed." $null }
}

# An upgrade and About both rely on the DLL carrying this build's version.
$dllVersion = (Get-Item (Join-Path $addInDir 'JameelCopilot.dll')).VersionInfo.FileVersion
if ($dllVersion -ne "$Version.0") { Stop-Build "JameelCopilot.dll has file version $dllVersion, not $Version.0." $null }

# 2. The custom action (produces JameelCopilot.Installer.Actions.CA.dll), from clean.
foreach ($dir in 'Actions\bin\Release', 'Actions\obj\Release') {
    $path = Join-Path $PSScriptRoot $dir
    if (Test-Path $path) { Remove-Item -Recurse -Force $path }
}
$log = Join-Path $logs 'installer-actions.log'
$actionArgs = @((Join-Path $PSScriptRoot 'Actions\JameelCopilot.Installer.Actions.csproj'), '-c', 'Release', "-p:JcCertThumbprint=$CertThumbprint", '-v', 'q', '-nologo')
if ($TimestampServer) { $actionArgs += "-p:JcTimestampUrl=$TimestampServer" }
dotnet build @actionArgs *> $log
if ($LASTEXITCODE -ne 0) { Stop-Build 'The installer custom action did not build.' $log }
$caDll = Join-Path $PSScriptRoot 'Actions\bin\Release\net472\JameelCopilot.Installer.Actions.CA.dll'
if (-not (Test-Path $caDll)) { Stop-Build "The custom action package was not produced: $caDll" $log }

# Where only signed code may load, the MSI's own signature does not cover a
# DLL once it is on disk. (A self-signed certificate gives UnknownError, which
# is still a signature.)
foreach ($dll in (Join-Path $addInDir 'JameelCopilot.dll'), $caDll) {
    $dllSignature = Get-AuthenticodeSignature -LiteralPath $dll
    if ($dllSignature.Status -eq 'NotSigned' -or $dllSignature.SignerCertificate.Thumbprint -ne $CertThumbprint) {
        Stop-Build "$(Split-Path $dll -Leaf) is not signed with the certificate $CertThumbprint." $null
    }
}

# 3. The MSI.
$log = Join-Path $logs 'msi.log'
dotnet build (Join-Path $PSScriptRoot 'JameelCopilot.Installer.wixproj') -c Release "-p:JcVersion=$Version" "-p:AddInDir=$addInDir" "-p:ActionsCaDll=$caDll" -v m -nologo *> $log
if ($LASTEXITCODE -ne 0) { Stop-Build 'The MSI did not build.' $log }
$built = Get-ChildItem (Join-Path $PSScriptRoot 'bin') -Recurse -Filter "JameelCopilot-$Version-x64.msi" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $built) { Stop-Build 'The MSI file was not found after the build.' $log }
$msi = Join-Path $OutDir $built.Name
Copy-Item -LiteralPath $built.FullName -Destination $msi -Force

# 4. Sign it.
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object Thumbprint -eq $CertThumbprint | Select-Object -First 1
if (-not $cert) { Stop-Build "No code-signing certificate with thumbprint $CertThumbprint in Cert:\CurrentUser\My." $null }
$signArgs = @{ FilePath = $msi; Certificate = $cert; HashAlgorithm = 'SHA256' }
if ($TimestampServer) { $signArgs.TimestampServer = $TimestampServer }
$signature = Set-AuthenticodeSignature @signArgs
# A self-signed certificate signs correctly but is not trusted: UnknownError is expected then.
Write-Host ("Signature: {0} ({1})" -f $signature.Status, $cert.Subject)

Write-Host ("MSI:    {0}" -f $msi)
Write-Host ("SHA256: {0}" -f (Get-FileHash -LiteralPath $msi -Algorithm SHA256).Hash)
