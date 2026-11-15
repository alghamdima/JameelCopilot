<#
.SYNOPSIS
  Builds the add-in and its tests, and optionally runs the tests.

.DESCRIPTION
  The add-in is an old-style VSTO project: it needs MSBuild from Visual Studio
  or Build Tools with the "Visual Studio Tools for Office (VSTO) build tools"
  component. Office does not have to be installed. The test project is
  SDK-style and is built with the .NET SDK, against the add-in just built.

.PARAMETER CertThumbprint
  A code-signing certificate in Cert:\CurrentUser\My that signs the add-in's
  DLL and manifests. For development, create one once:
    New-SelfSignedCertificate -Type CodeSigningCert -Subject "CN=Jameel Copilot dev" -CertStoreLocation Cert:\CurrentUser\My

.EXAMPLE
  .\tools\Build.ps1 -CertThumbprint <thumbprint> -Test
.EXAMPLE
  .\tools\Build.ps1 -CertThumbprint <thumbprint> -Test -Filter "FullyQualifiedName~RibbonTests"
#>
param(
    [Parameter(Mandatory = $true)][string]$CertThumbprint,
    [ValidateSet('Debug', 'Release')][string]$Configuration = 'Debug',
    [switch]$Test,
    [string]$Filter,
    [string]$LogDir,
    # Three-part version stamped on the DLL and its manifests (default 1.0.0).
    [ValidatePattern('^(\d+\.\d+\.\d+)?$')][string]$Version,
    [string]$TimestampServer
)

# Not 'Stop': Windows PowerShell turns every stderr line of a redirected native
# tool into an error record, and a failing test prints to stderr. Each step
# checks the tool's exit code instead.
$ErrorActionPreference = 'Continue'
$repo = Split-Path $PSScriptRoot -Parent
if (-not $LogDir) { $LogDir = Join-Path $repo 'out\logs' }
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null

function Find-MSBuild {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (Test-Path $vswhere) {
        $found = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
        if ($found) { return $found }
    }
    throw 'MSBuild was not found. Install Visual Studio or Build Tools with the VSTO build tools component.'
}

function Show-Problems([string]$Log) {
    Select-String -Path $Log -Pattern ': error |error MSB|error NETSDK' |
        ForEach-Object { ($_.Line -replace '^\s*', '') -replace '\s*\[[^\]]*\]\s*$', '' } |
        Select-Object -Unique | Select-Object -First 40 | ForEach-Object { Write-Host "  $_" }
}

$msbuild = Find-MSBuild
$sln = Join-Path $repo 'VSTO2\JameelCopilot.sln'
$addin = Join-Path $repo 'VSTO2\JameelCopilot\JameelCopilot.csproj'
$tests = Join-Path $repo 'VSTO2\JameelCopilot.Tests\JameelCopilot.Tests.csproj'

# 1. NuGet packages of the add-in (packages.config) into VSTO2\packages.
#    MSBuild skips the SDK-style test project here (NU1503); dotnet restores it in step 3.
$log = Join-Path $LogDir 'restore.log'
& $msbuild $sln /t:Restore /p:RestorePackagesConfig=true /v:m /nologo *> $log
if ($LASTEXITCODE -ne 0) { Write-Host 'Restore FAILED'; Show-Problems $log; exit 1 }

# 2. The add-in.
$log = Join-Path $LogDir 'addin.log'
$addinArgs = @($addin, "/p:Configuration=$Configuration", "/p:ManifestCertificateThumbprint=$CertThumbprint", '/p:JcSignAssembly=true', '/v:m', '/nologo')
if ($Version) { $addinArgs += "/p:JcVersion=$Version" }
if ($TimestampServer) { $addinArgs += "/p:JcTimestampUrl=$TimestampServer" }
& $msbuild @addinArgs *> $log
if ($LASTEXITCODE -ne 0) {
    Write-Host ("Add-in build FAILED: {0} error lines" -f (Select-String -Path $log -Pattern ': error ' | Measure-Object).Count)
    Show-Problems $log
    exit 1
}
Write-Host ("Add-in build OK ({0} warnings)" -f (Select-String -Path $log -Pattern ': warning ' | Measure-Object).Count)

# 3. The tests, against the add-in built above. VSToolsPath is cleared because
#    the .NET SDK cannot load the VSTO targets the add-in project imports.
$log = Join-Path $LogDir 'tests-build.log'
dotnet build $tests -c $Configuration -p:BuildProjectReferences=false '-p:VSToolsPath=' -v q -nologo *> $log
if ($LASTEXITCODE -ne 0) {
    Write-Host ("Tests build FAILED: {0} error lines" -f (Select-String -Path $log -Pattern ': error ' | Measure-Object).Count)
    Show-Problems $log
    exit 1
}
Write-Host 'Tests build OK'

if ($Test) {
    $dll = Join-Path $repo "VSTO2\JameelCopilot.Tests\bin\$Configuration\net472\JameelCopilot.Tests.dll"
    $log = Join-Path $LogDir 'tests-run.log'
    $testArgs = @($dll, '--nologo')
    if ($Filter) { $testArgs += @('--filter', $Filter) }
    dotnet test @testArgs *> $log
    $code = $LASTEXITCODE
    Select-String -Path $log -Pattern '^(Passed|Failed)!|No test matches' | ForEach-Object { Write-Host $_.Line.Trim() }
    Select-String -Path $log -Pattern '^\s+Failed \S' | Select-Object -First 40 | ForEach-Object { Write-Host ('  ' + $_.Line.Trim()) }
    exit $code
}
