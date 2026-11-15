<#
.SYNOPSIS
  Writes Jameel Copilot's settings to the registry on a developer or test PC, the same way
  the MSI does. Run from an elevated PowerShell.

.DESCRIPTION
  For running a developer build of the add-in, where the MSI is not used.
  Company PCs get their settings from the MSI (Deploy\README.txt).

  The API key is read from a file ({ "apiKey": "..." }, default jameelcopilot.local.json
  in the repository root) or typed hidden. It is stored encrypted for this
  machine, never printed, and never placed on a command line.

.EXAMPLE
  .\Installer\Set-JameelCopilotConfig.ps1 -ApiBaseUrl https://llm.example.internal/v1 -Model my-model
.EXAMPLE
  .\Installer\Set-JameelCopilotConfig.ps1 -ApiBaseUrl https://llm.example.internal/v1 -Model my-model -WriteTools outlook_create_draft
#>
param(
    [Parameter(Mandatory = $true)][string]$ApiBaseUrl,
    [Parameter(Mandatory = $true)][string]$Model,
    [string]$ApiKeyFile,
    [string]$WriteTools = '',
    [string]$ReasoningEffort = '',
    [string]$MaxBulkExportRows = '',
    # 0 to 2 (default 0.5), or "off" to send none.
    [string]$Temperature = '',
    [string]$TemperatureMax = '',
    [string]$MaxOutputTokens = ''
)

$ErrorActionPreference = 'Stop'

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not (New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run this from an elevated PowerShell: the settings live in HKLM.'
}
if ($ApiBaseUrl -notmatch '^https://') { throw 'ApiBaseUrl must begin with https://' }

if (-not $ApiKeyFile) { $ApiKeyFile = Join-Path (Split-Path $PSScriptRoot -Parent) 'jameelcopilot.local.json' }
$key = $null
if (Test-Path -LiteralPath $ApiKeyFile) { $key = (Get-Content -LiteralPath $ApiKeyFile -Raw | ConvertFrom-Json).apiKey }
if (-not $key) {
    $secure = Read-Host 'API key (typing is hidden)' -AsSecureString
    $key = (New-Object System.Net.NetworkCredential('', $secure)).Password
}
if (-not $key) { throw 'No API key was given.' }

# The format the add-in reads (ApiKeyProtector): DPAPI, machine scope, base64.
Add-Type -AssemblyName System.Security
$protected = [Convert]::ToBase64String([Security.Cryptography.ProtectedData]::Protect(
    [Text.Encoding]::UTF8.GetBytes($key), $null, [Security.Cryptography.DataProtectionScope]::LocalMachine))

# The 64-bit view, where the add-in looks even when Outlook is 32-bit.
$hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
$settings = $hklm.CreateSubKey('SOFTWARE\JameelCopilot')
try {
    $settings.SetValue('ApiBaseUrl', $ApiBaseUrl.TrimEnd('/'))
    $settings.SetValue('ApiKeyProtected', $protected)
    $settings.SetValue('Model', $Model)
    $settings.SetValue('ReasoningEffort', $ReasoningEffort)
    $settings.SetValue('EnabledWriteTools', $WriteTools)
    $settings.SetValue('MaxBulkExportRows', $MaxBulkExportRows)
    $settings.SetValue('Temperature', $Temperature)
    $settings.SetValue('TemperatureMax', $TemperatureMax)
    $settings.SetValue('MaxOutputTokens', $MaxOutputTokens)
}
finally {
    $settings.Dispose()
    $hklm.Dispose()
}
Write-Host 'Jameel Copilot settings written to HKLM\SOFTWARE\JameelCopilot. Restart Outlook to pick them up.'
