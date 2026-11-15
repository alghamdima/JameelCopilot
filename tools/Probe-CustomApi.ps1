<#
.SYNOPSIS
  Asks the Custom API's model for two tool calls, sends the results back, and
  prints what the service answers at each step.

.DESCRIPTION
  Shows whether the model calls tools the way Jameel Copilot expects. Run it from a PC
  that can reach the Custom API (the company network or VPN). Two requests:

    1. A streamed question that needs the tool twice. Every line of the
       answer is printed, then a summary of the tool calls found in it.
    2. The same conversation with the two tool results added, as Jameel Copilot sends it
       after running the tools. The service must accept it and answer in text.

  The address and key come from jameelcopilot.local.json in the repository root:
      { "apiBaseUrl": "https://.../v1", "apiKey": "...", "model": "..." }
  Without that file the script asks for them; the key is typed hidden. The
  key is sent in the Authorization header only. It is never printed and never
  placed on a command line.

  A good run ends with: RESULT: OK. Send the whole output to the developers
  either way; it holds no mail content and no key.

.EXAMPLE
  .\tools\Probe-CustomApi.ps1
.EXAMPLE
  .\tools\Probe-CustomApi.ps1 -Model another-model
#>
param(
    [string]$ApiBaseUrl,
    [string]$Model,
    [string]$SettingsFile
)

$ErrorActionPreference = 'Stop'

if (-not $SettingsFile -and $PSScriptRoot) {
    $SettingsFile = Join-Path (Split-Path $PSScriptRoot -Parent) 'jameelcopilot.local.json'
}

$key = $null
if ($SettingsFile -and (Test-Path -LiteralPath $SettingsFile)) {
    $saved = Get-Content -LiteralPath $SettingsFile -Raw | ConvertFrom-Json
    if (-not $ApiBaseUrl) { $ApiBaseUrl = $saved.apiBaseUrl }
    if (-not $Model) { $Model = $saved.model }
    $key = $saved.apiKey
}
if (-not $Model) { $Model = Read-Host 'Model name, as the Custom API knows it' }
if (-not $ApiBaseUrl) { $ApiBaseUrl = Read-Host 'Custom API base address (https://.../v1)' }
if (-not $key) {
    $secure = Read-Host 'API key (typing is hidden)' -AsSecureString
    $key = (New-Object System.Net.NetworkCredential('', $secure)).Password
}
if ($ApiBaseUrl -notmatch '^https://') { throw 'The base address must start with https://' }
if (-not $key) { throw 'No API key was given.' }

Add-Type -AssemblyName System.Net.Http
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

$tools = @(@{
    type     = 'function'
    function = @{
        name        = 'get_time'
        description = 'Returns the current time in a city.'
        parameters  = @{ type = 'object'; properties = @{ city = @{ type = 'string' } }; required = @('city') }
    }
})
$question = @{ role = 'user'; content = 'What time is it in Riyadh and in Jeddah? Call the tool once for each city, both in the same step.' }

$handler = New-Object System.Net.Http.HttpClientHandler
$handler.AllowAutoRedirect = $false
$client = New-Object System.Net.Http.HttpClient($handler)
$client.Timeout = [TimeSpan]::FromSeconds(120)
$url = $ApiBaseUrl.TrimEnd('/') + '/chat/completions'

# Sends one streamed request, prints each line, and returns the status, the
# text and the tool calls put together the way Jameel Copilot's reader does.
function Send-Round([object[]]$Messages) {
    $body = @{ model = $Model; stream = $true; tool_choice = 'auto'; messages = $Messages; tools = $tools } |
        ConvertTo-Json -Depth 12 -Compress
    $request = New-Object System.Net.Http.HttpRequestMessage([System.Net.Http.HttpMethod]::Post, $url)
    $request.Headers.Authorization = New-Object System.Net.Http.Headers.AuthenticationHeaderValue('Bearer', $key)
    $request.Content = New-Object System.Net.Http.StringContent($body, [Text.Encoding]::UTF8, 'application/json')
    $response = $client.SendAsync($request, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
    Write-Host ("HTTP {0} {1}  content-type: {2}" -f [int]$response.StatusCode, $response.StatusCode, $response.Content.Headers.ContentType)

    $round = @{ Status = [int]$response.StatusCode; Text = ''; Calls = @(); Indexes = @(); Done = $false }
    $open = @{}
    $reader = New-Object System.IO.StreamReader($response.Content.ReadAsStreamAsync().GetAwaiter().GetResult())
    while ($null -ne ($line = $reader.ReadLine())) {
        if ($line.Length -eq 0) { continue }
        Write-Host $line
        if (-not $line.StartsWith('data:')) { continue }
        $payload = $line.Substring(5).Trim()
        if ($payload -eq '[DONE]') { $round.Done = $true; continue }
        try { $chunk = $payload | ConvertFrom-Json } catch { continue }
        if (-not $chunk.choices) { continue }
        $delta = $chunk.choices[0].delta
        if (-not $delta) { continue }
        if ($delta.content) { $round.Text += $delta.content }
        foreach ($fragment in @($delta.tool_calls)) {
            if (-not $fragment) { continue }
            $index = 0
            if ($null -ne $fragment.index) { $index = [int]$fragment.index; $round.Indexes += $index } else { $round.Indexes += 'none' }
            $call = $open[$index]
            if (-not $call -or ($fragment.id -and $call.id -and $call.id -ne $fragment.id)) {
                $call = @{ id = ''; name = ''; arguments = '' }
                $open[$index] = $call
                $round.Calls += $call
            }
            if ($fragment.id -and -not $call.id) { $call.id = $fragment.id }
            if ($fragment.function.name -and -not $call.name) { $call.name = $fragment.function.name }
            if ($fragment.function.arguments) { $call.arguments += $fragment.function.arguments }
        }
    }
    return $round
}

$problems = @()
try {
    Write-Host "POST $url  (model $Model)"
    Write-Host '--- Request 1: a question that needs the tool twice ---'
    $first = Send-Round @($question)
    Write-Host ''
    Write-Host ("Tool calls found: {0}; index values seen: {1}; end marker: {2}" -f $first.Calls.Count, (($first.Indexes | Select-Object -Unique) -join ','), $first.Done)
    $n = 0
    foreach ($call in $first.Calls) {
        $n++
        $hasId = [bool]$call.id
        Write-Host ("  call {0}: name={1} arguments={2} has id={3}" -f $n, $call.name, $call.arguments, $hasId)
        if (-not $hasId) { $call.id = "call_probe_$n" }
        try { $null = $call.arguments | ConvertFrom-Json } catch { $problems += "call $n has arguments that are not JSON" }
    }
    if ($first.Status -ne 200) { $problems += "request 1 was answered HTTP $($first.Status)" }
    elseif ($first.Calls.Count -eq 0) { $problems += 'the model did not call the tool' }
    elseif ($first.Calls.Count -eq 1) { Write-Host '  NOTE: one call only; the model made no parallel calls in this run.' }
    if (($first.Calls | ForEach-Object { $_.id } | Select-Object -Unique | Measure-Object).Count -lt $first.Calls.Count) {
        $problems += 'two calls carry the same id'
    }

    if ($first.Calls.Count -gt 0 -and $first.Status -eq 200) {
        Write-Host ''
        Write-Host '--- Request 2: the same conversation with the tool results ---'
        $assistant = @{ role = 'assistant'; content = $null; tool_calls = @() }
        if ($first.Text) { $assistant.content = $first.Text }
        $results = @()
        foreach ($call in $first.Calls) {
            $arguments = '{}'
            if ($call.arguments) { $arguments = $call.arguments }
            $assistant.tool_calls += @{ id = $call.id; type = 'function'; function = @{ name = $call.name; arguments = $arguments } }
            $results += @{ role = 'tool'; tool_call_id = $call.id; content = '{"time":"10:30"}' }
        }
        $second = Send-Round (@($question, $assistant) + $results)
        Write-Host ''
        if ($second.Status -ne 200) { $problems += "request 2 (with tool results) was answered HTTP $($second.Status)" }
        elseif (-not $second.Text) { $problems += 'request 2 was accepted but the answer has no text' }
        else { Write-Host ("Final text: {0}" -f $second.Text) }
    }
}
finally {
    $client.Dispose()
}

Write-Host ''
if ($problems.Count -eq 0) { Write-Host 'RESULT: OK' }
else { Write-Host ('RESULT: PROBLEM - ' + ($problems -join '; ')) }
