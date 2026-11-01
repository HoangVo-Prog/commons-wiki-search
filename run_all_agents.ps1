param([int]$Count=0)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$config = Get-Content -Raw -Encoding UTF8 (Join-Path $root 'config.json') | ConvertFrom-Json
$max = [int]$config.agent_count
if ($Count -eq 0) { $Count = $max }
if ($Count -lt 1 -or $Count -gt $max) { throw "Count must be 1..$max according to config.json" }
$one = Join-Path $root 'run_one_agent.ps1'
for ($i=1; $i -le $Count; $i++) {
    Start-Process -FilePath 'powershell.exe' -ArgumentList @(
        '-NoExit', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $one + '"'), '-AgentID', "$i"
    )
    Start-Sleep -Milliseconds 200
}
Write-Host "Started $Count Codex windows. Each reserves its configured ID before collecting; API requests are globally throttled."
