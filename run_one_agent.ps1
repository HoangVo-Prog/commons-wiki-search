param(
    [Parameter(Mandatory=$true)]
    [int]$AgentID,
    [switch]$Resume
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$config = Get-Content -Raw -Encoding UTF8 (Join-Path $root 'config.json') | ConvertFrom-Json
if ($AgentID -lt 1 -or $AgentID -gt [int]$config.agent_count) { throw "AgentID must be 1..$($config.agent_count)" }
$id = '{0:D2}' -f $AgentID
$worker = Join-Path (Join-Path $root $config.paths.worker_root) ($config.paths.worker_template.Replace('{agent_id}', $id))
if (-not (Test-Path (Join-Path $worker $config.paths.agent_identity))) { throw "Worker $id not initialized" }
if (-not (Get-Command codex -ErrorAction SilentlyContinue)) { throw 'Codex CLI not found on PATH.' }
$template = Get-Content -Raw -Encoding UTF8 (Join-Path $root $config.paths.task_instructions)
$prompt = $template.Replace('{AGENT_ID}', $id).Trim()
Push-Location $root
try {
    if ($Resume) {
        $claims = Get-Content -Raw -Encoding UTF8 (Join-Path $root $config.paths.claims) | ConvertFrom-Json
        if (-not $claims.PSObject.Properties[$id]) { throw "Agent $id has no existing claim; remove -Resume and claim first." }
        Write-Warning "Resume mode: make sure the previous Codex session for agent $id has stopped."
    }
    else {
        & python claim_agent.py reserve --agent-id $id
        if ($LASTEXITCODE -ne 0) { throw "Unable to reserve agent $id. Already claimed or registry error." }
    }
    Write-Host "Launching API-only Codex agent $id from $root"
    & codex $prompt
}
finally { Pop-Location }
