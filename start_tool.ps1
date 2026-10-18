$ErrorActionPreference = 'Stop'
try {
    $exe = Join-Path $PSScriptRoot 'bg_tool_server.exe'
    $source = Join-Path $PSScriptRoot 'server.py'
    $runtime = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
    $expectedVersion = '2026.10.18-performance'
    $useSource = (Test-Path -LiteralPath $source) -and (Test-Path -LiteralPath $runtime) -and
        ((-not (Test-Path -LiteralPath $exe)) -or ((Get-Item -LiteralPath $source).LastWriteTimeUtc -gt (Get-Item -LiteralPath $exe).LastWriteTimeUtc))
    $ready = $false
    try {
        $probe = Invoke-RestMethod 'http://127.0.0.1:8765/api/auth/status' -TimeoutSec 2
        $ready = $probe.version -eq $expectedVersion -and $probe.app_dir -eq $PSScriptRoot
    } catch {}
    if (-not $ready) {
        $listeners = netstat -ano -p tcp | Select-String ':8765\s+.*LISTENING\s+\d+$'
        foreach ($line in $listeners) {
            $ownerPid = [int](($line.ToString().Trim() -split '\s+')[-1])
            $owner = Get-Process -Id $ownerPid -ErrorAction SilentlyContinue
            if ($owner -and (($owner.Path -eq $runtime) -or ($owner.Path -eq $exe))) {
                Stop-Process -Id $ownerPid -Force
            } else { throw 'Port 8765 is in use. Close the other application first.' }
        }
        if ($useSource) {
            Start-Process -FilePath $runtime -ArgumentList '-u', ('"' + $source + '"') -WorkingDirectory $PSScriptRoot -WindowStyle Hidden
        } elseif (Test-Path -LiteralPath $exe) {
            Start-Process -FilePath $exe -WorkingDirectory $PSScriptRoot -WindowStyle Hidden
        } elseif (Test-Path -LiteralPath $runtime) {
            Start-Process -FilePath $runtime -ArgumentList '-u', ('"' + $source + '"') -WorkingDirectory $PSScriptRoot -WindowStyle Hidden
        } else { throw 'Program missing. Extract the complete portable package again.' }
        for ($i=0; $i -lt 40; $i++) {
            Start-Sleep -Milliseconds 500
            try {
                $probe = Invoke-RestMethod 'http://127.0.0.1:8765/api/auth/status' -TimeoutSec 2
                $ready = $probe.version -eq $expectedVersion -and $probe.app_dir -eq $PSScriptRoot
            } catch {}
            if ($ready) { break }
        }
    }
    if (-not $ready) { throw 'Startup failed. Check tool_run.log and port 8765.' }
    Start-Process 'http://localhost:8765'
} catch {
    Write-Host $_.Exception.Message
    Read-Host 'Press Enter to close'
    exit 1
}
