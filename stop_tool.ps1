$exe = Join-Path $PSScriptRoot 'bg_tool_server.exe'
$runtime = Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$listeners = netstat -ano -p tcp | Select-String ':8765\s+.*LISTENING\s+\d+$'
foreach ($line in $listeners) {
    $ownerPid = [int](($line.ToString().Trim() -split '\s+')[-1])
    $owner = Get-Process -Id $ownerPid -ErrorAction SilentlyContinue
    if ($owner -and (($owner.Path -eq $exe) -or ($owner.Path -eq $runtime))) {
        Stop-Process -Id $ownerPid -Force
    }
}
Write-Host 'Project management service stopped.'
