param([int]$BuildPid = 0)

$excel = Get-Process -Name excel -ErrorAction SilentlyContinue
foreach ($p in $excel) {
  Write-Host "PID $($p.Id)  CPU $($p.CPU)  Started $($p.StartTime)  Title '$($p.MainWindowTitle)'  WS $($p.WorkingSet)"
}

if ($BuildPid) {
  $build = Get-Process -Id $BuildPid -ErrorAction SilentlyContinue
  if ($build) {
    Write-Host "BUILD PID $($build.Id)  CPU $($build.CPU)  Threads $($build.Threads.Count)  Handles $($build.HandleCount)"
  } else {
    Write-Host "BUILD PID $BuildPid HAS EXITED"
  }
} else {
  Write-Host "No BuildPid specified; pass a PID to inspect the build process."
}