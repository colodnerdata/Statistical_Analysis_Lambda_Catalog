$excel = Get-Process -Name excel -ErrorAction SilentlyContinue
foreach ($p in $excel) {
  Write-Host "PID $($p.Id)  CPU $($p.CPU)  Started $($p.StartTime)  Title '$($p.MainWindowTitle)'  WS $($p.WorkingSet)"
}
$build = Get-Process -Id 15716 -ErrorAction SilentlyContinue
if ($build) {
  Write-Host "BUILD PID $($build.Id)  CPU $($build.CPU)  Threads $($build.Threads.Count)  Handles $($build.HandleCount)"
} else {
  Write-Host "BUILD HAS EXITED"
}