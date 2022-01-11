Param(
        [parameter()]$ProcessName
     )
if($ProcessName -match '\.exe')
  {$ProcessName = $ProcessName -replace '\.exe'}
(Get-Process -Name $ProcessName).Kill()
