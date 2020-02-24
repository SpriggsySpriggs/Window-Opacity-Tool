$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'
(Get-Process | Where-Object {$_.MainWindowTitle -ne ""} | Select-Object MainWindowTitle | ft -auto -HideTableHeaders | Out-String).TrimStart() > openWindows.txt
