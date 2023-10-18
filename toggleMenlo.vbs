Set objShell = CreateObject("Wscript.shell")

disaToggle = "& { "& vbCrLf & _
"$path = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings'" & vbCrLf & _
"$menlo= 'https://pac.menlosecurity.com/af-5cf035477276/wpad.dat' "& vbCrLf & _
"try {" & vbCrLf & _
"$prox = Get-ItemPropertyValue -Path $path -Name AutoConfigURL" & vbCrLf & _
"If ($prox -eq $menlo) {Set-ItemProperty -Path $path -Name AutoConfigURL -Value ''}" & vbCrLf & _
"}" & vbCrLf & _
"catch { Set-ItemProperty -Path $path -Name AutoConfigURL -Value $menlo" & vbCrLf & _
"}" & vbCrLf & _
"finally { $Error.Clear() }" & vbCrLf & _
"}"

objShell.run("powershell -command """ & disaToggle & """ "),0