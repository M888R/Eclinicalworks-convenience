'Perhaps this would work as the whole script?'
' eclinicalWorks, username, password variables'
Dim eclinicalWorks, username, password, computerName

eclinicalWorks = "PATH TO ECLINICAL WORKS HERE" ' Insert its path here, or just write its filename and add to system PATH'
username = "USERNAME HERE" ' Put username Here '
password = "PASSWORD HERE" ' password Here '
computerName = "INSERT COMPUTER NAME HERE" ' Open "System" from the control panel and it'll be in there '

'process check func'
Function IsProcessRunning( strComputer, strProcess )
    Dim Process, strObject
    IsProcessRunning = False
    strObject   = "winmgmts://" & strComputer
    For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
    If UCase( Process.name ) = UCase( strProcess ) Then
        IsProcessRunning = True
        Exit Function
    End If
    Next
End Function

Dim tobjShell : Set tobjShell = WScript.CreateObject("WScript.Shell")
' Kill ECW '
tobjShell.Run ("TASKKILL" + eclinicalWorks, , True)
' Launch ECW '
tobjShell.Run ("START" + eclinicalWorks, , True)

while IsProcessRunning(computerName, eclinicalWorks) == false ' Check if ECW is running '
  tobjShell.Sleep 100
Wend

' Make it active '
tobjShell.AppActivate eclinicalWorks

' Begin password Enter Script '
tobjShell.SendKeys username
tobjShell.Sleep 300
tobjShell.SendKeys "  "
tobjShell.Sleep 300
tobjShell.SendKeys password
tobjShell.Sleep 300
tobjShell.SendKeys "{ENTER}"
