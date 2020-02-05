Const ForReading = 1

Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    ("servers.txt", ForReading)

Do Until objTextFile.AtEndOfStream
    strLine = objTextFile.ReadLine
    'Wscript.Echo strLine
arrayline = split(strLine,",")

'Call PingTest(arrayline(0))

'Call ConnectWMI(arrayline(0),arrayline(1),arrayline(2))

'Call ConnectSSH(arrayline(0),arrayline(1),arrayline(2))

Call AddAccount(arrayline(0))

Loop

objTextFile.Close

function AddAccount(strcomputer)
ON ERROR RESUME NEXT
Set objShell = CreateObject("Wscript.Shell")

Set objEnv = objShell.Environment("Process")

Set colAccounts = GetObject("WinNT://" & strComputer & ",computer")

Set objUser = colAccounts.Create("user", "LocalAccountName")

objUser.SetPassword "Pa55w0rd1"

Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
objPasswordExpirationFlag = ADS_UF_DONT_EXPIRE_PASSWD
objUser.Put "userFlags", objPasswordExpirationFlag

returtn = objUser.SetInfo

if return = 0 then
wscript.echo strcomputer & " account creation complete"
else
wscript.echo strcomputer & " account creation failed"
end if


Set Group = GetObject("WinNT://" & strComputer & "/Administrators,group")
return1 = Group.Add(objUser.ADspath)


if return1 = 0 then
wscript.echo strcomputer & " user added to group"
else
wscript.echo strcomputer & " user not added to group"
end if

End Function


Function ConnectWMI(strtarget,strusername,strpassword)

wscript.echo "WMI Connection Attempt to " & strtarget
ON ERROR RESUME NEXT
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objSWbemLocator.ConnectServer _
    (strtarget, "root\cimv2", strusername, strpassword)
objSWbemServices.Security_.ImpersonationLevel = 3

wscript.echo "WMI Connection Return Code = " & Err.Number
wscript.echo "WMI Connection Return Description = " & Err.Description

ON ERROR GOTO 0

End Function



Function PingTest(strtarget)
'#####
'ping a remote machine
'return the error code 0 = success 1 = fail

pingtest = WshShell.Run("ping.exe "& strtarget,1,1)

wscript.echo "Ping Return code = " & Err.Number

End Function


Function ConnectSSH(strtarget,strusername,strpassword)

strcommand = "Plink.exe " & strusername & "@" & strtarget & " -auto_store_key_in_cache -pw " & strpassword & " df -h"

wscript.echo "Attempting to SSH to server: " & strtarget

test = WshShell.Run(strcommand,1,1)
wscript.echo "Test Connection Return Code = " & test

Set oExec = WshShell.Exec(strcommand)

'We wait for the end of process
Do While oExec.Status = 0
     WScript.Sleep 100
Loop

wscript.echo oExec.StdOut.ReadAll()

End Function
