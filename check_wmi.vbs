'check wmi on a target
strComputer = "192.168.1.20"
strNamespace = "root\cimv2"
strUser = "LAB\test"
strPassword = "Pa55w0rd1"

Set objWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objWMIService = objwbemLocator.ConnectServer _
    (strComputer, strNamespace, strUser, strPassword)
objWMIService.Security_.authenticationLevel = 6

Set colItems = objWMIService.ExecQuery _
    ("Select * From Win32_OperatingSystem")
For Each objItem in ColItems
    Wscript.Echo strComputer & ": " & objItem.Caption
Next
