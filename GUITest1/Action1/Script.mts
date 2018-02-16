Set QtpApp = CreateObject("QuickTest.Application")
QC_Status = QtpApp.TDConnection.IsConnected
If QC_Status = "True" Then
Project = QtpApp.TDConnection.Project
If Project="Excel_Performance" Then
Else
QtpApp.TDConnection.Disconnect
wait(2)
QtpApp.TDConnection.Connect "http://10.11.17.79:8080/qcbin","DEFAULT","bison","admin","",False
End If
Else
QtpApp.TDConnection.Disconnect
wait(2)
QtpApp.TDConnection.Connect "http://10.11.17.79:8080/qcbin","DEFAULT","eew","admin","",False
End If
