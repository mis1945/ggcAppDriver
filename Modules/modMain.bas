Attribute VB_Name = "modMain"
Public oApp As clsAppDriver

Private Sub Main()
10       Set oApp = New clsAppDriver
20       If oApp.LoadEnv("GRider") = False Then
25          MsgBox "Unable to Load GhostRider Application Driver!!!", vbCritical, "Warning"
30          Exit Sub
40       End If
   
50       If oApp.LogIn("") = False Then
55          MsgBox "Unable to Process User Login!!!", vbCritical, "Warning"
60          Exit Sub
70       End If
80       Set oApp = Nothing
End Sub

