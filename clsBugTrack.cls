VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBugTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
' Marlon A. Sayson
' Kalyptus Systems Enterprises
' Caoayan Kiling, San Carlos City
' 2420 Pangasinan
'
' Copyright 2006 and beyond
' All Rights Reserved
'
'     BugTrack - Bug and Error Tracking System
' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
' @  All  rights reserved. No part of this  software  @@    This Software is Owned by      @
' @  may be reproduced or transmitted in any form or  @@                                   @
' @  by   any   means,  electronic   or  mechanical,  @@    GUANZON MERCHANDISING CORP.    @
' @  including recording, or by information  storage  @@     Guanzon Bldg. Perez Blvd.     @
' @  and  retrieval  systems, without  prior written  @@           Dagupan City            @
' @  from the author.                                 @@  Tel No. 522-1085 ; 522-0863      @
' @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'
' ==========================================================================================
'  Kalyptus [11/24/2006] [11:29 am - 05:40 pm]
'        Analysis and merging of my previous Error Tracking Class with the new
'        idea i want to implement.
'  Kalyptus [11/28/2006] [08:39 am - 05:36 pm]
'        The previous method of tracking the stack of the error seems so complex.
'        Aside from the fact that a lot of changes has to be performed from the existing
'        application just to implement BugTracker.
'
'        1. I simplify the error caching through saving the cache to a temporary
'           file xxxTempxxx.Log
'        2. Instead of receiving the data needed from appdriver one by one. I created a
'           sub procedure to get them at the same time. See GetDriver(oDriver As Object)!
'        3. Added ImportErrorLog(). See Sub ImportErrorLog() for details
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

Option Explicit

Private l_sLogNoxxx As String
Private l_sClientID As String
Private l_sProductId As String
Private l_sComputerID As String
Private l_sUserId As String
Private l_oConnection As ADODB.Connection

'*
'* Sub GetDriver(oDriver As Object)
'*    - Gets the necessary data from AppDriver
'* Parameter:
'*    oDriver => AppDriver object
'* Returns:
'*    Nothing
'* Notes:
'*    Nothing
'* @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Sub GetDriver(ByVal oDriver As Object)
Attribute GetDriver.VB_Description = "Gets the necessary data from the AppDriver"
   Set l_oConnection = oDriver.Connection
   l_sLogNoxxx = oDriver.LogNumber                  'Should be present in our AppDriver
   l_sClientID = oDriver.ClientId
   l_sProductId = oDriver.ProductID
   l_sComputerID = oDriver.ComputerID
   l_sUserId = oDriver.UserID
End Sub

'*
'* Sub LogError(nErrNo As Integer,
'*              sErrDesc As String,
'*              sModule As String,
'*              sMethod As String,
'*              nErrLine As Integer)
'*    - Logs the error information/cache to the xxxTempxxx.log
'* Parameter:
'*    nErrNo   = Error Number Generated by the system
'*    sErrDesc = Error Description
'*    sModule  = Module/FileName that causes the error
'*    sMethod  = Sub/Function that causes the error
'*    nErrLine = Line in the function that causes the error
'* Returns:
'*    Nothing
'* Notes:
'*    Use this procedure to replace the previous ShowError procedure
'* @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Sub LogError(nErrNo As Integer, _
             sErrDesc As String, _
             sModule As String, _
             sMethod As String, _
             nErrLine As Integer)
Attribute LogError.VB_Description = "Logs the error information/cache to the Temporary Error Log"

   Dim intLogFile As Integer
   
   'On Error Goto errProc
      
   '* Obtain a free file handle.
   intLogFile = FreeFile
   
   '* Open the error log text file in Append mode.
   Open Environ("windir") & "\" & "xxxTempxxx.log" For Append As #intLogFile
   
   '* Write the entry from the temporary log file
   Write #intLogFile, l_sLogNoxxx, _
                      l_sClientID, _
                      l_sProductId, _
                      l_sComputerID, _
                      l_sUserId, _
                      Now, _
                      nErrNo, _
                      sErrDesc, _
                      sModule, _
                      sMethod, _
                      nErrLine
   Close #intLogFile
endProc:
   Exit Sub
errProc:
   Resume Next
End Sub

'*
'* Sub ShowError()
'*    - Display the error to the user
'* Parameter:
'*    Nothing
'* Returns:
'*    Nothing
'* Notes:
'*    This should be called only once
'* @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Sub ShowError()
Attribute ShowError.VB_Description = "Displays the Error to the user and at the same time copy the error cache to the error log."
   Dim lsLogNoxxx As String
   Dim lsClientID As String
   Dim lsProductId As String
   Dim lsComputerID As String
   Dim lsUserId As String
   Dim ldDateErr As String
   Dim lnErrNo As Integer
   Dim lsErrDesc As String
   Dim lsModule As String
   Dim lsMethod As String
   Dim lnErrLine As Integer

   Dim lsMessage As String
   Dim strCaption As String
   
   Dim ofrm As New frmShowError
   
   'On Error Goto errProc
   
   'Open and close the temporarry file in append mode to make sure that it exists
   Open Environ("windir") & "\" & "xxxTempxxx.log" For Append As #1
   Close #1

   '* Open the necessary files
   Open Environ("windir") & "\" & "xxxTempxxx.log" For Input As #1
   Open Environ("windir") & "\" & "GGCISys.log" For Append As #2

   lsMessage = ""
   
   Do Until EOF(1)
      '* Read the contents of the temporary error log file
      Input #1, lsLogNoxxx, _
                lsClientID, _
                lsProductId, _
                lsComputerID, _
                lsUserId, _
                ldDateErr, _
                lnErrNo, _
                lsErrDesc, _
                lsModule, _
                lsMethod, _
                lnErrLine
      '* Write the data to the main error log file
      Write #2, lsLogNoxxx, _
                lsClientID, _
                lsProductId, _
                lsComputerID, _
                lsUserId, _
                ldDateErr, _
                lnErrNo, _
                lsErrDesc, _
                lsModule, _
                lsMethod, _
                lnErrLine
     
     If Len(lsMessage) = 0 Then
        ofrm.ErrorDesc = lsErrDesc
        lsMessage = "Error Cache: " & lsErrDesc & vbCrLf & vbCrLf
     End If
     
     lsMessage = lsMessage & _
                "Module     : " & lsModule & vbCrLf & _
                "Procedure  : " & lsMethod & vbCrLf & _
                "Line Number: " & lnErrLine & vbCrLf & vbCrLf
   Loop

   '* Display the message if there is an error
   If Len(lsMessage) > 0 Then
'      lsMessage = lsMessage & vbCrLf & _
                  "Please notify the author(s) of this application for assistance." & _
                  "  See Technical Support menu for details on how to reach us!" & vbCrLf & vbCrLf & _
                  "Be sure to provide us with the information shown in this dialog box " & _
                  "as well as an explanation of what you were doing when this error occurred."
      strCaption = App.EXEName & " " & App.Major & _
                "." & App.Minor & " - Unexpected Error!"
      
      ofrm.ErrorNo = lnErrNo
      ofrm.ErrorInfo = lsMessage
      ofrm.Caption = strCaption
      ofrm.Show 1
      ' MsgBox lsMessage, vbCritical, strCaption
   End If
   
   '* Close the files
   Close #1
   Close #2
   
   '* Delete the temporary log file
   Kill Environ("windir") & "\" & "xxxTempxxx.log"
   
endProc:
   Exit Sub
errProc:
   Resume Next
End Sub

'*
'* Sub ImportErrorLog()
'*    - Extracts the data inside the error log file, kalyptus.log
'* Parameter:
'*    Nothing
'* Returns:
'*    Nothing
'* Notes:
'*    - Deletes the kalyptus.log after the extraction process
'* @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Sub ImportErrorLog()
Attribute ImportErrorLog.VB_Description = "Extracts the data inside the error log file"
   Dim lsOldLogNo As String
   Dim lsTranCode As String
   Dim lsTranLine As String
   Dim lsLogNoxxx As String
   Dim lsClientID As String
   Dim lsProductId As String
   Dim lsComputerID As String
   Dim lsUserId As String
   Dim ldDateErr As String
   Dim lnErrNo As Integer
   Dim lsErrDesc As String
   Dim lsModule As String
   Dim lsMethod As String
   Dim lnErrLine As Integer
   
   Dim lbHasErr As Boolean
   Dim loRec As ADODB.Recordset
   Dim lsSQL As String
   
   'On Error Goto errProc
   
   '* Open and close the log file in append mode to make sure that it exists
   Open Environ("windir") & "\" & "xxxTempxxx.log" For Append As #1
   Close #1

   Open Environ("windir") & "\" & "GGCISys.log" For Input As #1
   
   Do Until EOF(1)
      '* Read the contents of the temporary error log file
      Input #1, lsLogNoxxx, _
                lsClientID, _
                lsProductId, _
                lsComputerID, _
                lsUserId, _
                ldDateErr, _
                lnErrNo, _
                lsErrDesc, _
                lsModule, _
                lsMethod, _
                lnErrLine
      
      If lsOldLogNo <> lsLogNoxxx Then
         '* Get the last transaction Code
         lsSQL = "SELECT sTranCode FROM xxxAETMaster" _
              & " ORDER BY sTranCode DESC" _
              & " LIMIT 1"
         Set loRec = l_oConnection.Execute(lsSQL)
         lsTranCode = Format(loRec("sTrancode") + 1, "0000000000")
         loRec.Close
         
         '* Save the logged error to the xxxAETMaster table
         lsSQL = "INSERT INTO xxxAETMaster" _
                    & "( sTranCode" _
                    & ", sClntNmbr" _
                    & ", sProjNmbr" _
                    & ", sCompCode" _
                    & ", sUserCode" _
                    & ", dTimeMeet" _
                    & ", nErrNmbrx" _
                    & ", sErrDescx" _
                    & ", cStatCode" _
                    & ", sModified" _
             & ") VALUES" _
                    & " (" & strParm(lsTranCode) _
                    & ", " & strParm(lsClientID) _
                    & ", " & strParm(lsProductId) _
                    & ", " & strParm(lsComputerID) _
                    & ", " & strParm(lsUserId) _
                    & ", " & dateParm(ldDateErr) _
                    & ", " & lnErrNo _
                    & ", " & strParm(lsErrDesc) _
                    & ", " & strParm("0") _
                    & ", " & strParm(l_sUserId) _
             & ")"
         l_oConnection.Execute lsSQL
         
         lsTranLine = "00"
         lsOldLogNo = lsLogNoxxx
      End If
      
      lsTranLine = Format(lsTranLine + 1, "00")
      
      'Save it to the logged error to the xxxAETDetail table
      lsSQL = "INSERT INTO xxxAETMaster" _
                 & "( sTranCode" _
                 & ", sTranLine" _
                 & ", sProjNmbr" _
                 & ", sModulexx" _
                 & ", sProcName" _
             & ") VALUES" _
                 & " (" & strParm(lsTranCode) _
                 & ", " & strParm(lsTranLine) _
                 & ", " & strParm(lsModule) _
                 & ", " & strParm(lsMethod) _
                 & ", " & lnErrLine _
             & ")"
      l_oConnection.Execute lsSQL
      
      lbHasErr = True
   Loop
   
   '* Close the file
   Close #1
   
   '* Delete the error log file
   Kill Environ("windir") & "\" & "GGCISys.log"
   
   '* Display the success message
   If lbHasErr Then
      MsgBox "Error log exported...! Thank you!", vbInformation, App.EXEName & " " & App.Major & "." & App.Minor
   End If

endProc:
   Exit Sub
errProc:
   Resume Next
End Sub


'* Function strParm(ByVal sStrParm As String) As String
'* Function dateParm(ByVal dDateParm As Date) As String
'*
'* Notes:
'*    - They are self-explanatory. Read the code.
'* @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
Private Function strParm(ByVal sStrParm As String) As String
    strParm = "'" & Replace(sStrParm, "'", "''") & "'"
End Function

Private Function dateParm(ByVal dDateParm As Date) As String
    dateParm = "'" & Format(dDateParm, "YYYY-MM-DD HH:MM:SS") & "'"
End Function
