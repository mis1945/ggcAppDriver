VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAppDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Rex S. Adversalo
' XerSys Computing
' Canaoalan, Binmaley, Pangasinan
'
' Copyright 2003-2004 and beyond
' All Rights Reserved
'
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  XerSys [ 05/20/2004 10:28 am ]
'     Start revising this object. i adopt the idea of the system environment
'  variables, definition and loading from the TBoxPlus v97.01 of the late Mr. Sison and
'  combine it with the idea of my dear friend Marlon Sayson (with his consent).
' ------------------------------------------------------------------------------------------
'  XerSys [ 12/10/2006 05:03 pm ]
'     Create a second version of this object.
'        -  Connection is MySQL
'        -  Refer to GCC SEGs Notes for detailed changes of this object
'  XerSys [ 12/13/2006 08:53 am ]
'     Continue revising this object
'  XerSys [ 01/27/2007 09:44 am ]
'     Add this function
'        - getConfiguration(ConfigCd as string) as Variant
'           - retrieves special configuration value set by each product
' ------------------------------------------------------------------------------------------
'  Kalyptus [ 12/20/2006 11:20 am ]
'     Inserts BugTracking
'  Kalyptus [ 12/27/2013 10:43 am ]
'     Implements Day-To-Day opening
'     See Config property
'         LoadEnv function
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Const xeDriverID As String = "GRider"
Private Const xeSignature As String = "08220326"
Private Const pxeMODULENAME = "clsAppDriver"
Const pxeDay2Day As String = "DT"

' system environment variables
Private p_sClientId As String
Private p_sClientNm As String
Private p_sAddressx As String
Private p_sTownName As String
Private p_sZippCode As String
Private p_sProvName As String
Private p_sTelNoxxx As String
Private p_sFaxNoxxx As String
Private p_sApproved As String
Private p_sSysAdmin As String
Private p_sProdctID As String
Private p_sProdctNm As String
Private p_sNetWarex As String
Private p_sMachinex As String
Private p_sApplPath As String
Private p_sReptPath As String
Private p_sImgePath As String
Private p_dSysDatex As Date
Private p_nNetError As Long
Private p_sBranchCd As String
Private p_dCapturex As Date
Private p_dLicencex As Date
Private p_sComptrID As String
Private p_sMenuName As String

' system-user variables
Private p_sUserIDxx As String
Private p_sLogNamex As String
Private p_sUserName As String
Private p_sEmployNo As String
Private p_sLogNoxxx As String
Private p_nUserLevl As xeUserRights
Private p_nHexCrypt As Integer

Private p_bAllowUpd As Boolean
Private p_bWithTran As Boolean

' system user variables
Private p_ocnMain As Connection
Private p_oMDIForm As MDIForm

Private p_oMod As New clsMainModules

' skin variables
Private p_sSkinCode As String
Private p_sFormPict As String
Private p_sQSImage As String
Private p_sFldName As String
Private p_anColor() As Long

' branch variables
Private p_sBranchNm As String
Private p_cMainOffc As Byte
Private p_cWareHous As Byte

Private p_sCompName As String
Private p_sProcName As String

Private pbErrorLog As Boolean
Private pnTransact As Integer
Private psModlName As String
Private psSQL As String

Private pbChkErrCt As Boolean
Private p_oRSConfg As Recordset

Private p_bBypass As Boolean

Property Get Connection() As Connection
   Set Connection = p_ocnMain
End Property

Property Get ClientId() As String
   ClientId = p_sClientId
End Property

Property Get ClientName() As String
   ClientName = p_sClientNm
End Property

Property Get Address() As String
   Address = p_sAddressx
End Property

Property Get TownCity() As String
   TownCity = p_sTownName
End Property

Property Get ZippCode() As String
   ZippCode = p_sZippCode
End Property

Property Get Province() As String
   Province = p_sProvName
End Property

Property Get TelNo() As String
   TelNo = p_sTelNoxxx
End Property

Property Get FaxNo() As String
   FaxNo = p_sFaxNoxxx
End Property

Property Get Approved() As String
   Approved = p_sApproved
End Property

Property Get SysAdmin() As String
   SysAdmin = p_sSysAdmin
End Property

Property Get ProductID() As String
   ProductID = p_sProdctID
End Property

Property Get ProductName() As String
   ProductName = p_sProdctNm
End Property

Property Get NetWarex() As String
   NetWarex = p_sNetWarex
End Property

Property Get Machinex() As String
   Machinex = p_sMachinex
End Property

Property Get AppPath() As String
   AppPath = p_sApplPath
End Property

Property Get SysDate() As Date
   SysDate = p_dSysDatex
End Property

Property Get NetError() As Long
   NetError = p_nNetError
End Property

Property Get BranchCode() As String
   BranchCode = p_sBranchCd
End Property

Property Get BranchName() As String
   BranchName = p_sBranchNm
End Property

Property Get IsMainOffice() As Boolean
   IsMainOffice = p_cMainOffc = 1
End Property

Property Get IsWarehouse() As Boolean
   IsWarehouse = p_cWareHous = 1
End Property

Property Get UserID() As String
   UserID = p_sUserIDxx
End Property

Property Get LogName() As String
   LogName = p_sLogNamex
End Property

Property Get LogNumber() As String
   LogNumber = p_sLogNoxxx
End Property

Property Get UserName() As String
   UserName = p_sUserName
End Property

Property Get EmployeeNo() As String
   EmployeeNo = p_sEmployNo
End Property

Property Get UserLevel() As Integer
   UserLevel = p_nUserLevl
End Property

Property Get ComputerName() As String
   ComputerName = p_sCompName
End Property

Property Get ComputerID() As String
   ComputerID = p_sComptrID
End Property

Property Get MenuName() As String
   MenuName = p_sMenuName
End Property

Property Let MenuName(Value As String)
   p_sMenuName = Value
End Property

Property Set MDIMain(Value As Object)
   Set p_oMDIForm = Value
   SetMDIMain
End Property

Property Get MDIMain() As Object
   Set MDIMain = p_oMDIForm
End Property

Property Get CaptureDate() As Date
   CaptureDate = p_dCapturex
End Property

Property Let ChkErrCnt(lbValue As Boolean)
   pbChkErrCt = lbValue
End Property

Property Let ByPassACID(lbValue As Boolean)
   p_bBypass = lbValue
End Property

'kalyptus - 2013.12.27 10:32am
'Added this property which will be used in accessing the value of
'fields in Branch_Others...
Property Get Config(fsValue As String) As Variant
   Dim lors As Recordset
   Dim lsSQL As String
   Select Case LCase(fsValue)
   Case "ssalesinv", "sreceiptx"
      If LCase(fsValue) = "ssalesinv" Then
         lsSQL = "SELECT sSalesInv" & _
                " FROM Branch_Others" & _
                " WHERE sBranchCD = " & p_oMod.strParm(p_sBranchCd)
         Set lors = p_ocnMain.Execute(lsSQL, , adCmdText)
         Config = lors("sSalesInv")
      Else
         lsSQL = "SELECT sReceiptx" & _
                " FROM Branch_Others" & _
                " WHERE sBranchCD = " & p_oMod.strParm(p_sBranchCd)
         Set lors = p_ocnMain.Execute(lsSQL, , adCmdText)
         
         Config = lors("sReceiptx")
      End If
      
      lors.Close
      Set lors = Nothing
   Case Else
      Config = p_oRSConfg(fsValue)
   End Select
End Property

'Function LogIn(sProduct As String) As Boolean
'   ' sProduct will be the controlling variable. If sProduct contains a valid value
'   '  then it will not allow the user to select an application to be executed; if it
'   '  contains an empty string then selection is available and the selected application
'   '  will be selected.
'   Dim loProduct As ADODB.Recordset
'   Dim loSysUser As ADODB.Recordset
'   Dim loForm As frmLogin
'   Dim loCrypt As clsCrypto
'   Dim lsUserName As String
'   Dim lsPassword As String
'   Dim lsProcName As String
'   Dim lsSQL As String
'   Dim lnCtr As Long
'   Dim lbLogIn As Boolean
'
'   lsProcName = "LogIn"
'   On Error GoTo errProc
'
'   Set loProduct = New Recordset
'   Set loSysUser = New Recordset
'   Set loForm = New frmLogin
'   Set loCrypt = New clsCrypto
'   loCrypt.Signature = xeSignature
'
'   With p_oMod
'      If sProduct = Empty Then
'         lsSQL = "SELECT" & _
'                     "  a.sProdctID" & _
'                     ", a.sProdctNm" & _
'                     ", a.sApplName" & _
'                     ", b.sApplPath" & _
'                  " FROM xxxSysObject a" & _
'                     ", xxxSysApplication b" & _
'                  " WHERE a.sProdctID = b.sProdctID" & _
'                     " AND b.sClientID = " & .strParm(p_sClientId) & _
'                     " AND a.sProdctID <> " & .strParm(xeDriverID)
'      Else
'         lsSQL = "SELECT" & _
'                     "  a.sProdctID" & _
'                     ", a.sProdctNm" & _
'                     ", a.sApplName" & _
'                     ", b.sApplPath" & _
'                  " FROM xxxSysObject a" & _
'                     ", xxxSysApplication b" & _
'                  " WHERE a.sProdctID = b.sProdctID" & _
'                     " AND b.sClientID = " & .strParm(p_sClientId) & _
'                     " AND a.sProdctID = " & .strParm(sProduct)
'      End If
'      loProduct.Open lsSQL, p_ocnMain, adOpenStatic, , adCmdText
'   End With
'
'   If loProduct.EOF Then
'      MsgBox "Invalid Application Registration Detected!"
'      GoTo endProc
'   End If
'
'   With loForm
'      loForm.lblCompany = p_sClientNm
'      loForm.lblAddress = Trim(p_sAddressx) & ", " & Trim(p_sTownName) & ", " & _
'      Trim(p_sProvName) & " " & Trim(p_sZippCode) & ", "
'      loForm.lblTelNo = "Tel No: " & p_sTelNoxxx
'      Do
'         .Combo1.AddItem (loProduct("sProdctNm"))
'         loProduct.MoveNext
'      Loop While Not loProduct.EOF
'      If sProduct <> Empty Then .Combo1.ListIndex = 0
'
'      loForm.txtDate.Text = Format(ServerDate, "YYYY/MM/DD")
'      loForm.txtDate.Visible = IsDate(p_oRSConfg("dOpenD2Dx"))
'      loForm.Label2.Visible = IsDate(p_oRSConfg("dOpenD2Dx"))
'   End With
'
'   lnCtr = 0
'
'   With loSysUser
'      Do
'         loForm.Show 1
'         If loForm.Cancel = True Then
'            ' user either clicked the cancel button or the close key
'            lnCtr = 2
'         Else
'            ' Encrypt the username and password
'            loCrypt.InBuffer = loForm.txtPassword
'            loCrypt.Encrypt
'            lsPassword = loCrypt.OutBuffer
'
'            loCrypt.InBuffer = LCase(loForm.txtUserName)
'            loCrypt.Encrypt
'            lsUserName = loCrypt.OutBuffer
'
'            ' search for the user
'            lsSQL = "SELECT *" & _
'                     " FROM xxxSysUser" & _
'                     " WHERE sLogNamex = " & p_oMod.strParm(Replace(lsUserName, "\", "\\")) & _
'                        " AND sPassword = " & p_oMod.strParm(Replace(lsPassword, "\", "\\"))
'
'            If .State = adStateOpen Then .Close
'            .Open lsSQL, p_ocnMain, , , adCmdText
'            If .EOF Then
'               MsgBox "Verify your log name and/or password", vbCritical, "Login Error"
'            Else
'               lbLogIn = True
'            End If
'         End If
'         lnCtr = lnCtr + 1
'      Loop Until (lbLogIn = True) Or (lnCtr = 3)
'
'      If lbLogIn = False Then GoTo endProc
'
'      loProduct.MoveFirst
'      loProduct.Find "sProdctNm = " & p_oMod.strParm(loForm.Combo1.Text), 0, adSearchForward, 0
'
'      If loProduct.EOF Then GoTo endProc
'
'      ' check first if user is allowed to use the application
'      If .Fields("cUserStat") = xeUserStateSuspended Then
'         MsgBox "User is currently suspended!!!" & vbCrLf & _
'                  "Application used is not allowed!!!", vbCritical, "Warning"
'         GoTo endProc
'      End If
'
'      If .Fields("cUserType") = xeUserTypeLocal Then
'         If .Fields("sProdctID") <> loProduct("sProdctID") Then
'            .Find "sProdctID = " & _
'                     p_oMod.strParm(loProduct("sProdctID")), 1, adSearchForward
'            If .EOF Then
'               MsgBox "User is not a member of this application!!!" & vbCrLf & _
'                        "Application used is not allowed!!!", vbCritical, "Warning"
'               GoTo endProc
'            End If
'         End If
'      End If
'
'      p_dSysDatex = CDate(loForm.txtDate)
'
'      If sProduct <> Empty Then
'         If LogUser(.Fields("sUserIDxx")) = False Then GoTo endProc
'         getSkin p_sSkinCode
'
'         'Kalyptus - 2013.12.26 10:46am
'         'Load Day-To-Day opening screen here
'         If IsDate(p_oRSConfg("dOpenD2Dx")) And LCase(HostName) = LCase(p_oRSConfg("sDBHostNm")) Then
'            If openDay2Day(.Fields("sUserIDxx")) = False Then GoTo endProc
'         End If
'
'         LogIn = True
'         GoTo endProc
'      Else
'         If IsDate(p_oRSConfg("dOpenD2Dx")) Then
'            If LCase(HostName) = LCase(p_oRSConfg("sDBHostNm")) Then
'               psSQL = "UPDATE xxxSysApplication SET" & _
'                           " dSysDatex = " & p_oMod.dateParm(p_dSysDatex) & _
'                        " WHERE sClientID = " & p_oMod.strParm(p_sClientId) & _
'                           " AND sProdctID = " & p_oMod.strParm(loProduct("sProdctID"))
'               p_ocnMain.Execute psSQL
'            End If
'         End If
'      End If
'
'      getSkin p_sSkinCode
'
'      'Kalyptus - 2013.12.26 10:46am
'      'Load Day-To-Day opening screen here
'      If IsDate(p_oRSConfg("dOpenD2Dx")) And LCase(HostName) = LCase(p_oRSConfg("sDBHostNm")) Then
'         If openDay2Day(.Fields("sUserIDxx")) = False Then GoTo endProc
'      End If
'
'      If Shell(loProduct("sApplPath") & "\" & _
'               loProduct("sApplName") & " " & _
'               loProduct("sProdctID") & " " & _
'               .Fields("sUserIDxx") & " " & _
'               .Fields("sUserName") & " " & _
'               .Fields("sPassword"), vbMaximizedFocus) = 0 Then GoTo endProc
'      LogIn = True
'   End With
'endProc:
'   Unload loForm
'   Set loCrypt = Nothing
'   Set loSysUser = Nothing
'   Set loForm = Nothing
'   Set loProduct = Nothing
'
'   Exit Function
'errProc:
'   ShowError lsProcName & "( " & sProduct & " )"
'End Function
Function LogIn(sProduct As String) As Boolean
   ' sProduct will be the controlling variable. If sProduct contains a valid value
   '  then it will not allow the user to select an application to be executed; if it
   '  contains an empty string then selection is available and the selected application
   '  will be selected.
   Dim loProduct As ADODB.Recordset
   Dim loSysUser As ADODB.Recordset
   Dim loForm As frmLogin
   Dim loCrypt As clsCrypto
   Dim lsUserName As String
   Dim lsPassword As String
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCtr As Long
   Dim lbLogIn As Boolean
   
   lsProcName = "LogIn"
   On Error GoTo errProc
   
   Set loProduct = New Recordset
   Set loSysUser = New Recordset
   Set loForm = New frmLogin
   'Set loCrypt = New clsCrypto
   'loCrypt.Signature = xeSignature

   With p_oMod
      If sProduct = Empty Then
         lsSQL = "SELECT" & _
                     "  a.sProdctID" & _
                     ", a.sProdctNm" & _
                     ", a.sApplName" & _
                     ", b.sApplPath" & _
                  " FROM xxxSysObject a" & _
                     ", xxxSysApplication b" & _
                  " WHERE a.sProdctID = b.sProdctID" & _
                     " AND b.sClientID = " & .strParm(p_sClientId) & _
                     " AND a.sProdctID <> " & .strParm(xeDriverID)
      Else
         lsSQL = "SELECT" & _
                     "  a.sProdctID" & _
                     ", a.sProdctNm" & _
                     ", a.sApplName" & _
                     ", b.sApplPath" & _
                  " FROM xxxSysObject a" & _
                     ", xxxSysApplication b" & _
                  " WHERE a.sProdctID = b.sProdctID" & _
                     " AND b.sClientID = " & .strParm(p_sClientId) & _
                     " AND a.sProdctID = " & .strParm(sProduct)
      End If
      loProduct.Open lsSQL, p_ocnMain, adOpenStatic, , adCmdText
   End With
               
   If loProduct.EOF Then
      MsgBox "Invalid Application Registration Detected!"
      GoTo endProc
   End If
   
   With loForm
      loForm.lblCompany = p_sClientNm
      loForm.lblAddress = Trim(p_sAddressx) & ", " & Trim(p_sTownName) & ", " & _
      Trim(p_sProvName) & " " & Trim(p_sZippCode) & ", "
      loForm.lblTelNo = "Tel No: " & p_sTelNoxxx
      Do
         .Combo1.AddItem (loProduct("sProdctNm"))
         loProduct.MoveNext
      Loop While Not loProduct.EOF
      If sProduct <> Empty Then .Combo1.ListIndex = 0
      
      loForm.txtDate.Text = Format(ServerDate, "YYYY/MM/DD")
      loForm.txtDate.Visible = IsDate(p_oRSConfg("dOpenD2Dx"))
      loForm.imgDate.Visible = IsDate(p_oRSConfg("dOpenD2Dx"))
   End With
   
   lnCtr = 0
   
   With loSysUser
      Do
         loForm.Show 1
         If loForm.Cancel = True Then
            ' user either clicked the cancel button or the close key
            lnCtr = 2
         Else
            ' Encrypt the username and password
            'loCrypt.InBuffer = loForm.txtPassword
            'loCrypt.Encrypt
            'lsPassword = loCrypt.OutBuffer
            lsPassword = Encrypt(loForm.txtPassword, xeSignature)
            
            'loCrypt.InBuffer = LCase(loForm.txtUserName)
            'loCrypt.Encrypt
            'lsUserName = loCrypt.OutBuffer
            lsUserName = Encrypt(LCase(loForm.txtUserName), xeSignature)
                        
            ' search for the user
            lsSQL = "SELECT *" & _
                     " FROM xxxSysUser" & _
                     " WHERE sLogNamex = " & p_oMod.strParm(Replace(lsUserName, "\", "\\")) & _
                        " AND sPassword = " & p_oMod.strParm(Replace(lsPassword, "\", "\\"))
            
            If .State = adStateOpen Then .Close
            .Open lsSQL, p_ocnMain, , , adCmdText
            If .EOF Then
               MsgBox "Verify your log name and/or password", vbCritical, "Login Error"
            Else
               lbLogIn = True
            End If
         End If
         lnCtr = lnCtr + 1
      Loop Until (lbLogIn = True) Or (lnCtr = 3)
      
      If lbLogIn = False Then GoTo endProc
      
      loProduct.MoveFirst
      loProduct.Find "sProdctNm = " & p_oMod.strParm(loForm.Combo1.Text), 0, adSearchForward, 0
      
      If loProduct.EOF Then GoTo endProc
   
      ' check first if user is allowed to use the application
      If .Fields("cUserStat") = xeUserStateSuspended Then
         MsgBox "User is currently suspended!!!" & vbCrLf & _
                  "Application used is not allowed!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      If .Fields("cUserType") = xeUserTypeLocal Then
         If .Fields("sProdctID") <> loProduct("sProdctID") Then
            .Find "sProdctID = " & _
                     p_oMod.strParm(loProduct("sProdctID")), 1, adSearchForward
            If .EOF Then
               MsgBox "User is not a member of this application!!!" & vbCrLf & _
                        "Application used is not allowed!!!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      End If
      
      p_dSysDatex = CDate(loForm.txtDate)
      
      If sProduct <> Empty Then
         If LogUser(.Fields("sUserIDxx")) = False Then GoTo endProc
         getSkin p_sSkinCode
         
         'Kalyptus - 2013.12.26 10:46am
         'Load Day-To-Day opening screen here
         
         If Left(LCase(HostName), 9) = Left(LCase(p_oRSConfg("sDBHostNm")), 9) _
         Or InStr(1, "M001»M0W1»M0W2»M0W3»M029»C001»C0CW»C0W1»C0W2", p_sBranchCd) > 1 Then
            If IsDate(p_oRSConfg("dOpenD2Dx")) Then
               If openDay2Day(.Fields("sUserIDxx")) = False Then GoTo endProc
            Else
               psSQL = "UPDATE xxxSysApplication SET" & _
                           " dSysDatex = " & p_oMod.dateParm(p_dSysDatex) & _
                        " WHERE sClientID = " & p_oMod.strParm(p_sClientId) & _
                           " AND sProdctID = " & p_oMod.strParm(loProduct("sProdctID"))
               p_ocnMain.Execute psSQL
            End If
         End If
         
         LogIn = True
         GoTo endProc
      Else
         If IsDate(p_oRSConfg("dOpenD2Dx")) Then
            If Left(LCase(HostName), 9) = Left(LCase(p_oRSConfg("sDBHostNm")), 9) _
            Or InStr(1, "M001»M0W1»M0W2»M0W3»M029»C001»C0CW»C0W1»C0W2", p_sBranchCd) > 1 Then
               psSQL = "UPDATE xxxSysApplication SET" & _
                           " dSysDatex = " & p_oMod.dateParm(p_dSysDatex) & _
                        " WHERE sClientID = " & p_oMod.strParm(p_sClientId) & _
                           " AND sProdctID = " & p_oMod.strParm(loProduct("sProdctID"))
               p_ocnMain.Execute psSQL
            End If
         End If
      End If
      
      getSkin p_sSkinCode
      
      'Kalyptus - 2013.12.26 10:46am
      'Load Day-To-Day opening screen here
      If IsDate(p_oRSConfg("dOpenD2Dx")) Then
         If Left(LCase(HostName), 9) = Left(LCase(p_oRSConfg("sDBHostNm")), 9) _
         Or InStr(1, "M001»M0W1»M0W2»M0W3»M029»C001»C0CW»C0W1»C0W2", p_sBranchCd) > 1 Then
            If openDay2Day(.Fields("sUserIDxx")) = False Then GoTo endProc
         End If
      End If
      
      If Shell(loProduct("sApplPath") & "\" & _
               loProduct("sApplName") & " " & _
               loProduct("sProdctID") & " " & _
               .Fields("sUserIDxx") & " " & _
               .Fields("sUserName") & " " & _
               .Fields("sPassword"), vbMaximizedFocus) = 0 Then GoTo endProc
      LogIn = True
   End With
endProc:
   Unload loForm
   Set loCrypt = Nothing
   Set loSysUser = Nothing
   Set loForm = Nothing
   Set loProduct = Nothing
   
   Exit Function
errProc:
   ShowError lsProcName & "( " & sProduct & " )"
End Function

Function LoadEnv(ByVal sProduct As String, _
                  Optional sUserID As Variant, _
                  Optional sUserName As Variant, _
                  Optional sPassword As Variant) As Boolean
   Dim lors As New Recordset
   'Dim loCrypt As clsCrypto
   Dim loINI As clsINIFile
   Dim lsServerNm As String, lsDatabase As String, lsPassword As String
   Dim lsUserName As String, lsClientCd As String, lsBranchCD As String
   Dim lsProcName As String
   Dim lsDriverNm As String
   Dim lsPortNmbr As String
   Dim lbUpdateDt As Boolean
   Dim lsMyStation As String
      
   lsProcName = "LoadEnv"
   On Error GoTo errProc
   
   LoadEnv = False
   lsMyStation = "GMC_SEG01"
   
   Set loINI = New clsINIFile
   Set lors = New Recordset
   'Set loCrypt = New clsCrypto
   'loCrypt.Signature = xeSignature
   
   ' prepare the ini file that will be used in comparing the records...
   loINI.FileName = Environ("windir") & "\GhostRiderXP.ini"
   
   If Not loINI.isFileExist Then
      Err.Raise 53
      Err.Clear
   End If
   
   'kalyptus - 2018.04.28 03:45pm
   'Load a new config variable that will determine the type of encryption to use...
   p_nHexCrypt = Val(loINI.getTextValue(sProduct, "CryptType"))
   
   ' only gRider Application has Entry to GhostRider
   lsServerNm = loINI.getTextValue(sProduct, "ServerName")
   lsDatabase = loINI.getTextValue(sProduct, "Database")
   lsClientCd = loINI.getTextValue(sProduct, "ClientID")
   lsDriverNm = loINI.getTextValue(sProduct, "DriverName")
   lsPortNmbr = loINI.getTextValue(sProduct, "Port")
   
   If Len(lsDriverNm) < 5 Then lsDriverNm = "MySQL ODBC 3.51 Driver"
   If Trim(lsPortNmbr) = "" Then lsPortNmbr = "3306"

'105      lsBranchCd = loINI.getTextValue(sProduct, "BranchCd")
   
   If sProduct = xeDriverID Then
      lsUserName = loINI.getTextValue(sProduct, "UserName")
      lsPassword = loINI.getTextValue(sProduct, "Password")
   Else
      lsUserName = loINI.getTextValue(sProduct, "UserName")
      lsPassword = loINI.getTextValue(sProduct, "Password")
   End If
   
   If lsServerNm = Empty Or _
         lsDatabase = Empty Or _
         lsUserName = Empty Or _
         lsDriverNm = Empty Or _
         lsPassword = Empty Then
      ShowError "INI file is corrupt"
   End If

   'kalyptus - 2018.04.28 03:43pm
   'Use the new decryption method
   '  password and username of a user will be used as connection properties
   lsUserName = Decrypt(lsUserName, xeSignature)
   'loCrypt.InBuffer = lsUserName
   'loCrypt.Encrypt
   'lsUserName = loCrypt.OutBuffer

   lsPassword = Decrypt(lsPassword, xeSignature)
   'loCrypt.InBuffer = lsPassword
   'loCrypt.Encrypt
   'lsPassword = loCrypt.OutBuffer
   
   psSQL = "DRIVER={" & lsDriverNm & "};" & _
               "SERVER=" & lsServerNm & ";" & _
               "PORT=" & lsPortNmbr & ";" & _
               "DATABASE=" & lsDatabase & ";" & _
               "USER=" & lsUserName & ";" & _
               "PASSWORD=" & lsPassword & ";" & _
               "OPTION=16395;"
   
   ' Set the connection and prepare the record to be extracted...
   p_ocnMain.CursorLocation = adUseClient
'   p_ocnMain.CursorLocation = adUseServer
   p_ocnMain.Open psSQL
   
   With p_oMod
      ' get the computer name first
      p_sCompName = .getComputerName()
      p_sComptrID = loINI.getTextValue(sProduct, "ComputerID")
      
      psSQL = "SELECT" & _
                  "  a.sClientID" & _
                  ", a.sClientNm" & _
                  ", a.sAddressx" & _
                  ", a.sTownName" & _
                  ", a.sZippCode" & _
                  ", a.sProvName" & _
                  ", a.sTelNoxxx" & _
                  ", a.sFaxNoxxx" & _
                  ", a.sApproved" & _
                  ", a.sBranchCd" & _
                  ", b.sProdctID" & _
                  ", b.sProdctNm" & _
                  ", b.sApplName" & _
                  ", c.sApplPath" & _
                  ", c.sReptPath" & _
                  ", c.sImgePath" & _
                  ", c.sSysAdmin" & _
                  ", c.sNetWarex" & _
                  ", c.sMachinex" & _
                  ", c.dSysDatex" & _
                  ", c.dLicencex" & _
                  ", c.nNetError" & _
                  ", c.sSkinCode" & _
                  ", d.sComptrNm"
      psSQL = psSQL & _
                  ", e.sBranchNm" & _
                  ", e.cWareHous" & _
                  ", e.cMainOffc" & _
               " FROM xxxSysClient a" & _
                  ", xxxSysObject b" & _
                  ", xxxSysApplication c" & _
                        " LEFT JOIN xxxSysWorkStation d" & _
                           " ON c.sClientID = d.sClientID" & _
                              " AND d.sComptrNm = " & .strParm(p_sCompName) & _
                  ", Branch e" & _
               " WHERE c.sClientID = a.sClientID" & _
                  " AND c.sProdctID = b.sProdctID" & _
                  " AND a.sBranchCd = e.sBranchCd" & _
                  " AND a.sClientID = " & .strParm(lsClientCd) & _
                  " AND b.sProdctID = " & .strParm(sProduct)
   End With
   
   Debug.Print psSQL
   
   lors.Open psSQL, p_ocnMain, , , adCmdText
   Set lors.ActiveConnection = Nothing
   If lors.EOF Then
      ShowError "Application is Not Registered!!!"
   End If
   
   If IsNull(lors("sComptrNm")) And p_sCompName <> lsMyStation Then
      MsgBox "Computer is Not Registered to Use The Selected Sytem"
      GoTo endProc
   End If
   Debug.Print psSQL
   
   p_nNetError = lors("nNetError").Value
   
   ' check error count
   If pbChkErrCt Then
      If p_nNetError > 100 And p_nNetError < 200 Then
         p_oMod.ShowWarning "Application has Reached 100 ApplicatiOn Error!!!" & vbCrLf & _
                        "Please Inform the GCC-SEG for this Application to Avoid Further Damages!!!"
      ElseIf p_nNetError > 200 Then
         ShowError "Error Limit has been Reached!!!" & vbCrLf & _
                     "The Application will Locked to Avert Further Damages!!!"
         GoTo endProc
      End If
   End If
   
   With lors
      p_sClientId = .Fields("sClientID")
      p_sClientNm = .Fields("sClientNm")
      p_sAddressx = .Fields("sAddressx")
      p_sTownName = .Fields("sTownName")
      p_sZippCode = .Fields("sZippCode")
      p_sProvName = .Fields("sProvName")
      p_sTelNoxxx = .Fields("sTelNoxxx")
      p_sFaxNoxxx = .Fields("sFaxNoxxx")
      p_sApproved = .Fields("sApproved")
      p_sBranchCd = .Fields("sBranchCd")
      p_sProdctID = .Fields("sProdctID")
      p_sProdctNm = .Fields("sProdctNm")
      p_sApplPath = .Fields("sApplPath")
      p_sReptPath = .Fields("sReptPath")
      p_sImgePath = .Fields("sImgePath")
      p_sSysAdmin = .Fields("sSysAdmin")
      p_sNetWarex = .Fields("sNetWarex")
      p_sMachinex = .Fields("sMachinex")
      p_dSysDatex = .Fields("dSysDatex")
      p_dLicencex = .Fields("dLicencex")
      p_nNetError = .Fields("nNetError")
      p_sSkinCode = .Fields("sSkinCode")
      p_sCompName = IIf(p_sCompName = lsMyStation, lsMyStation, .Fields("sComptrNm"))
      p_sBranchNm = .Fields("sBranchNm")
      p_cWareHous = .Fields("cWareHous")
      p_cMainOffc = .Fields("cMainOffc")
   End With
   
   ' decrypt the encrypted fields
   'loCrypt.InBuffer = p_sNetWarex
   'loCrypt.Decrypt
   'p_sNetWarex = loCrypt.OutBuffer
   p_sNetWarex = Decrypt(p_sNetWarex, xeSignature)
   Debug.Print p_sNetWarex
  
   'loCrypt.InBuffer = p_sMachinex
   'loCrypt.Decrypt
   'p_sMachinex = loCrypt.OutBuffer
   p_sMachinex = Decrypt(p_sMachinex, xeSignature)
   Debug.Print p_sMachinex
   
   ' check the authors signature
   If isSignatureOK = False Then GoTo endProc
   
   Call LoadOthers
   
   p_bAllowUpd = (LCase(HostName) = LCase(p_oRSConfg("sDBHostNm")))
         
   ' check system date
   If Format(lors("dSysDatex").Value, "YYYY-MM-DD") > Format(ServerDate, "YYYY-MM-DD") Then
      ShowError "Server's Date was Regressed!!!" & vbCrLf & _
               "Back Date is One Cause of Data Integrity Error!!!"
      GoTo endProc
   Else
      'kalyptus - 2013.12.26 12:48pm
      ' if day-2-day task is not active then update the sytem date
      '     if value of dOpenD2Dx is null then day-2-day task is not active
      ' else let the day-2-day task(openDay2Day) do the update
      
      If Not IsDate(p_oRSConfg("dOpenD2Dx")) Then
         p_dSysDatex = Format(ServerDate, "YYYY-MM-DD")
         
         lbUpdateDt = True
         
         If Not (p_cMainOffc = "1" Or p_cWareHous = "1") Then
            If Left(LCase(HostName), 9) <> Left(LCase(p_oRSConfg("sDBHostNm")), 9) Then lbUpdateDt = False
         End If
               
         If lbUpdateDt Then
            psSQL = "UPDATE xxxSysApplication SET" & _
                        " dSysDatex = " & p_oMod.dateParm(p_dSysDatex) & _
                     " WHERE sClientID = " & p_oMod.strParm(p_sClientId) & _
                       " AND sProdctID = " & p_oMod.strParm(p_sProdctID)
            p_ocnMain.Execute psSQL
         End If
      End If
   End If
   
   lors.Close

   If Not IsMissing(sUserID) Then
      'Retrieve the date and time from the server and set it as the date of this computer
'505         Date = ServerDate
'510         Time = Format(ServerDate, "HH:MM:SS")
      ImportErrorLog
      LoadEnv = LogUser(sUserID)
      If LoadEnv = False Then GoTo endProc
      getSkin p_sSkinCode
   End If
   LoadEnv = True
   
endProc:
   Set lors = Nothing
   'Set loCrypt = Nothing
   Set loINI = Nothing
   
   Exit Function
errProc:
   ShowError lsProcName & "( " & sProduct _
                       & ", " & p_oMod.IFNull(sUserID) & ")"
End Function

Function getUserName(ByVal sUserID As String) As String
   Dim lors As ADODB.Recordset
   Dim lsProcName As String
   
   lsProcName = "getUserName"
   On Error GoTo errProc
   
   If sUserID = Empty Then GoTo endProc
   
   Set lors = New Recordset
   lors.Open "SELECT sUserName" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx = " & p_oMod.strParm(sUserID), p_ocnMain, , , adCmdText
   
   If Not lors.EOF Then getUserName = Decrypt(lors("sUserName"), Machinex)
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sUserID & " )"
End Function

Function getLogName(ByVal sUserID As String) As String
   Dim lors As ADODB.Recordset
   Dim lsProcName As String
   
   lsProcName = "getLogName"
   On Error GoTo errProc
   
   If sUserID = Empty Then GoTo endProc
   
   Set lors = New Recordset
   lors.Open "SELECT sLogNamex" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx = " & p_oMod.strParm(sUserID), p_ocnMain, , , adCmdText
   
   If Not lors.EOF Then getLogName = Decrypt(lors("sLogNamex"), Machinex)
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sUserID & " )"
End Function

Function getColor(ColorCode As String) As Long
   Dim lnCode As Integer
   Dim lnTemp As Double
   
   lnCode = InStr(1, p_sFldName, ColorCode, vbTextCompare)
   lnTemp = lnCode / 4
   lnCode = Int(lnTemp)
   getColor = p_anColor(lnCode)
End Function

Function getPicture(FormType As Integer) As String
   Select Case FormType
   Case xeFormMaintenance
      getPicture = p_sFormPict
   Case xeFormTransaction
      getPicture = p_sFormPict
   Case xeFormQuickSearch
      getPicture = p_sQSImage
   End Select
End Function
   
Sub getSkin(lsSkin As String)
   Dim loSkin As ADODB.Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCtr As Integer

   lsProcName = "getSkin"
   On Error GoTo errProc
   
   Set loSkin = New Recordset

   lsSQL = "SELECT *" & _
            " FROM xxxSkin" & _
            " WHERE sSkinCode = " & p_oMod.strParm(lsSkin)
   loSkin.Open lsSQL, p_ocnMain, , , adCmdText
   
   ReDim p_anColor(loSkin.Fields.Count - 1) As Long
   p_sFldName = ""
   If loSkin.EOF Then
      For lnCtr = 2 To loSkin.Fields.Count - 3
         p_sFldName = p_sFldName & Right(loSkin(lnCtr).Name, 3) & ";"
         p_anColor(lnCtr - 2) = 0
      Next
      p_sQSImage = ""
      p_sFormPict = ""
   Else
      For lnCtr = 2 To loSkin.Fields.Count - 3
         p_sFldName = p_sFldName & Right(loSkin(lnCtr).Name, 3) & ";"
         p_anColor(lnCtr - 2) = loSkin(lnCtr)
      Next
      p_sQSImage = loSkin("sQSImagex")
      p_sFormPict = loSkin("sTitleBar")
   End If
   
endProc:
   Set loSkin = Nothing
   Exit Sub
errProc:
   ShowError lsProcName & "( " & lsSkin & " )"
End Sub

'mac 2021.04.19
'  add save entry to TMRS_Log(Transaction Mirroring Replication System) table
Function Add2TMRS(ByVal fsSourceNo As String, _
                  ByVal fsSourceCD As String, _
                  Optional ByVal fcTranStat As xeTransactionStatus = xeStateOpen, _
                  Optional ByVal fcSendStat As xeLogical = xeNo) As Long
                  
   Dim lsSQL As String
   Dim lsBranchCD As String
   Dim lsSendStat As String
   
   lsBranchCD = Right(p_sClientId, 4)
   
   lsSendStat = fcSendStat
   'tag as sent if branch is gco/gcc
   If LCase(lsBranchCD) = "m001" Or _
      LCase(lsBranchCD) = "m0w1" Then
      lsSendStat = xeYes 'sent
   End If
   
   lsSQL = "INSERT INTO TMRS_Log SET" & _
            "  sTransNox = " & p_oMod.strParm(p_oMod.GetNextCode("TMRS_Log", "sTransNox", True, p_ocnMain, True, lsBranchCD)) & _
            ", sSourceNo = " & p_oMod.strParm(fsSourceNo) & _
            ", sSourceCD = " & p_oMod.strParm(fsSourceCD) & _
            ", cTranStat = " & p_oMod.strParm(fcTranStat) & _
            ", cSendStat = " & p_oMod.strParm(lsSendStat) & _
            ", dCreatedx = " & p_oMod.dateParm(ServerDate())
            
   Add2TMRS = Execute(lsSQL, "TMRS_Log", lsBranchCD)
End Function

Function Execute(ByVal sQuery As String, _
                  ByVal sTableNme As String, _
                  Optional sBranchCd As Variant, _
                  Optional sDestinat As Variant = "") As Long
   Dim lsSQL As String
   Dim lnRow As Long
   Dim lsBranchCD As String
   Dim lbWithTran As Boolean
   
   lsBranchCD = Right(p_sClientId, 4)

   If sQuery = Empty Then Exit Function
   
   'Kalyptus - 2014-11-26 02:45pm
   'Do not allow execution of queries beyond 2048 in length
   If Len(sQuery) >= 2048 Then
      Err.Raise "30016", "AppDriver", "Query too long. Unable to Register SQL Query Statement"
      Exit Function
   ElseIf Not p_bAllowUpd Then
      Err.Raise "30015", "AppDriver", "OFFLINE detected. Connection does not allow update of records."
      Exit Function
   End If
      
   If Not p_bBypass Then
      If Not p_bWithTran Then
         lbWithTran = True
         p_ocnMain.BeginTrans
         'Call p_ocnMain.Execute("SET SESSION tx_isolation='SERIALIZABLE';", , adCmdText)
      End If
   End If
      
   lsSQL = "INSERT INTO xxxReplicationLog" & _
               " SET sTransNox = " & p_oMod.strParm(p_oMod.GetNextCode("xxxReplicationLog", "sTransNox", True, _
                     p_ocnMain, True, lsBranchCD)) & _
               ", sBranchCd = " & p_oMod.strParm(IIf(IsMissing(sBranchCd), _
                     lsBranchCD, sBranchCd)) & _
               ", sStatemnt = " & p_oMod.strParm(sQuery) & _
               ", sTableNme = " & p_oMod.strParm(sTableNme) & _
               ", sDestinat = " & p_oMod.strParm(sDestinat) & _
               ", sModified = " & p_oMod.strParm(p_sUserIDxx) & _
               ", dEntryDte = " & p_oMod.dateParm(ServerDate()) & _
               ", dModified = " & p_oMod.dateParm(ServerDate())
   p_ocnMain.Execute lsSQL, lnRow, adCmdText
   
   If lnRow <= 0 Then
      Err.Raise "30013", "AppDriver", "Statement Error. Unable to Register SQL Query Statement"
      Exit Function
   End If
   
   lnRow = 0
   Debug.Print sQuery
   p_ocnMain.Execute sQuery, lnRow, adCmdText
   Execute = lnRow

   If Not p_bBypass Then
      If lbWithTran Then
         p_ocnMain.CommitTrans
         'Call p_ocnMain.Execute("SET SESSION tx_isolation='REPEATABLE-READ';", , adCmdText)
      End If
   End If
End Function

Sub BeginTrans()
   If Not p_bWithTran Then
      p_bWithTran = True
      p_ocnMain.BeginTrans
      'Call p_ocnMain.Execute("SET SESSION tx_isolation='SERIALIZABLE';", , adCmdText)
   Else
      Err.Raise "30014", "AppDriver", "Multiple Transaction Initiation Not Allowed."
      Exit Sub
   End If
End Sub

Sub CommitTrans()
   If p_bWithTran Then
      p_bWithTran = False
      p_ocnMain.CommitTrans
      'Call p_ocnMain.Execute("SET SESSION tx_isolation='REPEATABLE-READ';", , adCmdText)
   End If
End Sub

Sub RollBackTrans()
   If p_bWithTran Then
      p_bWithTran = False
      p_ocnMain.RollBackTrans
      'Call p_ocnMain.Execute("SET SESSION tx_isolation='REPEATABLE-READ';", , adCmdText)
   End If
End Sub

Function RegisDelete(ByVal sQuery As String) As Boolean
   Dim lsSQL As String
   Dim lnRow As Long
   
   If sQuery = Empty Then Exit Function
   lsSQL = "INSERT INTO xxxDeletedRec " & _
               "( sTransNox" & _
               ", sBranchCd" & _
               ", sStatemnt" & _
               ", dModified" & _
            " ) VALUES ( " & _
               p_oMod.strParm(p_oMod.GetNextCode("xxxDeletedRec", "sTransNox", True, _
                     p_ocnMain, True, p_sBranchCd)) & _
               ", " & p_oMod.strParm(p_sBranchCd) & _
               ", " & p_oMod.strParm(sQuery) & _
               ", " & p_oMod.dateParm(ServerDate()) & " )"

   p_ocnMain.Execute lsSQL, lnRow, adCmdText
   If lnRow <= 0 Then
      Err.Raise "30013", "AppDriver", "Unable to Register SQL Query Statement"
      Exit Function
   End If
   RegisDelete = True
End Function

Function LogUser(ByVal sUserID As String) As Boolean
   Dim lors As ADODB.Recordset
   'Dim loCrypt As clsCrypto
   Dim lsUser As String
   Dim lsPassword As String
   Dim lbLogIn As Boolean
   Dim lsProcName As String
   Dim lnRow As Integer
   
   lsProcName = "LogUser"
   On Error GoTo errProc
   
   LogUser = False
   
   psSQL = "SELECT a.*, b.sEmployID, b.cRecdStat" & _
            " FROM xxxSysUser a" & _
               " LEFT JOIN Employee_Master001 b" & _
                  " ON a.sEmployNo = b.sEmployID" & _
            " WHERE a.sUserIDxx = " & p_oMod.strParm(sUserID)

   Set lors = New Recordset
   lors.Open psSQL, Connection, , , adCmdText
   Debug.Print psSQL
   If lors.EOF Then
      MsgBox "Invalid User Detected!"
      GoTo endProc
   End If
   
   p_sUserIDxx = lors("sUserIDxx")
   p_sLogNamex = lors("sLogNamex")
   p_sUserName = lors("sUserName")
   p_nUserLevl = lors("nUserLevl")
   p_sSkinCode = lors("sSkinCode")
   p_sEmployNo = p_oMod.IFNull(lors("sEmployID"), "")
   
   'mac 2019.02.12
   '  user account must have employee id
   '  action confirmed by sir marlon
   If p_oMod.IFNull(p_sEmployNo) = "" Then
      MsgBox "Employee ID is not detected on user's account info." & vbCrLf & vbCrLf & _
               "Please call MIS Support Group for configuration.", vbCritical, "Warning"
      GoTo endProc
   End If
   
   'mac 2022.11.11
   '  inactive employee accounts must be prohibited to login
   '  requested by HR Department
   If lors("cRecdStat") = xeNo Then
      MsgBox "Unable to login INACTIVE employee account.", vbCritical, "Warning"
      GoTo endProc
   End If

   'Set loCrypt = New clsCrypto
   'loCrypt.Signature = Machinex

   'loCrypt.InBuffer = p_sUserName
   'loCrypt.Decrypt
   'p_sUserName = loCrypt.OutBuffer
   p_sUserName = Decrypt(p_sUserName, Machinex)

   'loCrypt.InBuffer = p_sLogNamex
   'loCrypt.Decrypt
   'p_sLogNamex = loCrypt.OutBuffer
   p_sLogNamex = Decrypt(p_sLogNamex, Machinex)

   ' user with administrator level up don't follow standard login procedure
   If p_nUserLevl < xeSysAdmin Then
      ' check first if user is currently login to another computer
      If lors("cLogStatx") = "1" Then
         MsgBox "You are currently logged to another workstation!!!" & vbCrLf & _
                  "Loging to two station simultaneously is not permitted!!!" & vbCrLf & vbCrLf & _
                  App.EXEName & " " & App.Major & "." & App.Minor & " will now quit!!!", _
                  vbInformation, "Warning"
         GoTo endProc
      End If

      ' check if user right is lock or not
      If lors("cLockStat") = xeLockStateLock Then
         If Not UnlockUser Then
            MsgBox "Unlocking of user rights failed!!!" & vbCrLf & _
                  App.EXEName & " " & App.Major & "." & App.Minor & " will now quit!!!", _
                  vbInformation, "Warning"
            GoTo endProc
         End If
      End If
   End If

   With p_oMod
      p_sLogNoxxx = .GetNextCode("xxxSysUserLog", "sLogNoxxx", True, p_ocnMain, True, p_sBranchCd)
   
      ' insert the user log
      psSQL = "INSERT INTO xxxSysUserLog (" & _
                  "  sLogNoxxx" & _
                  ", sUserIDxx" & _
                  ", dLogInxxx" & _
                  ", sProdctID" & _
                  ", sComptrNm" & _
               ") VALUES (" & _
                  .strParm(p_sLogNoxxx) & _
                  ", " & .strParm(p_sUserIDxx) & _
                  ", " & .dateParm(ServerDate()) & _
                  ", " & .strParm(p_sProdctID) & _
                  ", " & .strParm(p_sCompName) & ")"
                  
      p_ocnMain.Execute psSQL, lnRow, adCmdText
      If lnRow <= 0 Then
         MsgBox "Unable to Insert User Log Info!!!", vbCritical, "Warning"
         GoTo endProc
      End If
      
   End With
   LogUser = True

endProc:
   'Set loCrypt = Nothing

   Exit Function
errProc:
   ShowError lsProcName & "( " & sUserID & " )"
End Function

Function LogOutUser() As Boolean
   Dim lsProcName As String
   Dim lnRow As Long
   
   lsProcName = "LogOutUser"
   On Error GoTo errProc
   
   ' logout user
   With p_oMod
      psSQL = "UPDATE xxxSysUser SET" & _
                  "  cLogStatx = " & .strParm(xeLogStateOut) & _
                  ", cLockStat = " & .strParm(xeNo) & _
                  ", sCompName = ''" & _
               " WHERE sUserIdxx = " & .strParm(p_sUserIDxx)

      p_ocnMain.Execute psSQL, lnRow, adCmdText
      If lnRow <= 0 Then
         MsgBox "Unable to Update User Info!!!", vbCritical, "Warning"
         GoTo endProc
      End If
         
      psSQL = "UPDATE xxxSysUserLog SET" & _
                  " dLogOutxx = " & .dateParm(ServerDate()) & _
               " WHERE sLogNoxxx = " & .strParm(p_sLogNoxxx) & _
                 " AND sUserIDxx = " & .strParm(p_sUserIDxx)
               
      p_ocnMain.Execute psSQL, lnRow, adCmdText
      If lnRow <= 0 Then
         MsgBox "Unable to Update User Log Info!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End With
   
   LogOutUser = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function
'Function setConfiguration(ConfigCd As String, Optional BranchCd As Variant) As Boolean
'
'End Function

Function getConfiguration(ConfigCd As String, Optional BranchCd As Variant) As String
   Dim lors As ADODB.Recordset
   Dim lsSQL As String
   Dim lsProcName As String
   Dim loField As ADODB.Field
   
   lsProcName = "getConfiguration"
   On Error GoTo errProc
   
   If IsMissing(BranchCd) Then
      lsSQL = "SELECT sValuexxx" & _
               " FROM xxxOtherConfig" & _
               " WHERE sProdctID = " & p_oMod.strParm(p_sProdctID) & _
                  " AND sConfigId = " & p_oMod.strParm(ConfigCd)
   Else
      lsSQL = "SELECT " & ConfigCd & " sValuexxx" & _
               " FROM xxxOtherInfo a" & _
                  " LEFT JOIN xxxSysClient b" & _
                     " ON a.sClientID = b.sClientID" & _
               " WHERE sBranchCd = " & p_oMod.strParm(BranchCd)
   End If

   Set lors = New Recordset
   With lors
      .Open lsSQL, p_ocnMain, , , adCmdText
      Debug.Print
      If .EOF Then GoTo endProc
            
      getConfiguration = .Fields("sValuexxx")
      
      .Close
   End With

endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Sub SetMDIMain()
   Dim lors As New Recordset
   Dim loCtrl As Control
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lsName As String
   Dim lnCtr As Integer
   
   Dim loFileSys As FileSystemObject
   
   lsProcName = "SetMDIMain"
   On Error GoTo errProc
   
   If UserLevel = 128 Then GoTo endProc
   
   'Create a new way of loading rights for the menu
   lsName = p_sApplPath & "\" & "config\" & p_sProdctID & "." & p_sUserIDxx & ".mn1"
   
   Set loFileSys = New Scripting.FileSystemObject
   If Not loFileSys.FolderExists(p_sApplPath & "\" & "config\") Then
      Call loFileSys.CreateFolder(p_sApplPath & "\" & "config\")
   End If
   
   Set lors = New Recordset
   If Not p_oMod.isFileExist(lsName) Then
      lsSQL = "SELECT" & _
                  "  a.sMenuIDxx" & _
                  ", a.sMenuName" & _
                  ", a.nUserRght" & _
                  ", a.nAddRight" & _
                  ", a.nUpdRight" & _
                  ", a.nCanRight" & _
                  ", a.nDelRight" & _
                  ", a.nAppRight" & _
                  ", b.nAddRight xAddRight" & _
                  ", b.nUpdRight xUpdRight" & _
                  ", b.nCanRight xCanRight" & _
                  ", b.nDelRight xDelRight" & _
                  ", b.nAppRight xAppRight" & _
               " FROM xxxMenuObject a" & _
                     " LEFT JOIN xxxMenuRightsObject b" & _
                        " ON a.sMenuIDxx = b.sMenuIDxx" & _
                           " AND b.sUserIDxx = " & p_oMod.strParm(p_sUserIDxx) & _
               " WHERE a.sProdctID = " & p_oMod.strParm(ProductID)
      lors.Open lsSQL, p_ocnMain, , , adCmdText
      
      If Not lors.EOF Then
         'Check here whether the directory is existing....
         lors.Save lsName, adPersistXML
      End If
   Else
      Set lors = New Recordset
      lors.Open lsName
   End If
   Set lors.ActiveConnection = Nothing
   
   For Each loCtrl In p_oMDIForm.Controls
      If TypeName(loCtrl) = "Menu" Then
         lsName = loCtrl.Name
         If lors.RecordCount > 0 Then
            lors.MoveFirst
            lors.Find "sMenuName = " & p_oMod.strParm(lsName), , adSearchForward, adBookmarkFirst
         End If
         
         If Not lors.EOF Then
            loCtrl.Tag = lors("sMenuIDxx") & _
                           "»" & lors("nAddRight") & _
                           "»" & lors("nUpdRight") & _
                           "»" & lors("nCanRight") & _
                           "»" & lors("nDelRight") & _
                           "»" & lors("nAppRight") & _
                           "»" & IIf(IsNull(lors("xAddRight")), "", lors("xAddRight")) & _
                           "»" & IIf(IsNull(lors("xUpdRight")), "", lors("xUpdRight")) & _
                           "»" & IIf(IsNull(lors("xCanRight")), "", lors("xCanRight")) & _
                           "»" & IIf(IsNull(lors("xDelRight")), "", lors("xDelRight")) & _
                           "»" & IIf(IsNull(lors("xAppRight")), "", lors("xAppRight"))

            If (lors("nUserRght") And UserLevel) = 0 Then
               On Error Resume Next
               If loCtrl.Visible Then loCtrl.Visible = False
'               On Error GoTo errProc
            End If
         End If
      End If
   Next
   
'   For Each loCtrl In p_oMDIForm.Controls
'      If TypeName(loCtrl) = "Menu" Then
'         lsName = loCtrl.Name
'         lsSQL = "SELECT" & _
'                     "  a.sMenuIDxx" & _
'                     ", a.nUserRght" & _
'                     ", a.nAddRight" & _
'                     ", a.nUpdRight" & _
'                     ", a.nCanRight" & _
'                     ", a.nDelRight" & _
'                     ", a.nAppRight" & _
'                     ", b.nAddRight xAddRight" & _
'                     ", b.nUpdRight xUpdRight" & _
'                     ", b.nCanRight xCanRight" & _
'                     ", b.nDelRight xDelRight" & _
'                     ", b.nAppRight xAppRight" & _
'                  " FROM xxxMenuObject a" & _
'                        " LEFT JOIN xxxMenuRightsObject b" & _
'                           " ON a.sMenuIDxx = b.sMenuIDxx" & _
'                              " AND b.sUserIDxx = " & p_oMod.strParm(p_sUserIDxx) & _
'                  " WHERE a.sProdctID = " & p_oMod.strParm(ProductID) & _
'                     " AND a.sMenuName = " & p_oMod.strParm(lsName)
'         lors.Open lsSQL, p_ocnMain, , , adCmdText
'
'         If Not lors.EOF Then
'            loCtrl.Tag = lors("sMenuIDxx") & _
'                           "»" & lors("nAddRight") & _
'                           "»" & lors("nUpdRight") & _
'                           "»" & lors("nCanRight") & _
'                           "»" & lors("nDelRight") & _
'                           "»" & lors("nAppRight") & _
'                           "»" & IIf(IsNull(lors("xAddRight")), "", lors("xAddRight")) & _
'                           "»" & IIf(IsNull(lors("xUpdRight")), "", lors("xUpdRight")) & _
'                           "»" & IIf(IsNull(lors("xCanRight")), "", lors("xCanRight")) & _
'                           "»" & IIf(IsNull(lors("xDelRight")), "", lors("xDelRight")) & _
'                           "»" & IIf(IsNull(lors("xAppRight")), "", lors("xAppRight"))
'
'            If (lors("nUserRght") And UserLevel) = 0 Then
'               On Error Resume Next
'               If loCtrl.Visible Then loCtrl.Visible = False
'               On Error GoTo errProc
'            End If
'         End If
'         lors.Close
'      End If
'   Next
      
   With p_oMDIForm
      lnCtr = .Statusbar1.Panels.Count
      .Statusbar1.Panels(lnCtr - 2).Text = ClientName + " (" + IIf(p_bAllowUpd, "ONLINE", "OFFLINE") + ")"
      .Statusbar1.Panels(lnCtr - 1).Text = UserName
      .Statusbar1.Panels(lnCtr).Text = Format(SysDate, "mm/dd/yyyy")
   End With
   
endProc:
   Set lors = Nothing
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Private Function UnlockUser() As Boolean
   Dim lors As ADODB.Recordset
   Dim loForm As frmUnlock
   Dim loCrypt As clsCrypto
   Dim loMod As clsMainModules
   Dim lsUserName As String
   Dim lsPassword As String
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lbLogIn As Boolean
   
   lsProcName = "UnlockUser"
   On Error GoTo errProc
   
   Set lors = New Recordset
   Set loCrypt = New clsCrypto
   loCrypt.Signature = p_sMachinex
   Set loForm = New frmUnlock
   
   lnCtr = 0
   lbLogIn = False
   Do
      loForm.AppPath = p_sApplPath
      loForm.Show 1
      If loForm.Cancel = True Then
         ' user either clicked the cancel button or the close key
         lnCtr = 2
      ElseIf loForm.xrButton(2) = True Then
         ' user either clicked the close key
         lnCtr = 3
         GoTo endProc
      Else
         ' Encrypt the username and password
         'loCrypt.InBuffer = loForm.txtPassword
         'loCrypt.Encrypt
         'lsPassword = loCrypt.OutBuffer
         lsPassword = Encrypt(loForm.txtPassword, xeSignature)
         
         'loCrypt.InBuffer = LCase(loForm.txtUserName)
         'loCrypt.Encrypt
         'lsUserName = loCrypt.OutBuffer
         lsUserName = Encrypt(LCase(loForm.txtUserName), xeSignature)
         
         With p_oMod
            ' search for the user
            lsSQL = "SELECT *" & _
                     " FROM xxxSysUser" & _
                     " WHERE sLogNamex = " & .strParm(lsUserName) & _
                        " AND sPassword = " & .strParm(lsPassword)
         End With
         
         If lors.State = adStateOpen Then lors.Close
         
         lors.Open lsSQL, Connection, , , adCmdText
         If lors.EOF Then
            MsgBox "Verify your log name and/or password", vbCritical, "Login Error"
         Else
            If lors("nUserLevl") < xeManager Then
               MsgBox "User does Not Have the Right to Unlock your Account!!!" & vbCrLf & _
                        "Please Seek User with Rights to Unlock Account!!!", vbCritical, "Warning"
            Else
               lbLogIn = True
            End If
         End If
      End If
      lnCtr = lnCtr + 1
   Loop Until (lbLogIn = True) Or (lnCtr = 3)
   
   UnlockUser = lbLogIn
   
endProc:
   Unload loForm
   Set lors = Nothing
   Set loCrypt = Nothing
   Set loForm = Nothing
   
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Sub Class_Initialize()
   pbErrorLog = False
   psModlName = "AppDriver"
   pbChkErrCt = True
   
   Set p_ocnMain = New Connection
End Sub

Public Function ServerDate() As Variant
   Dim lors As ADODB.Recordset
   
   'Set lors = p_ocnMain.Execute("SELECT SYSDATE()")
   Set lors = p_ocnMain.Execute("SELECT NOW()")
   ServerDate = lors(0)
   
   Set lors = Nothing
End Function

Private Sub Class_Terminate()
   Set p_ocnMain = Nothing
   Set p_oMod = Nothing
End Sub

Private Function isSignatureOK() As Boolean
   Dim lors As ADODB.Recordset
   Dim lsSQL As String
   Dim lsProcName As String
   
   
   lsProcName = "isSignatureOK"
   On Error GoTo errProc
   
   If p_nHexCrypt <> 1 Then
      isSignatureOK = True
      GoTo endProc
   End If
   
   If p_sMachinex <> xeSignature Then GoTo endProc
   
   ' check the existence of the author and the system admin
   lsSQL = "SELECT sUserIDxx" & _
               ", sUserName" & _
               ", sLogNamex" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx in (" & _
               p_oMod.strParm(p_sMachinex) & ", " & p_oMod.strParm(p_sNetWarex) & ")"
   
   Set lors = New Recordset
   
   lors.Open lsSQL, p_ocnMain, , , adCmdText
   
   With lors
      If .RecordCount <> 2 Then
         MsgBox "Unregistered Copy of " & p_sProdctNm & " Detected!", vbCritical, "Warning"
         GoTo endProc
      End If
      
      Do
         If .Fields("sUserIDxx") = p_sNetWarex Then
            If Trim(.Fields("sLogNamex")) <> Trim(p_sSysAdmin) Then
               MsgBox "Unregistered Copy of " & p_sProdctNm & " Detected!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
         .MoveNext
      Loop While Not .EOF
   End With
   isSignatureOK = True
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Sub ShowError(ByVal lsProcName As String)
   xLogError Err.Number, Err.Description, "clsAppDriver", lsProcName, Erl
   With Err
'      .Raise .Number, .Source, .Description
   End With
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
Sub xLogError( _
    ByVal nErrNo As Variant, _
    sErrDesc As String, _
    sModule As String, _
    sMethod As String, _
    ByVal nErrLine As Variant _
    )

   Dim intLogFile As Integer
   
   On Error GoTo errProc
      
   '* Obtain a free file handle.
   intLogFile = FreeFile
   
   '* Open the error log text file in Append mode.
   Open p_sApplPath & "\Temp\" & "xxxTempxxx.log" For Append As #intLogFile
   
   '* Write the entry from the temporary log file
   Write #intLogFile, LogNumber, _
                      ClientId, _
                      ProductID, _
                      ComputerID, _
                      UserID, _
                      Format(Now, "YYYY-MM-DD HH:MM:SS"), _
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
Sub xShowError()
   Dim lsLogNoxxx As String
   Dim lsClientID As String
   Dim lsProductId As String
   Dim lsComputerID As String
   Dim lsUserID As String
   Dim ldDateErr As String
   Dim lnErrNo As Integer
   Dim lsErrDesc As String
   Dim lsModule As String
   Dim lsMethod As String
   Dim lnErrLine As Integer

   Dim lsMessage As String
   Dim strCaption As String
   
   Dim ofrm As New frmShowError
   
   On Error GoTo errProc
   
   'Open and close the temporarry file in append mode to make sure that it exists
'160      Open p_sApplPath & "\" & "xxxTempxxx.log" For Append As #1
'170      Close #1
'180      Kill p_sApplPath & "\" & "xxxTempxxx.log"

   '* Open the necessary files
   Open p_sApplPath & "\Temp\" & "xxxTempxxx.log" For Input As #1
   Open p_sApplPath & "\" & "GGCISys.log" For Append As #2

   lsMessage = ""
   
   Do Until EOF(1)
      '* Read the contents of the temporary error log file
      Input #1, lsLogNoxxx, _
                lsClientID, _
                lsProductId, _
                lsComputerID, _
                lsUserID, _
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
                lsUserID, _
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
   Kill p_sApplPath & "\Temp\" & "xxxTempxxx.log"
   
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
'kalyptus-2013.05.09
'Disable ImportErrorLog
'10       Dim lsOldLogNo As String
'20       Dim lsTranCode As String
'30       Dim lsTranLine As String
'40       Dim lsLogNoxxx As String
'50       Dim lsClientID As String
'60       Dim lsProductId As String
'70       Dim lsComputerID As String
'80       Dim lsUserID As String
'90       Dim ldDateErr As String
'100      Dim lnErrNo As Integer
'110      Dim lsErrDesc As String
'120      Dim lsModule As String
'130      Dim lsMethod As String
'140      Dim lnErrLine As Integer
'
'150      Dim lbHasErr As Boolean
'160      Dim loRec As ADODB.Recordset
'170      Dim lsSQL As String
'
'180      On Error GoTo errProc
'
'   'Open and close the temporarry file in append mode to make sure that it exists
'190      Open p_sApplPath & "\" & "xxxTempxxx.log" For Append As #1
'200      Close #1
'
'   '* Open the necessary files
'210      Open p_sApplPath & "\" & "xxxTempxxx.log" For Input As #1
'220      Open p_sApplPath & "\" & "GGCISys.log" For Append As #2
'
'230      Do Until EOF(1)
'      '* Read the contents of the temporary error log file
'240         Input #1, lsLogNoxxx, _
'                lsClientID, _
'                lsProductId, _
'                lsComputerID, _
'                lsUserID, _
'                ldDateErr, _
'                lnErrNo, _
'                lsErrDesc, _
'                lsModule, _
'                lsMethod, _
'                lnErrLine
'250         If Trim(lsClientID) = "" Then
'260            lsClientID = ClientId
'270         End If
'
'280         If Trim(lsProductId) = "" Then
'290            lsProductId = "gRider"
'300         End If
'
'310         If Trim(lsComputerID) = "" Then
'320            lsComputerID = p_sCompName
'330         End If
'
'340         If Trim(lsUserID) = "" Then
'350            lsUserID = "N-O-N-E"
'360         End If
'
'      '* Write the data to the main error log file
'370         Write #2, lsLogNoxxx, _
'                lsClientID, _
'                lsProductId, _
'                lsComputerID, _
'                lsUserID, _
'                ldDateErr, _
'                lnErrNo, _
'                lsErrDesc, _
'                lsModule, _
'                lsMethod, _
'                lnErrLine
'380      Loop
'390      Close #1
'400      Close #2
'
'   '* Open and close the log file in append mode to make sure that it exists
''   Open p_sApplPath & "\" & "GGCISys.log" For Append As #1
''   Close #1
'
'410      Open p_sApplPath & "\" & "GGCISys.log" For Input As #1
'420      lsTranLine = "00"
'430      lsOldLogNo = "XXX"
'440      Do Until EOF(1)
'      '* Read the contents of the temporary error log file
'
'450         Input #1, lsLogNoxxx, _
'                lsClientID, _
'                lsProductId, _
'                lsComputerID, _
'                lsUserID, _
'                ldDateErr, _
'                lnErrNo, _
'                lsErrDesc, _
'                lsModule, _
'                lsMethod, _
'                lnErrLine
'460         With p_oMod
'470            If lsOldLogNo <> lsLogNoxxx Then
'            '* Get the last transaction Code
'480               lsSQL = "SELECT sTranCode FROM xxxAETMaster" _
'                 & " WHERE sClientID = " & .strParm(lsClientID) _
'                   & " AND sProdctID = " & .strParm(lsProductId) _
'                 & " ORDER BY sTranCode DESC" _
'                 & " LIMIT 1"
'490               Set loRec = Connection.Execute(lsSQL, , adCmdText)
'500               If loRec.EOF Then
'510                  lsTranCode = Format(1, "0000000000")
'520               Else
'530                  lsTranCode = Format(loRec("sTrancode") + 1, "0000000000")
'540               End If
'550               loRec.Close
'            '* Save the logged error to the xxxAETMaster table
'560               lsSQL = "INSERT INTO xxxAETMaster" _
'                       & "( sTranCode" _
'                       & ", sLogNoxxx" _
'                       & ", sClientID" _
'                       & ", sProdctID" _
'                       & ", sCompCode" _
'                       & ", sUserIDxx" _
'                       & ", dTimeMeet" _
'                       & ", nErrNmbrx" _
'                       & ", sErrDescx" _
'                       & ", cStatCode" _
'                       & ", sModified" _
'                & ") VALUES" _
'                       & " (" & .strParm(lsTranCode) _
'                       & ", " & .strParm(lsLogNoxxx) _
'                       & ", " & .strParm(lsClientID) _
'                       & ", " & .strParm(lsProductId) _
'                       & ", " & .strParm(lsComputerID) _
'                       & ", " & .strParm(lsUserID) _
'                       & ", " & .dateParm(Replace(ldDateErr, "#", "")) _
'                       & ", " & lnErrNo _
'                       & ", " & .strParm(lsErrDesc) _
'                       & ", " & .strParm("0") _
'                       & ", " & .strParm(.Encrypt(UserID, Machinex)) _
'                & ")"
'
'570               Connection.Execute lsSQL
'
'            '* Increment the neterror of the user based on UserID
'580               lsSQL = " UPDATE xxxSysUser" _
'                  & " SET" _
'                        & " nSysError = nSysError+1" _
'                  & " WHERE" _
'                        & " sUserIDxx=" & .strParm(lsUserID)
'590               Connection.Execute lsSQL
'
'            '* Increment the neterror neterror of project
'600               lsSQL = " UPDATE xxxSysApplication" _
'                  & " SET" _
'                        & " nNetError = nNetError+1" _
'                  & " WHERE" _
'                        & " sProdctID =" & .strParm(lsProductId) _
'                    & " AND sClientID =" & .strParm(lsClientID)
'610               Connection.Execute lsSQL
'
'            '* Initialize lsTranline to be used in our ledger
'620               lsTranLine = "00"
'630               lsOldLogNo = lsLogNoxxx
'640            End If
'
'650            lsTranLine = Format(lsTranLine + 1, "00")
'
'         'Save it to the logged error to the xxxAETDetail table
'660            lsSQL = "INSERT INTO xxxAETDetail" _
'                    & "( sTranCode" _
'                    & ", sClientID" _
'                    & ", sProdctID" _
'                    & ", sTranLine" _
'                    & ", sModulexx" _
'                    & ", sProcName" _
'                    & ", nErrLinex" _
'                & ") VALUES" _
'                    & " (" & .strParm(lsTranCode) _
'                    & ", " & .strParm(lsClientID) _
'                    & ", " & .strParm(lsProductId) _
'                    & ", " & .strParm(lsTranLine) _
'                    & ", " & .strParm(lsModule) _
'                    & ", " & .strParm(lsMethod) _
'                    & ", " & lnErrLine _
'                & ")"
'670            Connection.Execute lsSQL
'680         End With
'690         lbHasErr = True
'700      Loop
'
'   '* Close the file
'710      Close #1
'
'   '* Delete the error log file
'720      Kill p_sApplPath & "\" & "GGCISys.log"
'730      Kill p_sApplPath & "\" & "xxxTempxxx.log"
'   '* Display the success message
''355      If lbHasErr Then
''360         MsgBox "Error log exported...! Thank you!", vbInformation, App.EXEName & " " & App.Major & "." & App.Minor
''365      End If
'
'endProc:
'740      Exit Sub
'errProc:
'750      Resume Next
End Sub

Sub showRemarks(Optional lsRemarks As Variant, _
                        Optional lsMode As Variant)
'   If p_bShowRem = False Then Exit Sub
   
   ' show remarks and mode in status bar
   
   If Not (p_oMDIForm Is Nothing) Then
      If Not IsMissing(lsRemarks) Then p_oMDIForm.Statusbar1.Panels(1).Text = lsRemarks
      If Not IsMissing(lsMode) Then p_oMDIForm.Statusbar1.Panels(2).Text = lsMode
   End If
End Sub

'Public Function isConnected() As Boolean
'   Dim lsSQL As String
'   Dim lors As ADODB.Recordset
'   Dim lsOldProc As String
'   Dim loINI As clsINIFile
'
'   Dim lsUserName As String
'   Dim lsUserPass As String
'   Dim lsDatabase As String
'   Dim lsServerNm As String
'   Dim loCon As Connection
'
'   On Error GoTo errProc
'   lsOldProc = "isConnected()"
'   Debug.Print lsOldProc
'
'   'Test if connection has been established
'   If TypeName(p_ocnMain) <> "Nothing" Then
'      If p_ocnMain.State = adStateOpen Then
'         Call ServerDate
'         isConnected = True
'         GoTo endProc
'      End If
'   End If
'
'
'   'Load the INI File here
'   Set loINI = New clsINIFile
'   loINI.FileName = Environ("windir") & "\GhostRiderXP.ini"
'   If Not loINI.isFileExist Then
'      Err.Raise 53
'      Err.Clear
'   End If
'
'   lsUserName = loINI.getTextValue(p_sProdctID, "UserName")
'   lsUserPass = loINI.getTextValue(p_sProdctID, "Password")
'   lsDatabase = loINI.getTextValue(p_sProdctID, "Database")
'   lsServerNm = loINI.getTextValue(p_sProdctID, "ServerName")
'
'   'Establish the connection
'   Set loCon = New ADODB.Connection
'   lsSQL = "DRIVER={MySQL ODBC 3.51 Driver};" & _
'            "SERVER=" & lsServerNm & ";" & _
'            "DATABASE=" & lsDatabase & ";" & _
'            "USER=" & Decrypt(lsUserName, "08220326") & ";" & _
'            "PASSWORD=" & Decrypt(lsUserPass, "08220326") & ";" & _
'            "OPTION=16395;"
'
'   loCon.CursorLocation = adUseClient
'   loCon.Open lsSQL
'
'   'If it reach this area, i am sure its connected
'   Set p_ocnMain = loCon
'
'   isConnected = True
'
'endProc:
'   Exit Function
'errProc:
'   GoTo endProc
'End Function

Public Function isConnected() As Boolean
   Dim lsSQL As String
   Dim lors As ADODB.Recordset
   Dim lsOldProc As String
   Dim loINI As clsINIFile
   Dim lsUserName As String
   Dim lsUserPass As String
   Dim lsDatabase As String
   Dim lsServerNm As String
   Dim loCon As Connection

   On Error GoTo errProc
   lsOldProc = "isConnected()"
   Debug.Print lsOldProc

   'Test if connection has been established
   If TypeName(p_ocnMain) <> "Nothing" Then
      If p_ocnMain.State = adStateOpen Then
         lsSQL = "X"
         'Set lors = p_ocnMain.Execute("SELECT SYSDATE()")
         Set lors = p_ocnMain.Execute("SELECT NOW()")
         Set lors = Nothing
         isConnected = True

         GoTo endProc
      End If
   End If

XRECONNECT:
   lsSQL = ""
   'Load the INI File here
   Set loINI = New clsINIFile
   loINI.FileName = Environ("windir") & "\GhostRiderXP.ini"

   If Not loINI.isFileExist Then
      Err.Raise 53
      Err.Clear
   End If

   lsUserName = loINI.getTextValue(p_sProdctID, "UserName")
   lsUserPass = loINI.getTextValue(p_sProdctID, "Password")
   lsDatabase = loINI.getTextValue(p_sProdctID, "Database")
   lsServerNm = loINI.getTextValue(p_sProdctID, "ServerName")

   'Establish the connection
   Set loCon = New ADODB.Connection
   lsSQL = "DRIVER={MySQL ODBC 3.51 Driver};" & _
            "SERVER=" & lsServerNm & ";" & _
            "DATABASE=" & lsDatabase & ";" & _
            "USER=" & Decrypt(lsUserName, "08220326") & ";" & _
            "PASSWORD=" & Decrypt(lsUserPass, "08220326") & ";" & _
            "OPTION=16395;"

   loCon.CursorLocation = adUseClient
   loCon.Open lsSQL

   'If it reach this area, i am sure its connected
   Set p_ocnMain = loCon

   isConnected = True

endProc:
   Exit Function
errProc:
   If lsSQL = "X" Then GoTo XRECONNECT
   GoTo endProc
End Function

'kalyptus - 2013.12.27 11:19am
'detect the hostname of the server
Public Function HostName() As String
   Dim lsSQL As String
   Dim lors As Recordset
   lsSQL = "SELECT @@hostname"
   
   Set lors = p_ocnMain.Execute(lsSQL, , adCmdText)
   
   HostName = lors(0)
   
   Set lors = Nothing
End Function

'Private Function openDay2Day(ByVal sUserID As String) As Boolean
'   Dim lsSQL As String
'   Dim lsSQLx As String
'   Dim loRS As Recordset
'   Dim lorsx As Recordset
'   Dim loForm As frmNewDay2Day
'   Dim lbDayOK As Boolean
'   Dim lnCtr As Integer
'   Dim lsApproved As String
'   Dim loCode As clsCodeGen
'
'   Set loForm = frmNewDay2Day
'
'   Do
'      loForm.txtDate = Format(ServerDate, "YYYY-MM-DD")
'      loForm.Show 1
'      If loForm.Cancel = True Then
'         ' user either clicked the cancel button or the close key
'         lnCtr = 2
'      Else
'         'Open Day-To-Day Record...
'         lsSQLx = "SELECT" & _
'                        "  sBranchCd" & _
'                        ", dTransact" & _
'                        ", nBalanceF" & _
'                        ", nTotlSale" & _
'                        ", nAdvPaymx" & _
'                        ", nUseAdvPy" & _
'                        ", nDepositd" & _
'                        ", nBalanceX" & _
'                        ", sApproved" & _
'                        ", cActivexx" & _
'                        ", cPostedxx" & _
'                        ", sPostedxx" & _
'                        ", dPostedxx" & _
'                 " FROM Daily_Summary" & _
'                 " WHERE sBranchCD = " & p_oMod.strParm(p_sBranchCd) & _
'                   " AND dTransact = " & p_oMod.dateParm(loForm.txtDate)
'         Set loRS = New Recordset
'         With loRS
'            .Open lsSQLx, p_ocnMain, adOpenStatic, adLockOptimistic, adCmdText
'            Set .ActiveConnection = Nothing
'
'            'Create a day-to-day record if a day-to-day record is not yet existing for this day
'            If .EOF Then
'               'Initialize the value of the first 3 fields
'               .AddNew
'               .Fields("sBranchCd") = p_sBranchCd
'               .Fields("dTransact") = loForm.txtDate
'               .Fields("nBalanceF") = 0#
'               .Fields("nTotlSale") = 0#
'               .Fields("nAdvPaymx") = 0#
'               .Fields("nUseAdvPy") = 0#
'               .Fields("nDepositd") = 0#
'               .Fields("nBalanceX") = 0#
'               .Fields("sApproved") = ""
'               .Fields("cActivexx") = "0"
'               .Fields("cPostedxx") = "0"
'               .Fields("sPostedxx") = ""
'
'               lsSQL = p_oMod.ADO2SQL(loRS, "Daily_Summary")
'               Call Execute(lsSQL, "Daily_Summary")
'            End If
'         End With
'
''         'Reopen newly created day-to-day transaction
''         If p_oMod.IFNull(lors("sBranchCD").OriginalValue, "") = "" Then
''            Set lors = New Recordset
''            With lors
''               .Open lsSQLx, p_ocnMain, adOpenStatic, adLockOptimistic, adCmdText
''               Set .ActiveConnection = Nothing
''            End With
''         End If
'
'         lbDayOK = True
'      End If
'      lnCtr = lnCtr + 1
'   Loop Until (lbDayOK = True) Or (lnCtr = 3)
'
'   p_dSysDatex = loForm.txtDate
'   lsApproved = loForm.txtApproval
'
'   Unload loForm
'   Set loForm = Nothing
'
'   If lbDayOK = False Then GoTo endProc
'
'   With loRS
'      If .Fields("cPostedxx") = "0" Then
'         If CDate(.Fields("dTransact")) < CDate(Format(ServerDate, "MM/DD/YYYY")) Then
'            'Is there an approval code entered by the user...
'            If lsApproved = "" Then
'               MsgBox "The date needs an approval code!" & _
'                      "You are not allowed to enter the system...", vbOKOnly + vbCritical, "Sytem Security"
'               GoTo endProc
'            '98 is the encode DT-> Day-To-Day
'            ElseIf Mid(lsApproved, 7, 2) <> "98" Then
'               MsgBox "Invalid approval code detected!" & _
'                      "You are not allowed to enter the system...", vbOKOnly + vbCritical, "Sytem Security"
'               GoTo endProc
'            End If
'
'            Set loCode = New clsCodeGen
'            loCode.Branch = p_sBranchCd
'            loCode.DateRequested = .Fields("dTransact")
'            loCode.IssuedBy = Mid(lsApproved, 4, 1)
'            loCode.System = pxeDay2Day
'
'            If .Fields("cActivexx") = "0" Then
'               If .Fields("sApproved") = "" Then
'                  loCode.MiscInfo = "1"
'               Else
'                  lsSQL = "SELECT sReferNox" & _
'                         " FROM xxxSCA_Usage" & _
'                         " WHERE sTransNox LIKE " & p_oMod.strParm(p_sBranchCd & "%") & _
'                           " AND sReferNox LIKE " & p_oMod.strParm("___" & Mid(lsApproved, 4)) & _
'                         " ORDER BY sTransNox DESC" & _
'                         " LIMIT 1"
'                  Set lorsx = p_ocnMain.Execute(lsSQL, , adCmdText)
'                  If lorsx.EOF Then
'                     loCode.MiscInfo = "1"
'                  Else
'                     loCode.MiscInfo = (Val("&H" & Mid(lorsx("sReferNox"), 2, 2)) - 70) + 1
'                  End If
'               End If
'
'               'Encode data
'               Call loCode.Encode
'
'               'Compare possible result with the encoded approval
'               '(0)»equal; (-)»not equal; (+)»more than
'               If loCode.Equal(loCode.Result, lsApproved) >= 0 Then
'                  lsSQL = "UPDATE Daily_Summary" & _
'                         " SET sApproved = " & p_oMod.strParm(lsApproved) & _
'                            ", cActivexx = '1'" & _
'                         " WHERE sBranchCD = " & p_oMod.strParm(p_sBranchCd) & _
'                           " AND dTransact = " & p_oMod.dateParm(p_dSysDatex)
'                  Call p_ocnMain.Execute(lsSQL, , adCmdText)
'
'
'                  lsSQL = "UPDATE Daily_Summary" & _
'                         " SET cActivexx = '0'" & _
'                         " WHERE sBranchCD = " & p_oMod.strParm(p_sBranchCd) & _
'                           " AND dTransact <> " & p_oMod.dateParm(p_dSysDatex) & _
'                           " AND cActivexx = '1'"
'                  Call p_ocnMain.Execute(lsSQL, , adCmdText)
'
''                  p_oRSConfg("sApproved") = lsApproved
'
'                  If TypeName(lorsx) <> "Nothing" Then
'                     If lorsx.EOF Then
'                        lsSQL = p_oMod.GetNextCode("xxxSCA_Usage", "sTransNox", True, _
'                                                   p_ocnMain, True, p_sBranchCd)
'                        lsSQL = "INSERT INTO xxxSCA_Usage" & _
'                               " SET sTransNox = " & p_oMod.strParm(lsSQL) & _
'                                  ", sReferNox = " & p_oMod.strParm(lsApproved) & _
'                                  ", sModified = " & p_oMod.strParm(p_oMod.Encrypt(sUserID)) & _
'                                  ", dModified = " & p_oMod.dateParm(ServerDate)
'                        Execute lsSQL, "xxxSCA_Usage"
'                     End If
'                  End If
'               Else
'                  MsgBox "Invalid approval code detected!" & _
'                         "You are not allowed to enter the system...", vbOKOnly + vbCritical, "Sytem Security"
'                  GoTo endProc
'               End If
'            Else
'               If .Fields("sApproved") = "" Then
'                  MsgBox "Invalid cached approval code detected!" & _
'                         "You are not allowed to enter the system...", vbOKOnly + vbCritical, "Sytem Security"
'                  GoTo endProc
'               Else
'                  'If the date to be updated is the current date. just compare the approval code
'                  'entered by the user with the approval code save in the system...
'                  If .Fields("sApproved") <> lsApproved Then
'                     MsgBox "Invalid approval code confirmation detected!" & _
'                            "You are not allowed to enter the system...", vbOKOnly + vbCritical, "Sytem Security"
'                     GoTo endProc
'                  End If
'               End If
'            End If
'         Else
'            lsSQL = "UPDATE xxxSysApplication" & _
'                   " SET dSysDatex = " & p_oMod.dateParm(p_dSysDatex) & _
'                   " WHERE sClientID = " & p_oMod.strParm(p_sClientId) & _
'                     " AND sProdctID = " & p_oMod.strParm(p_sProdctID)
'            Call p_ocnMain.Execute(lsSQL, , adCmdText)
'         End If
'      Else
'         MsgBox "The Day-To-Day transaction for this date was already posted!" & vbCrLf & _
'                "Updates should be made through adjustment..." & vbCrLf & _
'                "For more info please contact SEG/SSG!"
'         GoTo endProc
'      End If
'   End With
'
'   openDay2Day = True
'
'endProc:
'   Exit Function
'errProc:
'   GoTo endProc
'End Function

Private Function openDay2Day(ByVal sUserID As String) As Boolean
   Dim lsSQL As String
   Dim lsSQLx As String
   Dim lors As Recordset
   Dim lorsx As Recordset
   Dim lbDayOK As Boolean
   Dim lnCtr As Integer
   
   Dim lsApproveCD As String
   Dim lsApproveID As String
   Dim lsApproveNm As String
   
   Dim loCode As clsCodeGen
               
   'Open Day-To-Day Record...
   lsSQLx = "SELECT" & _
                  "  sBranchCd" & _
                  ", dTransact" & _
                  ", nBalanceF" & _
                  ", nTotlSale" & _
                  ", nAdvPaymx" & _
                  ", nUseAdvPy" & _
                  ", nDepositd" & _
                  ", nBalanceX" & _
                  ", sApproved" & _
                  ", cActivexx" & _
                  ", cPostedxx" & _
                  ", sPostedxx" & _
                  ", dPostedxx" & _
           " FROM Daily_Summary" & _
           " WHERE sBranchCD = " & p_oMod.strParm(p_sBranchCd) & _
             " AND dTransact = " & p_oMod.dateParm(p_dSysDatex)
   Set lors = New Recordset
   With lors
      .Open lsSQLx, p_ocnMain, adOpenStatic, adLockOptimistic, adCmdText
      Set .ActiveConnection = Nothing
      
      'Create a day-to-day record if a day-to-day record is not yet existing for this day
      If .EOF Then
         'Initialize the value of the first 3 fields
         .AddNew
         .Fields("sBranchCd") = p_sBranchCd
         .Fields("dTransact") = p_dSysDatex
         .Fields("nBalanceF") = 0#
         .Fields("nTotlSale") = 0#
         .Fields("nAdvPaymx") = 0#
         .Fields("nUseAdvPy") = 0#
         .Fields("nDepositd") = 0#
         .Fields("nBalanceX") = 0#
         .Fields("sApproved") = ""
         .Fields("cActivexx") = "0"
         .Fields("cPostedxx") = "0"
         .Fields("sPostedxx") = ""
         
         lsSQL = p_oMod.ADO2SQL(lors, "Daily_Summary")
         Call Execute(lsSQL, "Daily_Summary")
      End If
   End With
      
   With lors
      If .Fields("cPostedxx") = "0" Then
         If CDate(.Fields("dTransact")) < CDate(Format(ServerDate, "MM/DD/YYYY")) Then
            
            If p_oMod.GetCodeApproval(Me, lsApproveCD, lsApproveID, lsApproveNm) = False Then
               MsgBox "The date needs an approval code!" & _
                      "You are not allowed to enter the system...", vbOKOnly + vbCritical, "Sytem Security"
               GoTo endProc
            End If
            
            '98 is the encode DT-> Day-To-Day
            If Mid(lsApproveCD, 7, 2) <> "98" Then
               MsgBox "Invalid approval code detected!" & _
                      "You are not allowed to enter the system...", vbOKOnly + vbCritical, "Sytem Security"
               GoTo endProc
            End If
            
            Set loCode = New clsCodeGen
            loCode.Branch = p_sBranchCd
            loCode.DateRequested = .Fields("dTransact")
            loCode.IssuedBy = Mid(lsApproveCD, 4, 1)
            loCode.System = pxeDay2Day
            
            If .Fields("cActivexx") = "0" Then
               If .Fields("sApproved") = "" Then
                  loCode.MiscInfo = "1"
               Else
                  lsSQL = "SELECT REVERSE(sApprCode) sApprCode" & _
                         " FROM xxxSCA_Usage" & _
                         " WHERE sTransNox LIKE " & p_oMod.strParm(p_sBranchCd & "%") & _
                           " AND sApprCode LIKE " & p_oMod.strParm(Left(lsApproveCD, 6)) & _
                         " ORDER BY sTransNox DESC" & _
                         " LIMIT 1"
                  Set lorsx = p_ocnMain.Execute(lsSQL, , adCmdText)
                  If lorsx.EOF Then
                     loCode.MiscInfo = "1"
                  Else
                     loCode.MiscInfo = (Val("&H" & Mid(lorsx("sApprCode"), 2, 2)) - 70) + 1
                  End If
               End If
               
               'Encode data
               Call loCode.Encode
               
               'Compare possible result with the encoded approval
               '(0)»equal; (-)»not equal; (+)»more than
               If loCode.Equal(loCode.Result, lsApproveCD) >= 0 Then
                  lsSQL = "UPDATE Daily_Summary" & _
                         " SET sApproved = " & p_oMod.strParm(lsApproveCD) & _
                            ", cActivexx = '1'" & _
                         " WHERE sBranchCD = " & p_oMod.strParm(p_sBranchCd) & _
                           " AND dTransact = " & p_oMod.dateParm(p_dSysDatex) & _
                           " AND cPostedxx = '0'"
                           
                  Call p_ocnMain.Execute(lsSQL, , adCmdText)
               
               
                  lsSQL = "UPDATE Daily_Summary" & _
                         " SET cActivexx = '0'" & _
                         " WHERE sBranchCD = " & p_oMod.strParm(p_sBranchCd) & _
                           " AND dTransact <> " & p_oMod.dateParm(p_dSysDatex) & _
                           " AND cActivexx = '1'" & _
                           " AND cPostedxx = '0'"
                  Call p_ocnMain.Execute(lsSQL, , adCmdText)
               
'                  p_oRSConfg("sApproved") = lsApproved
                  
                  If TypeName(lorsx) <> "Nothing" Then
                     If lorsx.EOF Then
                        lsSQL = p_oMod.GetNextCode("xxxSCA_Usage", "sTransNox", True, _
                                                   p_ocnMain, True, p_sBranchCd)
                        lsSQL = "INSERT INTO xxxSCA_Usage" & _
                               " SET sTransNox = " & p_oMod.strParm(lsSQL) & _
                                  ", sApprCode = " & p_oMod.strParm(StrReverse(lsApproveCD)) & _
                                  ", sApproved = " & p_oMod.strParm(lsApproveID) & _
                                  ", sSystemCD = " & p_oMod.strParm(pxeDay2Day) & _
                                  ", sSourceNo = " & p_oMod.strParm("") & _
                                  ", sSourceCD = " & p_oMod.strParm("AppD") & _
                                  ", sModified = " & p_oMod.strParm(sUserID) & _
                                  ", dModified = " & p_oMod.dateParm(ServerDate)
                        
                        Execute lsSQL, "xxxSCA_Usage"
                     End If
                  End If
               Else
                  MsgBox "Invalid approval code detected!" & _
                         "You are not allowed to enter the system...", vbOKOnly + vbCritical, "Sytem Security"
                  GoTo endProc
               End If
            Else
               If .Fields("sApproved") = "" Then
                  MsgBox "Invalid cached approval code detected!" & _
                         "You are not allowed to enter the system...", vbOKOnly + vbCritical, "Sytem Security"
                  GoTo endProc
               Else
                  'If the date to be updated is the current date. just compare the approval code
                  'entered by the user with the approval code save in the system...
                  If .Fields("sApproved") <> lsApproveCD Then
                     MsgBox "Invalid approval code confirmation detected!" & _
                            "You are not allowed to enter the system...", vbOKOnly + vbCritical, "Sytem Security"
                     GoTo endProc
                  End If
               End If
            End If
         Else
            lsSQL = "UPDATE xxxSysApplication" & _
                   " SET dSysDatex = " & p_oMod.dateParm(p_dSysDatex) & _
                   " WHERE sClientID = " & p_oMod.strParm(p_sClientId) & _
                     " AND sProdctID = " & p_oMod.strParm(p_sProdctID)
            Call p_ocnMain.Execute(lsSQL, , adCmdText)
         End If
      Else
         MsgBox "The Day-To-Day transaction for this date was already posted!" & vbCrLf & _
                "Updates should be made through adjustment..." & vbCrLf & _
                "For more info please contact SEG/SSG!"
         GoTo endProc
      End If
   End With
      
   openDay2Day = True
   
endProc:
   Exit Function
errProc:
   GoTo endProc
End Function

Private Sub LoadOthers()
   Dim lsSQL As String
   lsSQL = "SELECT *" & _
          " FROM Branch_Others" & _
          " WHERE sBranchCD = " & p_oMod.strParm(p_sBranchCd)
   Set p_oRSConfg = p_ocnMain.Execute(lsSQL, , adCmdText)
End Sub

Private Function getShortAppCode( _
      ByVal fsBranchCD As String _
    , ByVal fdTransact As Date _
    , ByVal fnTimes As Integer) As String
   
   Dim lnCtr As Integer
   Dim lnBTotal As Integer
   Dim lnDTotal As Integer
   
   'getBranchTotal
   For lnCtr = 2 To 4
      lnBTotal = lnBTotal + Asc(Mid(fsBranchCD, lnCtr, 1))
   Next
   
   lnDTotal = Month(fdTransact) + Day(fdTransact)
   
   getShortAppCode = p_oMod.Dec2Hex(lnBTotal) & Format(lnDTotal, "00") & fnTimes
End Function

'kalyptus - 2018.04.03 02:50pm
'An encryption method that will make the encryption equivalent
'with the new encryption method in our java library
Public Function Encrypt(ByVal Code As String, Optional Signature As Variant) As String
   If p_nHexCrypt = 1 Then
      If IsMissing(Signature) Then
         Encrypt = p_oMod.StringToHex(p_oMod.Encrypt(Code))
      Else
         Encrypt = p_oMod.StringToHex(p_oMod.Encrypt(Code, Signature))
      End If
   Else
      If IsMissing(Signature) Then
         Encrypt = p_oMod.Encrypt(Code)
      Else
         Encrypt = p_oMod.Encrypt(Code, Signature)
      End If
   End If
End Function

'kalyptus - 2018.04.03 02:50pm
'A decryption method that will make the decryption equivalent
'with the new decryption method in our java library
Public Function Decrypt(ByVal Code As String, Optional Signature As Variant) As String
   If p_nHexCrypt = 1 Then
      If IsMissing(Signature) Then
         Decrypt = p_oMod.Decrypt(p_oMod.HexToString(Code))
      Else
         Decrypt = p_oMod.Decrypt(p_oMod.HexToString(Code), Signature)
      End If
   Else
      If IsMissing(Signature) Then
         Decrypt = p_oMod.Decrypt(Code)
      Else
         Decrypt = p_oMod.Decrypt(Code, Signature)
      End If
   End If
End Function

'kalyptus - 2013.12.26 12:50pm
'Day-To-Day approval code analysis
'
'AABBCDEE -> EEDCBBAA
'
'AA=Hexadecimal code of the minutes when the approval code was created
'BB=ADD month and day
'c = No of Times
'D = Random Number from 0-9
'EE=Add all ASCII VALUE of the numeric part of branch code of the requesting branch
'
'Example #1
'Branch:     M028
'Date       : 2013-12-26
'No of Times: 1
'Approval Code = 0F+38+1+5+9A
'
'Branch:     M028
'Date       : 2013-12-26
'No of Times: 2
'Approval Code = 3B+38+2+4+9A
'
'   MsgBox Mid("3B38249A", 7, 2) ' -> getBranch
'   MsgBox Mid("3B38249A", 3, 2) ' -> getDate
'   MsgBox Mid("3B38249A", 5, 1) ' -> getNoTimes

