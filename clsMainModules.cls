VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMainModules"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'-------------------------------------------------------------------------------------'
'  Execute
'-------------------------------------------------------------------------------------'
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Const INFINITE = &HFFFFFFFF       '  Infinite timeout
Private Const SYNCHRONIZE = &H100000
Private Const STILL_ACTIVE = 0
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
'-------------------------------------------------------------------------------------'
'  Execute
'-------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------'
'  Internet Connection
'-------------------------------------------------------------------------------------'
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
   (ByVal hInet As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, _
      ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long

Private Declare Function InternetCloseHandle Lib "wininet.dll" _
   (ByVal hInet As Long) As Long

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
   (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As _
      String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long

Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
'-------------------------------------------------------------------------------------'
'  Internet Connection
'-------------------------------------------------------------------------------------'

Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetCompName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetFocus Lib "user32.dll" () As Long

Public Type FieldProperty
   Name As String
   Value As Variant
   Type As String
   Size As Integer
   Help As String
   Format As String
   Reference As Boolean
   Initialize As Boolean
   Empty As Boolean
   Query As String
   QColumn As String
   QTitle As String
   QFormat As String
   QType As String
   QReference As String
   FieldType As Integer
End Type

Public Enum xeLogical
   xeUnknown = -1
   xeNo = 0
   xeYes = 1
End Enum

Public Enum xeTransactionStatus
   xeStateOpen = 0
   xeStateClosed = 1
   xeStatePosted = 2
   xeStateCancelled = 3
   xeStateUnknown = 4
End Enum

Public Enum xeRecordStatus
   xeRecStateInactive = 0
   xeRecStateActive = 1
   xeRecStateUnknown = 2
End Enum

Public Enum xeUserStatus
   xeUserStateSuspended = 0
   xeUserStateActive = 1
End Enum

Public Enum xeUserType
   xeUserTypeLocal = 0
   xeUserTypeGlobal = 1
End Enum

Public Enum xeLogStatus
   xeLogStateOut = 0
   xeLogStateIn = 1
End Enum

Public Enum xeLockStatus
   xeLockStateUnLock = 0
   xeLockStateLock = 1
End Enum

Public Enum xeEditMode
   xeModeUnknown = -1
   xeModeReady = 0
   xeModeAddNew = 1
   xeModeUpdate = 2
   xeModeDelete = 3
End Enum

Public Enum xeUserRights
   xeEncoder = 1
   xeSupervisor = 2
   xeManager = 4
   xeAudit = 8
   xeSysAdmin = 16
   xeSysOwner = 32
   xeEngineer = 64
   xeSysMaster = 128
End Enum

Public Enum xeFormType
   xeFormMaintenance = 0
   xeFormTransaction = 1
   xeFormQuickSearch = 2
   xeFormMessageBox = 3
   xeFormLedger = 4
   xeFormTransDetail = 5
   xeFormTransMaintenance = 6
   xeFormTransEqualLeft = 7
   xeFormTransEqualRight = 8
   xeFormUnknown = 9
End Enum

Public Enum MCLocation
   xeLocWarehouse = 0
   xeLocBranch = 1
   xeLocSupplier = 2
   xeLocCustomer = 3
   xeLocUnknown = 4
   xeLocServiceCenter = 5
   xeLocServicePhone = 6
End Enum

Public Enum ClientType
   xeClientCustomer = 0
   xeClientSupplier = 1
   xeClientTrucking = 2
End Enum

Public Enum AccoutStat
   xeActStatActive = 0
   xeActStatClosed = 1
   xeActStatDead = 2
   xeActStatImpounded = 3
   xeActStatDiscarded = 4
End Enum

Public Enum xeErrorType
   xeErr_Unknown = -1
   xeErr_Not_ADO = 0
   xeErr_Duplicate = 1
   xeErr_No_Update = 2
End Enum

Public Enum xeAccountType
   xeActTypActive = 0
   xeActTypProblem = 1
   xeActTypInactive = 2
   xeActTypDead = 3
End Enum

Public Enum xeMCStatus
   xeMCStatActive = 0
   xeMCStatImpounded = 1
   xeMCStatSurrendered = 2
   xeMCStatCarnap = 3
   xeMCStatFLA = 4
   xeMCStatMissingUnit = 5
   xeMCStat4Carnap = 7
   xeMCStat4FLA = 8
   xeMCStat4MissingUnit = 9
End Enum

Public Function strParm(ByVal sStrParm As String) As String
   sStrParm = Replace(sStrParm, Chr(34), Chr(34) & Chr(34))
   sStrParm = Replace(sStrParm, "\", "\\")
   strParm = "'" & Replace(sStrParm, "'", "''") & "'"
End Function

Public Function dateParm(ByVal dDateParm As Date) As String
   dateParm = "'" & Format(dDateParm, "yyyy-mm-dd hh:mm:ss") & "'"
End Function

Public Function KwikSearch(oAppDriver As clsAppDriver, _
                           ByVal sSource As String, _
                           Optional sFieldName As Variant, _
                           Optional sFieldDesc As Variant, _
                           Optional sFieldPict As Variant, _
                           Optional sFieldCrit As Variant) As Variant
   Dim loLookUp As frmLookUp1
   Dim lsProcName As String
   
   lsProcName = "KwikSearch"
   On Error GoTo errProc
   
   Set loLookUp = New frmLookUp1
   Set loLookUp.AppDriver = oAppDriver
   
   loLookUp.SQLSource = sSource
   If Not IsMissing(sFieldName) Then
      If sFieldName <> "" Then loLookUp.FldName = sFieldName
   End If
   If Not IsMissing(sFieldDesc) Then
      If sFieldDesc <> "" Then loLookUp.FldTitle = sFieldDesc
   End If
   If Not IsMissing(sFieldPict) Then
      If sFieldPict <> "" Then loLookUp.FldFormat = sFieldPict
   End If
   If Not IsMissing(sFieldCrit) Then
      If sFieldCrit <> "" Then loLookUp.FldCriteria = sFieldCrit
   End If

   loLookUp.Show 1
   KwikSearch = loLookUp.SelectedItem
   
endProc:
   Set loLookUp = Nothing
   Exit Function
errProc:
   ShowError oAppDriver _
           , lsProcName & "( " & "oAppDriver" _
                       & ", " & sSource _
                       & ", " & sFieldName _
                       & ", " & sFieldDesc _
                       & ", " & sFieldPict _
                       & ", " & sFieldCrit & " )"
End Function

'      lsSQL = KwikBrowse(p_oAppDrivr, lors, _
'                                       "sTransNox»dTransact»xEmployee", _
'                                       "Trans No»Date»Employee", _
'                                       "@@@@-@@@@@@@@»@»@")

Public Function KwikBrowse(oAppDriver As clsAppDriver, _
                           oSource As ADODB.Recordset, _
                           Optional sFieldName As Variant, _
                           Optional sFieldDesc As Variant, _
                           Optional sFieldPict As Variant, _
                           Optional sFieldCrit As Variant) As Variant
   
   Dim loLookUp As frmLookUp
   Dim lsProcName As String
   
   lsProcName = "KwikBrowse"
   On Error GoTo errProc
   
   Set loLookUp = New frmLookUp
   Set loLookUp.AppDriver = oAppDriver
   
   Set loLookUp.RowSource = oSource
   If Not IsMissing(sFieldName) Then
      If sFieldName <> "" Then loLookUp.FldName = sFieldName
   End If
   If Not IsMissing(sFieldDesc) Then
      If sFieldDesc <> "" Then loLookUp.FldTitle = sFieldDesc
   End If
   If Not IsMissing(sFieldPict) Then
      If sFieldPict <> "" Then loLookUp.FldFormat = sFieldPict
   End If
   If Not IsMissing(sFieldCrit) Then
      If sFieldCrit <> "" Then loLookUp.FldCriteria = sFieldCrit
   End If

   loLookUp.Show 1
   KwikBrowse = loLookUp.SelectedItem
   
endProc:
   Set loLookUp = Nothing
   Exit Function
errProc:
   ShowError oAppDriver _
           , lsProcName & "( " & "oAppDriver" _
                       & ", " & "oSource" _
                       & ", " & sFieldName _
                       & ", " & sFieldDesc _
                       & ", " & sFieldPict _
                       & ", " & sFieldCrit & " )"
End Function

Public Function isAppDriverOK(oAppDriver) As Boolean
   isAppDriverOK = True
   If oAppDriver Is Nothing Then
      MsgBox "Application is not yet set!!!" & vbCrLf & _
               "Can not further proceed!!!", vbCritical, "Warning"
      isAppDriverOK = False
   End If
End Function

Public Function GetApproval(oAppDriver As clsAppDriver, _
                            nApproveRights As Integer, _
                            sApproveID As String, _
                            sApproveName As String, _
                            sMenuID As String) As Boolean
   
   Dim loForm As frmApproval
   Dim lors As ADODB.Recordset
   'Dim loCrypt As clsCrypto
   Dim lsUserName As String
   Dim lsPassword As String
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lbLogIn As Boolean
   Dim lnRights As Integer
   
   lsProcName = "GetApproval"
   On Error GoTo errProc
   GetApproval = False
   
   Set lors = New Recordset
   'Set loCrypt = New clsCrypto
   'loCrypt.Signature = oAppDriver.Machinex
   
   Set loForm = New frmApproval
   
   lnCtr = 0
   lbLogIn = False
   Do
      loForm.AppPath = oAppDriver.AppPath
      loForm.Show 1
      If loForm.Cancel = True Then
         ' user either clicked the cancel button or the close key
         lnCtr = 2
      Else
         ' Encrypt the username and password
         'loCrypt.InBuffer = loForm.txtPassword
         'loCrypt.Encrypt
         'lsPassword = loCrypt.OutBuffer
         lsPassword = oAppDriver.Encrypt(loForm.txtPassword, oAppDriver.Machinex)
             
         'loCrypt.InBuffer = LCase(loForm.txtUserName)
         'loCrypt.Encrypt
         'lsUserName = loCrypt.OutBuffer
         lsUserName = oAppDriver.Encrypt(LCase(loForm.txtUserName), oAppDriver.Machinex)
         
         ' search for the user
'         lsSQL = "SELECT" & _
'                     "  a.sUserIDxx" & _
'                     ", a.sUserName" & _
'                     ", a.nUserLevl" & _
'                     ", a.cLockStat" & _
'                     ", a.cUserStat" & _
'                     ", a.cUserType" & _
'                     ", b.nAppRight" & _
'                     ", a.sProdctID" & _
'                     ", c.nAppRight xAppRight" & _
'                  " FROM xxxSysUser a" & _
'                        " LEFT JOIN xxxMenuObject b" & _
'                           " ON b.sMenuIDxx = " & strParm(sMenuID) & _
'                        " LEFT JOIN xxxMenuRightsObject c" & _
'                           " ON a.sUserIDxx = c.sUserIDxx" & _
'                              " AND c.sMenuIDxx = " & strParm(sMenuID) & _
'                  " WHERE sLogNamex = " & strParm(lsUserName) & _
'                     " AND sPassword = " & strParm(lsPassword) & _
'                     " AND ( cUserType = '1' OR a.sProdctID = " & strParm(oAppDriver.ProductID) & " )"
         
         lsSQL = "SELECT" & _
                     "  a.sUserIDxx" & _
                     ", a.sUserName" & _
                     ", a.nUserLevl" & _
                     ", a.cLockStat" & _
                     ", a.cUserStat" & _
                     ", a.cUserType" & _
                     ", b.nAppRight" & _
                     ", a.sProdctID" & _
                     ", c.nAppRight xAppRight" & _
                  " FROM xxxSysUser a" & _
                        " LEFT JOIN xxxMenuObject b" & _
                           " ON b.sMenuIDxx = " & strParm(sMenuID) & _
                        " LEFT JOIN xxxMenuRightsObject c" & _
                           " ON a.sUserIDxx = c.sUserIDxx" & _
                              " AND b.sMenuIDxx = c.sMenuIDxx" & _
                  " WHERE sLogNamex = " & strParm(lsUserName) & _
                     " AND sPassword = " & strParm(lsPassword) & _
                     " AND ( cUserType = '1' OR a.sProdctID = " & strParm(oAppDriver.ProductID) & " )"
         
'         If lors.State = adStateOpen Then lors.Close
         Set lors = New ADODB.Recordset
         lors.Open lsSQL, oAppDriver.Connection, , , adCmdText
         Debug.Print lsSQL
         If lors.EOF Then
            MsgBox "Verify your log name and/or password", vbCritical, "Login Error"
         Else
            ' default approval rights
            lnRights = 230
            
            ' check first if user has the right to approved procedure entry
'            If IsNull(lors("nAppRight")) = False Then
'               If IsNull(lors("xAppRight")) = False Then
'                  lnRights = lors("xAppRight")
'               Else
'                  lnRights = lors("xAppRight")
'               End If
'            End If
            
            If IsNull(lors("nAppRight")) = False Then
               lnRights = lors("nAppRight")
            End If
            
            If (lors("nUserLevl") And lnRights) = 0 Then
               MsgBox "User has no right for this procedure!!!" & vbCrLf & vbCrLf & _
                        "Seek other user's assistance then try again!!!", _
                        vbCritical, "Warning"
            Else
               lbLogIn = True
               Debug.Print lsSQL
            End If
         End If
      End If
      lnCtr = lnCtr + 1
   Loop Until (lbLogIn = True) Or (lnCtr = 3)
   
   If lbLogIn = False Then GoTo endProc
      
   ' check first if user is not suspended
   If lors("cUserStat") = xeUserStateSuspended Then
      MsgBox "User is currently Suspended!!!" & vbCrLf & _
               "User has no Rights for Procedure Approval!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   ' then check if its lock
   If lors("cLockStat") = xeYes Then
      MsgBox "User Rights is Currently Locked!!!" & vbCrLf & _
               "User has no Rights for Procedure Approval!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   ' finally check if user is within the group
   If lors("cUserType") = xeUserTypeLocal Then
      If Trim(lors("sProdctID")) <> Trim(oAppDriver.ProductID) Then
         MsgBox "User is not a Member of this Application!!!" & vbCrLf & _
                  "User has no Rights for Procedure Approval!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End If
   
   nApproveRights = lors("nUserLevl")
   sApproveID = lors("sUserIDxx")
   sApproveName = lors("sUserName")
   
   ' decrypt User Name
   'loCrypt.InBuffer = sApproveName
   'loCrypt.Decrypt
   'sApproveName = loCrypt.OutBuffer
   sApproveName = oAppDriver.Decrypt(sApproveName, oAppDriver.Machinex)
   
   GetApproval = True
   
endProc:
   Unload loForm
   Set lors = Nothing
'   Set loCrypt = Nothing
   Set loForm = Nothing
   
   Exit Function
errProc:
   ShowError oAppDriver _
           , lsProcName & "( " & "oAppDriver" _
                       & ", " & nApproveRights _
                       & ", " & sApproveID _
                       & ", " & sApproveName _
                       & ", " & sMenuID & " )"
End Function

Public Function GetCodeApproval(oAppDriver As clsAppDriver, _
                            sApprovalCde As String, _
                            sApproveID As String, _
                            sApproveName As String) As Boolean


   Dim loForm As frmCodeApproval
   Dim lors As ADODB.Recordset
   Dim loCrypt As clsCrypto
   Dim lsProcName As String
   Dim lsClientID As String
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lbLogIn As Boolean
   
   lsProcName = "GetDayApproval"
   On Error GoTo errProc
   GetCodeApproval = False
   
   Set lors = New Recordset
   Set loForm = New frmCodeApproval
   
   lnCtr = 0
   lbLogIn = False
   Do
      Set loForm.AppDriver = oAppDriver
      loForm.AppPath = oAppDriver.AppPath
      loForm.Show 1
      If loForm.Cancel = True Then
         ' user either clicked the cancel button or the close key
         lnCtr = 2
      Else
         If Mid(UCase(loForm.CodeApproval), 4, 1) <> loForm.IssueeType Then
            MsgBox "Invalid Approval Parameter Detected." & vbCrLf & vbCrLf & _
                     "Kindly verify your entry.", vbCritical, "Approval Error"
         Else
            lbLogIn = True
         End If
      End If
      lnCtr = lnCtr + 1
   Loop Until (lbLogIn = True) Or (lnCtr = 3)
   
   If lbLogIn = False Then GoTo endProc
        
   sApprovalCde = UCase(loForm.CodeApproval)
   sApproveID = loForm.UserID
   sApproveName = loForm.Issuee
   
   GetCodeApproval = True
   
endProc:
   Unload loForm
   Set lors = Nothing
   Set loForm = Nothing
   
   Exit Function
errProc:
   ShowError oAppDriver _
           , lsProcName & "( " & "oAppDriver" _
                       & ", " & sApprovalCde _
                       & ", " & sApproveID _
                       & ", " & sApproveName & ")"
End Function

Public Function GetDayApproval(oAppDriver As clsAppDriver, _
                            sApprovalCde As String, _
                            sApproveID As String, _
                            sApproveName As String) As Boolean
End Function

Public Function GetUserConfirmation(oAppDriver As clsAppDriver, _
                                    sUserID As String, _
                                    sUserName As String) As Boolean
   
   Dim lors As ADODB.Recordset
   Dim loForm As frmApproval
   'Dim loCrypt As clsCrypto
   Dim loMod As clsMainModules
   Dim lsUser As String
   Dim lsPassword As String
   Dim lsSQL As String
   Dim lnCtr As Integer
   Dim lbLogIn As Boolean
   Dim lsProcName As String
   
   lsProcName = "GetUserConfirmation"
   On Error GoTo errProc
   GetUserConfirmation = False
      
   Set lors = New Recordset
   'Set loCrypt = New clsCrypto
   'loCrypt.Signature = oAppDriver.Machinex
   
   Set loForm = New frmApproval
   Set loMod = New clsMainModules
   
   lnCtr = 0
   lbLogIn = False
   Do
      loForm.AppPath = oAppDriver.AppPath
      loForm.Show 1
      If loForm.Cancel = True Then
         ' user either clicked the cancel button or the close key
         lnCtr = 2
      Else
         ' Encrypt the username and password
         'loCrypt.InBuffer = loForm.txtPassword
         'loCrypt.Encrypt
         'lsPassword = loCrypt.OutBuffer
         lsPassword = oAppDriver.Encrypt(loForm.txtPassword, oAppDriver.Machinex)
        
         'loCrypt.InBuffer = LCase(loForm.txtUserName)
         'loCrypt.Encrypt
         'lsUser = loCrypt.OutBuffer
         lsUser = oAppDriver.Encrypt(LCase(loForm.txtUserName), oAppDriver.Machinex)
              
         With loMod
            ' search for the user
            lsSQL = "SELECT *" & _
                     " FROM xxxSysUser" & _
                     " WHERE sLogNamex = " & .strParm(lsUser) & _
                        " AND sPassword = " & .strParm(lsPassword) & _
                        " AND ( cUserType = '1' OR sProdctID = " & .strParm(oAppDriver.ProductID) & " )"
         End With
         
         If lors.State = adStateOpen Then lors.Close
         
         lors.Open lsSQL, oAppDriver.Connection, , , adCmdText
         If lors.EOF Then
            MsgBox "Verify your log name and/or password", vbCritical, "Login Error"
         Else
            lbLogIn = True
         End If
      End If
      lnCtr = lnCtr + 1
   Loop Until (lbLogIn = True) Or (lnCtr = 3)
   
   If lbLogIn = False Then GoTo endProc
      
   ' check first if user is not suspended
   If lors("cUserStat") = xeUserStateSuspended Then
      MsgBox "User is currently Suspended!!!" & vbCrLf & _
               "User has no Rights for Procedure Approval!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If lors("cUserType") = xeUserTypeLocal Then
      If Trim(lors("sProdctID")) <> Trim(oAppDriver.ProductID) Then
         MsgBox "User is not a Member of this Application!!!" & vbCrLf & _
                  "User has no Rights for Procedure Approval!!!", vbCritical, "Warning"
         GoTo endProc
      End If
   End If
   
   sUserID = lors("sUserIDxx")
   sUserName = lors("sUserName")
   'loCrypt.InBuffer = sUserName
   'loCrypt.Decrypt
   'sUserName = loCrypt.OutBuffer
   sUserName = oAppDriver.Decrypt(sUserName, oAppDriver.Machinex)
   
   GetUserConfirmation = True
   
endProc:
   Unload loForm
   Set lors = Nothing
'   Set loCrypt = Nothing
   Set loForm = Nothing
   Set loMod = Nothing
   
   Exit Function
errProc:
   ShowError oAppDriver _
           , lsProcName & "( " & "oAppDriver" _
                       & ", " & sUserID _
                       & ", " & sUserName & " )"
End Function

Public Function GetNextCode(ByVal Table As String, _
                            ByVal Field As String, _
                            ByVal YearFormat As Boolean, _
                            Connection As Connection, _
                            Optional ByBranch As Boolean = False, _
                            Optional Branch As String = "") As String
   Dim lors As ADODB.Recordset
   Dim lsField As String
   Dim lsSQL As String
   Dim lnCounter As Integer
   Dim lnCode As Long
   Dim lnLen As Long
   Dim ldDate As Date
      
   'INSERT THIS TO MAKE BRANCH 01 THE DEFAULT BRANCH
   'Branch = "M001"
      
   If ByBranch = True And Branch = "" Then Exit Function
   
   Set lors = New Recordset
   
   lsField = IIf(ByBranch, Branch, "")
   If YearFormat Then
'      lors.Open "SELECT SYSDATE()", Connection, , , adCmdText
      lors.Open "SELECT NOW()", Connection, , , adCmdText
      If lors.EOF Then Exit Function
      
      ldDate = lors(0)
      lsField = lsField & Format(lors(0), "YY")
      lors.Close
      lnCounter = Len(lsField)
   Else
      lnCounter = Len(lsField)
   End If
   
   lsSQL = "SELECT " & Field & _
            " FROM " & Table & _
            " WHERE " & Field & " LIKE " & strParm(lsField & "%") & _
            " ORDER BY " & Field & " DESC LIMIT 1"
   lors.Open lsSQL, Connection, , , adCmdText
   
   lsSQL = IIf(lors.EOF, Empty, lors(0))

   lnLen = lors(0).DefinedSize
   lnCode = 1
   If lsSQL <> Empty Then
      lnCode = CLng(Mid(lsSQL, lnCounter + 1)) + 1
   Else
      lsSQL = IIf(ByBranch, Branch, "")
      If YearFormat Then
         lsSQL = lsSQL & Format(ldDate, "YY")
         lnCounter = Len(Branch) + 2
      Else
         lnCounter = Len(lsSQL)
      End If
   End If

   If lsSQL = Empty Then
      GetNextCode = Format(lnCode, String(lnCounter, "0"))
   Else
      GetNextCode = Left(lsSQL, lnCounter) & Format(lnCode, String(lnLen - lnCounter, "0"))
   End If
End Function

Public Function GetNextReference(ByVal Table As String, _
                                 ByVal Field As String, _
                                 ByVal Order As String, _
                                 ByVal BranchField As String, _
                                 ByVal Branch As String, _
                                 Connection As Connection) As String
   Dim lors As ADODB.Recordset
   Dim lsSQL As String
   Dim lnCode As Double
   Dim lnLen As Long
   Dim lsProcName As String

   lsProcName = "GetNextReference"
   On Error GoTo errProc
   
   lsSQL = "SELECT " & Field & _
            " FROM " & Table & _
            " WHERE " & BranchField & " LIKE " & strParm(Branch & "%")
   lsSQL = lsSQL & " ORDER BY " & Order & " DESC" & _
         ", " & Field & " DESC LIMIT 1"
   Set lors = New Recordset
   lors.Open lsSQL, Connection, , , adCmdText
         Debug.Print lsSQL
   If lors.EOF Then
      lsSQL = 0
      lnLen = lors(0).DefinedSize
   Else
      lsSQL = IIf(lors(0) = "", 0, lors(0))
      lnLen = lors(0).ActualSize
   End If
   lnCode = CDbl(lsSQL) + 1

   GetNextReference = Format(lnCode, String(lnLen, "0"))

endProc:
   Exit Function
errProc:
   MsgBox "Error in " & lsProcName & "( " & Table _
                       & ", " & Field _
                       & ", " & Order _
                       & ", " & BranchField _
                       & ", " & Branch _
                       & ", " & "Connection" & " )"
End Function

' transalate the query to retrieve an empty record
Public Function getEmptyRec(ByVal sSQL As String) As String
   Dim lnPosition As Integer
   
   ' parse first the Order by clause
   lnPosition = InStr(1, sSQL, "ORDER BY", vbTextCompare)
   If lnPosition > 0 Then sSQL = Trim(Left(sSQL, lnPosition - 1))
      
   ' parse the where clause too
   lnPosition = InStr(1, sSQL, "WHERE", vbTextCompare)
   If lnPosition > 0 Then sSQL = Trim(Left(sSQL, lnPosition - 1))
   
   ' the goal is to retrieve an empty record, so just create the simpliest query to do it
   getEmptyRec = sSQL & " WHERE 0 = 1"
End Function

' adds the condition passed to a query statement
Public Function AddCondition(ByVal sSQL As String, ByVal sCondition As String) As String
   Dim lnPos As Integer
   Dim lsOrder As String
   
   Dim lnUPos As Integer
   Dim lsUnion As String
   Dim lsTemp As String
   
   Dim lsSubQueryBeg As String, lsSubQueryEnd As String
   
   ' check first if sCondition contains a valid value
   If sCondition = Empty Then
      AddCondition = sSQL
      Exit Function
   End If

   ' parse first the subquery
   lnPos = InStr(1, sSQL, "FROM (", vbTextCompare)
   If lnPos > 0 Then
      lsSubQueryBeg = Trim(Left(sSQL, lnPos + Len("FROM (") - 1))
      sSQL = Trim(Mid(sSQL, lnPos + Len("FROM (")))
      
      lnPos = InStrRev(sSQL, ")", -1, vbTextCompare)
      If lnPos = 0 Then
         AddCondition = sSQL
         Exit Function
      End If
      
      lsSubQueryEnd = Mid(sSQL, lnPos)
      sSQL = Mid(sSQL, 1, lnPos - 1)
   End If
   
   ' parse first the group by clause
   lnPos = InStr(1, sSQL, "GROUP BY", vbTextCompare)
   If lnPos > 0 Then
      lsOrder = Trim(Mid(sSQL, lnPos))
      sSQL = Trim(Left(sSQL, lnPos - 1))
   Else
      lnPos = InStr(1, sSQL, "HAVING", vbTextCompare)
      If lnPos > 0 Then
         lsOrder = Trim(Mid(sSQL, lnPos))
         sSQL = Trim(Left(sSQL, lnPos - 1))
      Else
         lnPos = InStr(1, sSQL, "ORDER BY", vbTextCompare)
         If lnPos > 0 Then
            lsOrder = Trim(Mid(sSQL, lnPos))
            sSQL = Trim(Left(sSQL, lnPos - 1))
         End If
         
         lnUPos = InStr(1, sSQL, "UNION", vbTextCompare)
         If lnUPos > 0 Then
            lsUnion = Trim(Mid(sSQL, lnUPos + Len("UNION")))
            sSQL = Trim(Left(sSQL, lnUPos - 1))
         End If
      End If
   End If
         
   lnPos = InStr(1, sSQL, "WHERE", vbTextCompare)
   If lnPos > 0 Then
      sSQL = Trim(sSQL) & " AND " & sCondition
   Else
      sSQL = Trim(sSQL) & " WHERE " & sCondition
   End If
   
   Do Until lnUPos = 0
      lnUPos = InStr(1, lsUnion, "UNION", vbTextCompare)
      If lnUPos > 0 Then
         lsTemp = Trim(Mid(lsUnion, lnUPos + Len("UNION")))
         lsUnion = Trim(Left(lsUnion, lnUPos - 1))
      End If
   
      lnPos = InStr(1, lsUnion, "WHERE", vbTextCompare)
      If lnPos > 0 Then
         lsUnion = Trim(lsUnion) & " AND " & sCondition
      Else
         lsUnion = Trim(lsUnion) & " WHERE " & sCondition
      End If
      
      sSQL = sSQL & " UNION " & lsUnion
      lsUnion = lsTemp
   Loop
   
   AddCondition = Trim(lsSubQueryBeg & " " & Trim(sSQL) & " " & lsOrder & " " & lsSubQueryEnd)
End Function

Public Sub SetNextFocus()
   keybd_event &H9, 0, 0, 0
   keybd_event &H9, 0, &H2, 0
End Sub

Public Sub SetPreviousFocus()
   keybd_event &H10, 0, 0, 0
   keybd_event &H9, 0, 0, 0
   keybd_event &H10, 0, &H2, 0
End Sub

Public Sub CenterChildForm(oMDIForm As Object, ByRef oChild As Object)
   Dim lnX, lnY As Long
   
   lnX = oMDIForm.ScaleWidth
   lnY = oMDIForm.ScaleHeight
   
   oChild.Left = CLng((lnX - oChild.Width) / 2)
   oChild.Top = CLng((lnY - oChild.Height) / 2)
   
   If oChild.Left < 0 Then oChild.Left = 0
   If oChild.Top < 0 Then oChild.Top = 0
End Sub

Public Function Encrypt(ByVal Code As String, Optional Signature As Variant) As String
   Dim loCrypt As clsCrypto
   
   Set loCrypt = New clsCrypto
   
   If Not IsMissing(Signature) Then loCrypt.Signature = Signature
   loCrypt.InBuffer = Code
   loCrypt.Encrypt
   
   Encrypt = loCrypt.OutBuffer
   
   Set loCrypt = Nothing
End Function

Public Function Decrypt(ByVal Code As String, Optional Signature As Variant) As String
   Dim loCrypt As clsCrypto
   
   Set loCrypt = New clsCrypto
   
   If Not IsMissing(Signature) Then loCrypt.Signature = Signature
   
   loCrypt.InBuffer = Code
   loCrypt.Decrypt
   Decrypt = loCrypt.OutBuffer
   Set loCrypt = Nothing
End Function

Public Function NumToText(ByVal dblVal As Double) As String
    Static sOnes(0 To 9) As String
    Static sTeens(0 To 9) As String
    Static sTens(0 To 9) As String
    Static sThousands(0 To 4) As String
    Static bInit As Boolean
    
    Dim I As Integer
    Dim bAllZeros As Boolean
    Dim bShowsThousands As Boolean
    Dim sValue As String
    Dim sBuffer As String
    Dim sTemp As String
    Dim iCol As Integer
    Dim iChar As Integer
'
' Warning - this routine only handles positive values
    
    If bInit = False Then
        'Initialize array
        bInit = True
        sOnes(0) = "Zero"
        sOnes(1) = "One"
        sOnes(2) = "Two"
        sOnes(3) = "Three"
        sOnes(4) = "Four"
        sOnes(5) = "Five"
        sOnes(6) = "Six"
        sOnes(7) = "Seven"
        sOnes(8) = "Eight"
        sOnes(9) = "Nine"
        sTeens(0) = "Ten"
        sTeens(1) = "Eleven"
        sTeens(2) = "Twelve"
        sTeens(3) = "Thirteen"
        sTeens(4) = "Fourteen"
        sTeens(5) = "Fifteen"
        sTeens(6) = "Sixteen"
        sTeens(7) = "Seventeen"
        sTeens(8) = "Eighteen"
        sTeens(9) = "Nineteen"
        sTens(0) = ""
        sTens(1) = "Ten"
        sTens(2) = "Twenty"
        sTens(3) = "Thirty"
        sTens(4) = "Forty"
        sTens(5) = "Fifty"
        sTens(6) = "Sixty"
        sTens(7) = "Seventy"
        sTens(8) = "Eighty"
        sTens(9) = "Ninety"
        sThousands(0) = ""
        sThousands(1) = "Thousand"   'US numbering
        sThousands(2) = "Million"
        sThousands(3) = "Billion"
        sThousands(4) = "Trillion"
    End If
    '
    ' Setup Error Handler
    '
    On Error GoTo vbErrorHandler
    '
    ' Get fractional part of value (if any)
    '
    
    sValue = Format$((dblVal - Int(dblVal)) * 100, "00")
      
    If Val(sValue) > 0 Then
      For I = Len(sValue) To 1 Step -1
          iChar = Val(Mid$(sValue, I, 1))
          iCol = (Len(sValue) - I) + 1
          '
          'Action depends on 1's, 10's or 100's column
          '
          Select Case (iCol Mod 3)
              Case 1  '1's position
                  bShowsThousands = True
                  If I = 1 Then
                      sTemp = sOnes(iChar) & " "
                  ElseIf Mid$(sValue, I - 1, 1) = "1" Then
                      sTemp = sTeens(iChar) & " "
                      I = I - 1
                  ElseIf iChar > 0 Then
                      sTemp = sOnes(iChar) & " "
                  Else
                      bShowsThousands = False
                      If Mid$(sValue, I - 1, 1) <> "0" Then
                          bShowsThousands = True
                      ElseIf I > 2 Then
                          If Mid$(sValue, I - 2, 1) <> "0" Then
                              bShowsThousands = True
                          End If
                      End If
                      sTemp = ""
                  End If
                  If bShowsThousands Then
                      If iCol > 1 Then
                          sTemp = sTemp & sThousands(iCol \ 3)
                          If bAllZeros Then
                              sTemp = sTemp & " "
                          Else
                              sTemp = sTemp & " "
                          End If
                      End If
                      bAllZeros = False
                  End If
                  sBuffer = sTemp & sBuffer
              Case 2
                  If iChar > 0 Then
                      If Mid$(sValue, I + 1, 1) <> "0" Then
                          sBuffer = sTens(iChar) & " " & sBuffer
                      Else
                          sBuffer = sTens(iChar) & " " & sBuffer
                      End If
                  End If
              Case 0
                  If iChar > 0 Then
                      sBuffer = sOnes(iChar) & " Hundred " & sBuffer
                  End If
          End Select
      Next I
      sBuffer = "Pesos And " & sBuffer & "Centavo(s) Only"
    Else
      sBuffer = "Pesos Only"
    End If
    '
    ' Convert main part to string
    '
    sValue = CStr(Int(dblVal))
    
    bAllZeros = True
    
    For I = Len(sValue) To 1 Step -1
        iChar = Val(Mid$(sValue, I, 1))
        iCol = (Len(sValue) - I) + 1
        '
        'Action depends on 1's, 10's or 100's column
        '
        Select Case (iCol Mod 3)
            Case 1  '1's position
                bShowsThousands = True
                If I = 1 Then
                    sTemp = sOnes(iChar) & " "
                ElseIf Mid$(sValue, I - 1, 1) = "1" Then
                    sTemp = sTeens(iChar) & " "
                    I = I - 1
                ElseIf iChar > 0 Then
                    sTemp = sOnes(iChar) & " "
                Else
                    bShowsThousands = False
                    If Mid$(sValue, I - 1, 1) <> "0" Then
                        bShowsThousands = True
                    ElseIf I > 2 Then
                        If Mid$(sValue, I - 2, 1) <> "0" Then
                            bShowsThousands = True
                        End If
                    End If
                    sTemp = ""
                End If
                If bShowsThousands Then
                    If iCol > 1 Then
                        sTemp = sTemp & sThousands(iCol \ 3)
                        If bAllZeros Then
                            sTemp = sTemp & " "
                        Else
                            sTemp = sTemp & " "
                        End If
                    End If
                    bAllZeros = False
                End If
                sBuffer = sTemp & sBuffer
            Case 2
                If iChar > 0 Then
                    If Mid$(sValue, I + 1, 1) <> "0" Then
                        sBuffer = sTens(iChar) & " " & sBuffer
                    Else
                        sBuffer = sTens(iChar) & " " & sBuffer
                    End If
                End If
            Case 0
                If iChar > 0 Then
                    sBuffer = sOnes(iChar) & " Hundred " & sBuffer
                End If
        End Select
    Next I
  
    sBuffer = UCase$(Left$(sBuffer, 1)) & Mid$(sBuffer, 2)

EndNumToText:
    NumToText = sBuffer
    Exit Function

vbErrorHandler:
    sBuffer = "#Error#"
    Resume EndNumToText
End Function

Public Function Price2Code(nPrice As Double, sCode As String) As String
   Dim lacDigitCode(9) As String * 1
   Dim lsPrice, lsConPrice As String
   Dim lnCtr As Long
   
   lsPrice = Trim(Str(nPrice))
   If Len(sCode) < 10 Then
      If Len(sCode) <> 9 Then
         Price2Code = "XXX"
         Exit Function
      End If
      sCode = sCode & "Z"
   End If
   
   lacDigitCode(0) = Mid(sCode, 10, 1)
   For lnCtr = 1 To UBound(lacDigitCode)
      lacDigitCode(lnCtr) = Mid(sCode, lnCtr, 1)
   Next
   
   lsConPrice = ""
   For lnCtr = 1 To Len(lsPrice)
      If Mid(lsPrice, lnCtr, 1) = "." Then
         lsConPrice = lsConPrice & "."
      Else
         lsConPrice = lsConPrice & lacDigitCode(Int(Mid(lsPrice, lnCtr, 1)))
      End If
   Next
   If InStr(1, lsConPrice, ".", vbTextCompare) = 0 Then
      lsConPrice = lsConPrice & "." & lacDigitCode(0) & lacDigitCode(0)
   End If
   Price2Code = lsConPrice
End Function

Public Function Code2Price(ByVal sPrice, ByVal sCode) As Double
   Dim lsConPrice As String
   Dim lnCtr As Integer
   Dim lnValue As Integer

   sPrice = Trim(sPrice)
   If Len(sCode) < 10 Then
      If Len(sCode) <> 9 Then
         Code2Price = 0#
         Exit Function
      End If
      sCode = sCode & "Z"
   End If

   lsConPrice = ""
   For lnCtr = 1 To Len(sPrice)
      If Mid(sPrice, lnCtr, 1) = "." Then
         lsConPrice = lsConPrice & "."
      Else
         lnValue = InStr(1, sCode, Mid(sPrice, lnCtr, 1), vbTextCompare)
         If lnValue = 10 Then lnValue = 0

         lsConPrice = lsConPrice & Trim(Str(lnValue))
      End If
   Next
   If InStr(1, lsConPrice, ".", vbTextCompare) = 0 Then
      lsConPrice = lsConPrice & ".00"
   End If
   Code2Price = CDbl(lsConPrice)
End Function

Public Function TitleCase(sValue As String) As String
   Dim lsString() As String
   Dim lnCtr As Integer
   Dim lsValue As String
   
   Dim lsString2() As String
   Dim lnCtr2 As Integer
   Dim lsValue2 As String
   
   On Error GoTo errProc
   
   lsString = Split(sValue, vbCrLf)
   lsValue = Empty
   
   If UBound(lsString) > 0 Then
      For lnCtr = 0 To UBound(lsString)
         Select Case lnCtr
         Case 0
            lsString2 = Split(Trim(lsString(lnCtr)), " ")
            lsValue2 = Empty
         
            If UBound(lsString2) > 0 Then
               For lnCtr2 = 0 To UBound(lsString2)
                  If Not IsNumeric(lsString2(lnCtr2)) Then
                     If Len(lsString2(lnCtr2)) = 1 Then
                        lsValue2 = Trim(lsValue2) & " " & UCase(lsString2(lnCtr2))
                     Else
                        If Trim(lsString2(lnCtr2)) <> "" Then
                           lsValue2 = Trim(lsValue2) & " " & UCase(Left(lsString2(lnCtr2), 1)) & Right(lsString2(lnCtr2), Len(lsString2(lnCtr2)) - 1)
                        End If
                     End If
                  Else
                     lsValue2 = lsValue2 & " " & lsString2(lnCtr2)
                  End If
               Next
         
               lsValue = Trim(lsValue2)
            Else
               If Trim(lsString(lnCtr)) <> Empty Then
                  If Not IsNumeric(lsString(lnCtr)) Then
                     If Len(lsString(lnCtr)) = 1 Then
                        lsValue = UCase(lsString(lnCtr))
                     Else
                        If Trim(lsString(lnCtr2)) <> "" Then
                           lsValue = Trim(UCase(Left(lsString(lnCtr), 1)) & Right(lsString(lnCtr), Len(lsString(lnCtr)) - 1))
                        End If
                     End If
                  Else
                     lsValue = Trim(lsString(lnCtr))
                  End If
               End If
            End If
         Case Else
            lsString2 = Split(Trim(lsString(lnCtr)), " ")
            lsValue2 = Empty
         
            If UBound(lsString2) > 0 Then
               For lnCtr2 = 0 To UBound(lsString2)
                  If Not IsNumeric(lsString2(lnCtr2)) Then
                     If Len(lsString2(lnCtr2)) = 1 Then
                        lsValue2 = Trim(lsValue2) & " " & UCase(lsString2(lnCtr2))
                     Else
                        If Trim(lsString2(lnCtr2)) <> "" Then
                           lsValue2 = Trim(lsValue2) & " " & UCase(Left(lsString2(lnCtr2), 1)) & Right(lsString2(lnCtr2), Len(lsString2(lnCtr2)) - 1)
                        End If
                     End If
                  Else
                     lsValue2 = lsValue2 & " " & lsString2(lnCtr2)
                  End If
               Next
         
               lsValue2 = Trim(lsValue2)
            Else
               If Trim(lsString(lnCtr)) <> Empty Then
                  If Not IsNumeric(lsString(lnCtr)) Then
                     If Len(lsString(lnCtr)) = 1 Then
                        lsValue2 = UCase(lsString(lnCtr))
                     Else
                        If Trim(lsString2(lnCtr2)) <> "" Then
                           lsValue2 = Trim(UCase(Left(lsString(lnCtr), 1)) & Right(lsString(lnCtr), Len(lsString(lnCtr)) - 1))
                        End If
                     End If
                  Else
                     lsValue2 = Trim(lsString(lnCtr))
                  End If
               End If
            End If
         
            lsValue = lsValue & vbCrLf & lsValue2
         End Select
      Next
      
      TitleCase = lsValue
   Else
      lsString2 = Split(Trim(sValue), " ")
      lsValue2 = Empty
   
      If UBound(lsString2) > 0 Then
         For lnCtr2 = 0 To UBound(lsString2)
            If Not IsNumeric(lsString2(lnCtr2)) Then
               If Len(lsString2(lnCtr2)) = 1 Then
                  lsValue2 = Trim(lsValue2) & " " & UCase(lsString2(lnCtr2))
               Else
                  If Trim(lsString2(lnCtr2)) <> "" Then
                     lsValue2 = Trim(lsValue2) & " " & UCase(Left(lsString2(lnCtr2), 1)) & Right(lsString2(lnCtr2), Len(lsString2(lnCtr2)) - 1)
                  End If
               End If
            Else
               lsValue2 = lsValue2 & " " & lsString2(lnCtr2)
            End If
         Next
         
         TitleCase = Trim(lsValue2)
      Else
         If Trim(sValue) <> Empty Then
            If Not IsNumeric(sValue) Then
               If Len(sValue) = 1 Then
                  TitleCase = UCase(sValue)
               Else
                  TitleCase = Trim(UCase(Left(sValue, 1)) & Right(sValue, Len(sValue) - 1))
               End If
            Else
               TitleCase = Trim(sValue)
            End If
         End If
      End If
   End If
   
endProc:
   Exit Function
errProc:
   TitleCase = sValue
End Function

Public Function getComputerName() As String
   Dim lnLen As Long
   Dim lsCompName As String
   
   'Create a buffer
   lnLen = MAX_COMPUTERNAME_LENGTH + 1
   lsCompName = String(lnLen, "X")
   
   'Get the computer name
   GetCompName lsCompName, lnLen
   
   'get only the actual data
   lsCompName = Left(lsCompName, lnLen)
   
   getComputerName = lsCompName
End Function

Public Function getEmployeeDepartment(oAppDriver As clsAppDriver, Optional ByVal fsEmployID As String = "") As String
   Dim lsSQL As String
   Dim lors As Recordset
   
   getEmployeeDepartment = ""
   
   'Remove any trailing space from fsEmployID
   fsEmployID = Trim(fsEmployID)
    
   If fsEmployID = "" Then
        fsEmployID = Trim(oAppDriver.EmployeeNo)
   End If
    
   'Exit from function if fsEmployID is empty
   If fsEmployID = "" Then
        Exit Function
   End If
   
   Set lors = New Recordset
   
   lsSQL = "SELECT sDeptIDxx" & _
          " FROM Employee_Master001" & _
          " WHERE sEmployID = " & strParm(fsEmployID)
      
   lors.Open oAppDriver.Connection.Execute(lsSQL, , adCmdText)
   
   If Not lors.EOF Then
        getEmployeeDepartment = lors("sDeptIDxx")
   End If

End Function

Public Function GotFocus() As Long
   GotFocus = GetFocus
End Function

Function IFNull(ByVal Value As Variant, Optional ByVal Repl As Variant = "")
   IFNull = IIf(IsNull(Value), Repl, IIf(IsMissing(Value), Repl, Value))
End Function

Function NullIf(ByVal Value As Variant, Optional ByVal NValue As Variant = "")
   NullIf = IIf(Value = NValue, Null, Value)
End Function

Public Function AllowRecSave(oAppDrivr As clsAppDriver, _
                             cRecdStat As xeRecordStatus, _
                             nEditMode As xeEditMode, _
                             sRemarks As String, _
                             dModified As Date, _
                             sApprovID As String) As Boolean
   Dim lnUserRght As Integer, lsUserName As String, lsUserIDxx As String
   Dim lasRights() As String, lnRights As Integer
   Dim lnSeek As Integer, lbSeek As Boolean
   Dim lsProcName As String

   lsProcName = "AllowRecSave"
   On Error GoTo errProc

   AllowRecSave = False

   With oAppDrivr
      If Trim(sRemarks) = "" Then
         ReDim lasRights(5) As String
         lasRights(1) = 231
         lasRights(2) = 230
         lasRights(3) = 228
         lasRights(4) = 192
         lasRights(5) = 230
      Else
         lasRights = Split(sRemarks, "»")
      End If
      lbSeek = False
      
      If nEditMode = xeModeAddNew Then
         ' Add Rights
         lnRights = CInt(lasRights(1))
         If (lnRights And .UserLevel) = 0 Then
            lbSeek = True
         End If
      ElseIf nEditMode = xeModeUpdate Then
         ' update Rights
         lnRights = CInt(lasRights(2))
         If (lnRights And .UserLevel) = 0 Then
            Select Case .UserLevel
            Case xeEncoder
               lbSeek = True
            Case xeSupervisor
               If DateDiff("d", dModified, .ServerDate) > 2 Then lbSeek = True
            Case xeManager
               If DateDiff("d", dModified, .ServerDate) > 7 Then lbSeek = True
            Case xeAudit
               lbSeek = True
            Case xeSysAdmin
               lbSeek = True
            End Select
         End If
      End If
      
      If lbSeek = False Then
         sApprovID = .UserID
         AllowRecSave = True
         GoTo endProc
      End If

      lnSeek = MsgBox("User Doesn't Have the Right to Update Record!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
      If lnSeek = vbNo Then GoTo endProc

      If GetApproval(oAppDrivr, lnUserRght, lsUserIDxx, lsUserName, sRemarks) = False Then GoTo endProc

      If nEditMode = xeModeAddNew Then
         If (lnUserRght And lnRights) = 0 Then
            MsgBox "Approving User Doesn't Have the Right to Create New Record!!!" & _
                  "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
            GoTo endProc
         End If
      ElseIf nEditMode = xeModeUpdate Then
         If (lnUserRght And lnRights) = 0 Then
            Select Case lnUserRght
            Case xeSupervisor
               If DateDiff("d", dModified, oAppDrivr.ServerDate) <= 2 Then lbSeek = False
            Case xeManager
               If DateDiff("d", dModified, oAppDrivr.ServerDate) <= 7 Then lbSeek = False
            Case xeSysAdmin
               If .UserLevel = xeSysAdmin Then lbSeek = False
            End Select

            If lbSeek Then
               MsgBox "Approving User Doesn't Have the Right to Update Record!!!" & _
                     "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      End If
   End With

   sApprovID = lsUserIDxx
   AllowRecSave = True

endProc:
   Exit Function
errProc:
   ShowError oAppDrivr _
           , lsProcName & "( " & "oAppDriver" _
                       & ", " & cRecdStat _
                       & ", " & nEditMode _
                       & ", " & sRemarks _
                       & ", " & dModified _
                       & ", " & sApprovID & " )"
End Function

Public Function AllowRec2Delete(oAppDrivr As clsAppDriver, _
                             cRecdStat As xeRecordStatus, _
                             sRemarks As String, _
                             dModified As Date, _
                             sApprovID As String) As Boolean
   Dim lnUserRght As Integer, lsUserName As String, lsUserIDxx As String
   Dim lasRights() As String, lnRights As Integer
   Dim lnSeek As Integer, lbSeek As Boolean
   Dim lsProcName As String

   lsProcName = "AllowRec2Delete"
   On Error GoTo errProc

   AllowRec2Delete = False

   With oAppDrivr
      If Trim(sRemarks) = "" Then
         ReDim lasRights(5) As String
         lasRights(1) = 231
         lasRights(2) = 230
         lasRights(3) = 228
         lasRights(4) = 192
         lasRights(5) = 230
      Else
         lasRights = Split(sRemarks, "»")
      End If
      lbSeek = False
      
      lnRights = CInt(lasRights(4))
      If (lnRights And .UserLevel) = 0 Then
         Select Case .UserLevel
         Case xeEncoder, xeSupervisor, xeAudit, xeSysAdmin, xeManager
            lbSeek = True
         End Select
      End If

      If lbSeek = False Then
         sApprovID = .UserID
         AllowRec2Delete = True
         GoTo endProc
      End If

      lnSeek = MsgBox("User Doesn't Have the Right to Delete Record!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
      If lnSeek = vbNo Then GoTo endProc

      If GetApproval(oAppDrivr, lnUserRght, lsUserIDxx, lsUserName, sRemarks) = False Then GoTo endProc

      If (lnUserRght And lnRights) = 0 Then
         If lbSeek Then
            MsgBox "Approving User Doesn't Have the Right to Update Record!!!" & _
                  "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
   End With

   sApprovID = lsUserIDxx
   AllowRec2Delete = True

endProc:
   Exit Function
errProc:
   ShowError oAppDrivr _
           , lsProcName & "( " & "oAppDriver" _
                       & ", " & cRecdStat _
                       & ", " & sRemarks _
                       & ", " & dModified _
                       & ", " & sApprovID & " )"
End Function

Public Function AllowRec2Cancel(oAppDrivr As clsAppDriver, _
                             cRecdStat As xeRecordStatus, _
                             sRemarks As String, _
                             dModified As Date, _
                             sApprovID As String) As Boolean
   Dim lnUserRght As Integer, lsUserName As String, lsUserIDxx As String
   Dim lasRights() As String, lnRights As Integer
   Dim lnSeek As Integer, lbSeek As Boolean
   Dim lsProcName As String

   lsProcName = "AllowRec2Cancel"
   On Error GoTo errProc

   AllowRec2Cancel = False

   With oAppDrivr
      If Trim(sRemarks) = "" Then
         ReDim lasRights(5) As String
         lasRights(1) = 231
         lasRights(2) = 230
         lasRights(3) = 228
         lasRights(4) = 192
         lasRights(5) = 230
      Else
         lasRights = Split(sRemarks, "»")
      End If
      lbSeek = False
      
      lnRights = CInt(lasRights(3))
      If (lnRights And .UserLevel) = 0 Then
         Select Case .UserLevel
         Case xeEncoder, xeSupervisor, xeAudit, xeSysAdmin
            lbSeek = True
         Case xeManager
            If DateDiff("d", dModified, .ServerDate) > 7 Then lbSeek = True
         End Select
      End If

      If lbSeek = False Then
         sApprovID = .UserID
         AllowRec2Cancel = True
         GoTo endProc
      End If

      lnSeek = MsgBox("User Doesn't Have the Right to Update Record!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
      If lnSeek = vbNo Then GoTo endProc

      If GetApproval(oAppDrivr, lnUserRght, lsUserIDxx, lsUserName, sRemarks) = False Then GoTo endProc

      If (lnUserRght And lnRights) = 0 Then
         Select Case lnUserRght
         Case xeManager
            If DateDiff("d", dModified, oAppDrivr.ServerDate) <= 7 Then lbSeek = False
         End Select

         If lbSeek Then
            MsgBox "Approving User Doesn't Have the Right to Update Record!!!" & _
                  "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
   End With

   sApprovID = lsUserIDxx
   AllowRec2Cancel = True

endProc:
   Exit Function
errProc:
   ShowError oAppDrivr _
           , lsProcName & "( " & "oAppDriver" _
                       & ", " & cRecdStat _
                       & ", " & sRemarks _
                       & ", " & dModified _
                       & ", " & sApprovID & " )"
End Function

Public Function AllowTransSave(oAppDrivr As clsAppDriver, _
                               cTranStat As xeTransactionStatus, _
                               nEditMode As xeEditMode, _
                               sRemarks As String, _
                               dModified As Date, _
                               sApprovID As String) As Boolean
   Dim lnUserRght As Integer, lsUserName As String, lsUserIDxx As String
   Dim lasRights() As String, lnRights As Integer
   Dim lnSeek As Integer, lbSeek As Boolean
   Dim lsProcName As String

   lsProcName = "AllowTransSave"
   On Error GoTo errProc

   AllowTransSave = False

   If cTranStat = xeStatePosted Then
      If oAppDrivr.UserLevel = xeSysMaster Then
         AllowTransSave = True
      Else
         MsgBox "Modification of posted or cancelled transaction is not allowed!!!" & vbCrLf & vbCrLf & _
               "Request not granted!!!", vbCritical, "Warning"
      End If
      GoTo endProc
   End If

   With oAppDrivr
      If Trim(sRemarks) = "" Then
         ReDim lasRights(5) As String
         lasRights(1) = 231
         lasRights(2) = 230
         lasRights(3) = 228
         lasRights(4) = 192
         lasRights(5) = 230
      Else
         lasRights = Split(sRemarks, "»")
      End If

      lbSeek = False
      If nEditMode = xeModeAddNew Then
         ' Add Rights
         lnRights = CInt(lasRights(1))
         If (lnRights And .UserLevel) = 0 Then
            lnSeek = MsgBox("User Doesn't Have the Right to Create New Transaction!!!" & _
                              vbCrLf & "Seek Authorized User's Approval?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
            If lnSeek = vbNo Then GoTo endProc
            lbSeek = True
         End If
      ElseIf nEditMode = xeModeUpdate Then
         ' update Rights
         lnRights = CInt(lasRights(2))

         If (lnRights And .UserLevel) = 0 Then
            Select Case .UserLevel
            Case xeEncoder
               lbSeek = True
            Case xeSupervisor
               If DateDiff("d", dModified, .ServerDate) > 2 Then lbSeek = True
            Case xeManager
               If DateDiff("d", dModified, .ServerDate) > 7 Then lbSeek = True
            Case xeAudit
               lbSeek = True
            Case xeSysAdmin
               lbSeek = True
            End Select
         End If
      End If

      If lbSeek = False Then
         sApprovID = .UserID
         AllowTransSave = True
         GoTo endProc
      End If

      lnSeek = MsgBox("User Doesn't Have the Right to Update Transaction!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
      If lnSeek = vbNo Then GoTo endProc

      If GetApproval(oAppDrivr, lnUserRght, lsUserIDxx, lsUserName, sRemarks) = False Then GoTo endProc

      If nEditMode = xeModeAddNew Then
         If (lnUserRght And lnRights) = 0 Then
            MsgBox "Approving User Doesn't Have the Right to Create New Transaction!!!" & _
                  "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
            GoTo endProc
         End If
      ElseIf nEditMode = xeModeUpdate Then
         If (lnUserRght And lnRights) = 0 Then
            Select Case lnUserRght
            Case xeSupervisor
               If DateDiff("d", dModified, oAppDrivr.ServerDate) <= 2 Then lbSeek = False
            Case xeManager
               If DateDiff("d", dModified, oAppDrivr.ServerDate) <= 31 Then lbSeek = False
            Case xeSysAdmin
               If .UserLevel <> xeSysAdmin Then lbSeek = False
            End Select

            If lbSeek Then
               MsgBox "Approving User Doesn't Have the Right to Update Transaction!!!" & _
                     "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      End If
   End With

   sApprovID = lsUserIDxx
   AllowTransSave = True

endProc:
   Exit Function
errProc:
   ShowError oAppDrivr _
           , lsProcName & "( " & "oAppDriver" _
                       & ", " & cTranStat _
                       & ", " & nEditMode _
                       & ", " & sRemarks _
                       & ", " & dModified _
                       & ", " & sApprovID & " )"
End Function

Public Function GetSplitedName(ByVal lsName As String) As Variant
   Dim lasName() As String
   Dim lsLName As String, lsFName As String, lsMName As String
   Dim lnCtr As Integer
   
   If lsName = Empty Then GoTo endProc
   
   lasName = Split(Trim(lsName), ",")
   lsLName = Trim(UCase(Left(lasName(0), 1)) & Mid(lasName(0), 2))
   If UBound(lasName) > 0 Then
      If Trim(lasName(1)) <> Empty Then
         lasName = Split(Trim(lasName(1)), " ")
         lsFName = Trim(UCase(Left(lasName(0), 1)) & Mid(lasName(0), 2))
         If UBound(lasName) > 0 Then
            lsMName = Trim(UCase(Left(lasName(UBound(lasName)), 1)) & Mid(lasName(UBound(lasName)), 2))
            For lnCtr = 1 To UBound(lasName) - 1
               lsFName = lsFName & " " & Trim(UCase(Left(lasName(lnCtr), 1)) & Mid(lasName(lnCtr), 2))
            Next
         End If
      End If
   End If
   
   ReDim lasName(2) As String
   lasName(0) = lsLName
   lasName(1) = lsFName
   lasName(2) = lsMName
   
   GetSplitedName = lasName
   
endProc:
   Exit Function
End Function

Public Sub ShowWarning(Message As String)
   MsgBox Message
End Sub

Private Sub ShowError(p_oAppDrivr As clsAppDriver, ByVal lsProcName As String)
    With p_oAppDrivr
       .xLogError Err.Number, Err.Description, "clsMainModules", lsProcName, Erl
    End With
    With Err
       .Raise .Number, .Source, .Description
    End With
End Sub

Function ErrorCause(oCon As ADODB.Connection) As xeErrorType
   If oCon.Errors.Count > 0 Then
      Select Case oCon.Errors.Item(0).NativeError
      'Case 1062 'ER_DUP_ENTRY » Duplicate key Entry; Insert or Update will cause a duplicate record
      'Case 1088 'ER_ALTER_INFO » Records: %ld Duplicates: %ld
      'Case 1092 'ER_INSERT_INFO » Records: %ld Duplicates: %ld Warnings: %ld
      'Case 1169 'ER_DUP_UNIQUE » Can't write, because of unique constraint, to table '%s'
      'Case 1020 'ER_CHECKREAD » Record has changed since last read in table '%s'
      'Case 1022 'ER_DUP_KEY » Can't write; duplicate key in table '%s'
      Case 1062, 1088, 1092, 1169, 1022
         ErrorCause = 1
      Case 32   'Error cause by updating a record. Value of fields updated has the same value with
                 'the old data
         ErrorCause = 2
      Case Else
         ' MsgBox oCon.Errors.Item(0).NativeError & " » " & oCon.Errors.Item(0).Description
         ErrorCause = -1
      End Select
      MsgBox oCon.Errors.Item(0).NativeError & " » " & oCon.Errors.Item(0).Description
   Else
      ErrorCause = 0
   End If
End Function

Function FieldParam(ByVal lnFieldType As Integer, ByVal lvFieldValue) As String
   Select Case lnFieldType
   Case 129, 130, 202, 200 ' string
      If IsNull(lvFieldValue) Then
         FieldParam = "NULL"
      Else
         FieldParam = strParm(lvFieldValue)
      End If
   Case 2, 3, 11, 17, 72      ' numeric without decimal point
      If IsNumeric(lvFieldValue) = True Or IsNull(lvFieldValue) = True Then
         'set the value to 0 if value is empty/null
         FieldParam = IIf(IsEmpty(IFNull(lvFieldValue, 0)), 0, IFNull(lvFieldValue, 0))
      Else
         FieldParam = strParm(lvFieldValue)
      End If
   Case 4, 5, 6, 131             ' numeric with decimal point
      FieldParam = CDbl(IFNull(lvFieldValue, 0))
   Case 7, 133 ' datetime
      If IsNull(lvFieldValue) Or IsEmpty(lvFieldValue) Or lvFieldValue = Empty Then
         FieldParam = "NULL"
      Else
         FieldParam = dateParm(lvFieldValue)
      End If
   Case 134 'time
     If IsNull(lvFieldValue) Or IsEmpty(lvFieldValue) Or lvFieldValue = Empty Then
         FieldParam = "NULL"
      Else
         FieldParam = "'" & Format(lvFieldValue, "hh:mm:ss") & "'"
      End If
   Case 135 ' date
     If IsNull(lvFieldValue) Or IsEmpty(lvFieldValue) Or lvFieldValue = Empty Then
         FieldParam = "NULL"
      Else
         FieldParam = dateParm(lvFieldValue)
      End If
   Case Else
      If IsNull(lvFieldValue) Then
         FieldParam = "NULL"
      Else
         FieldParam = lvFieldValue
      End If
   End Select
End Function

Function ADO2SQL( _
                 lors As ADODB.Recordset, _
                 ByVal lsTable As String, _
                 Optional ByVal lsFilter As String = "", _
                 Optional ByVal lsModified As String = "", _
                 Optional ByVal ldModified As String = "", _
                 Optional ByVal lsExcluded As String = "" _
                 ) As String
   Dim lnCtr As Integer
   Dim lsSQL As String
   Dim lnFldCnt As Integer
   
   'Adjust the number of fields to be extracted
   lnFldCnt = lors.Fields.Count - 1
   
'   If lsModified <> "" Then lnFldCnt = lnFldCnt - 1
'   If ldModified <> "" Then lnFldCnt = lnFldCnt - 1
   
   If lsModified <> "" Then lsExcluded = lsExcluded & "»sModified"
   If ldModified <> "" Then lsExcluded = lsExcluded & "»dModified"
   
   'Extract fields
   lsSQL = ""
   For lnCtr = 0 To lnFldCnt
      If InStr(1, lsExcluded, lors.Fields(lnCtr).Name, 1) = 0 Then
         If lsFilter = "" Then
            lsSQL = lsSQL & ", " & AddField(lors.Fields(lnCtr).Name, FieldParam(lors(lnCtr).Type, lors(lnCtr).Value))
         Else
            If IsNull(lors(lnCtr).OriginalValue) Then
               If IsNull(lors(lnCtr).Value) = False Then
                  lsSQL = lsSQL & ", " & AddField(lors.Fields(lnCtr).Name, FieldParam(lors(lnCtr).Type, lors(lnCtr).Value))
               End If
            ElseIf IsNull(lors(lnCtr).Value) Then
               If IsNull(lors(lnCtr).OriginalValue) = False Then
                  lsSQL = lsSQL & ", " & AddField(lors.Fields(lnCtr).Name, FieldParam(lors(lnCtr).Type, lors(lnCtr).Value))
               End If
            ElseIf lors(lnCtr).OriginalValue <> lors(lnCtr).Value Then
               If IsNull(lors(lnCtr).Value) = False Then
                  lsSQL = lsSQL & ", " & AddField(lors.Fields(lnCtr).Name, FieldParam(lors(lnCtr).Type, lors(lnCtr).Value))
               End If
            End If
         End If
      End If
   Next
   
   If lsSQL = "" Then
      ADO2SQL = ""
   Else
      'if sModified and dModified are present then add them
      lsSQL = Mid(lsSQL, 2)
      If lsModified <> "" Then lsSQL = lsSQL & ", sModified = " & strParm(lsModified)
      If ldModified <> "" Then lsSQL = lsSQL & ", dModified = " & dateParm(ldModified)
      If lsFilter <> "" Then
         lsSQL = "UPDATE " & lsTable & " SET " & _
                     lsSQL & _
                 " WHERE " & lsFilter
      Else
         lsSQL = "INSERT INTO " & lsTable & " SET " & _
                     lsSQL
      End If
      ADO2SQL = lsSQL
   End If
   
endProc:
   Exit Function
End Function

Function AddField(ByVal sFieldName As String, ByVal vValue As Variant) As String
   AddField = sFieldName & " = " & vValue
   
'   Select Case UCase(TypeName(vValue))
'      Case "STRING", "EMPTY"
'         If Len(vValue) = 1 Then
'            If Left(sFieldName, 1) = "c" Then
'               vValue = IIf(Asc(vValue) = 0, "", vValue)
'               AddField = sFieldName & " = " & strParm(vValue)
'            Else
'               If Asc(vValue) = 0 Then
'                  AddField = sFieldName & " = " & strParm("")
'               Else
'                  AddField = sFieldName & " = " & strParm(vValue)
'               End If
'            End If
'         Else
'            Select Case LCase(Left(sFieldName, 1))
'            Case "n"
'               AddField = sFieldName & " = 0"
'            Case "d"
'               AddField = sFieldName & " = NULL"
'            Case "c"
'               vValue = IIf(Asc(vValue) = 0, "", vValue)
'               AddField = sFieldName & " = " & strParm(vValue)
'            Case Else
'               AddField = sFieldName & " = " & strParm(vValue)
'            End Select
'         End If
'      Case "NULL"
'         AddField = sFieldName & " = NULL"
'      Case "DATE"
'         AddField = sFieldName & " = " & dateParm(vValue)
'      Case Else
'         If Left(sFieldName, 1) = "c" Then
'            AddField = sFieldName & " = " & strParm(vValue)
'         Else
'            AddField = sFieldName & " = " & vValue
'         End If
'   End Select
End Function


'Returns
'  0 - valid
'  1 - invalid advance date
'  2 - invalid back date
'  3 - invalid date
Function validateTranDate(oAppDrivr As clsAppDriver, dTransact As Variant) As Long
   Dim lnAllow As Long
   
   validateTranDate = 3
   
   If IsDate(dTransact) Then
      validateTranDate = 0
      
      'check if date was valid from advance date limit
      lnAllow = IFNull(oAppDrivr.getConfiguration("AdvDatex"), 0)
      If DateDiff("d", oAppDrivr.ServerDate, dTransact) > lnAllow Then
         validateTranDate = 1
         Exit Function
      End If
      
      'check if date was valid from back date limit
      lnAllow = IFNull(oAppDrivr.getConfiguration("Backdate"), 0)
      If DateDiff("d", dTransact, oAppDrivr.ServerDate) > lnAllow Then
         validateTranDate = 2
         Exit Function
      End If
   End If
End Function

Function getValidTranDate(SysDate As Date, TransactDate As Variant) As Date
   getValidTranDate = SysDate
   
   If IsDate(TransactDate) Then
      Debug.Print SysDate, TransactDate, DateDiff("d", TransactDate, SysDate)
      If DateDiff("d", SysDate, TransactDate) > 0 Then
         MsgBox "Advancement of date is not allowed!" & vbCrLf & _
                  "Verify your entry then Try Again!", vbCritical, "Warning"
      Else
         getValidTranDate = CDate(TransactDate)
      End If
   End If
End Function

Function getValidDate(oAppDrivr As clsAppDriver, dTransact As Variant) As Date
   Dim lnAllow As Integer
   getValidDate = oAppDrivr.ServerDate
   
   If IsDate(dTransact) Then
      lnAllow = IFNull(oAppDrivr.getConfiguration("AdvDatex"), 0)
   
      If DateDiff("d", oAppDrivr.ServerDate, dTransact) > lnAllow Then
         MsgBox "Advancement of date exceeds the limit!" & vbCrLf & _
                  "Verify your entry then Try Again!", vbCritical, "Warning"
      Else
         getValidDate = CDate(dTransact)
      End If
   End If
End Function

Function isFileExist(file As String) As Boolean
    Dim loFileSys As FileSystemObject

    Set loFileSys = New Scripting.FileSystemObject
    isFileExist = loFileSys.FileExists(file)
    Set loFileSys = Nothing
End Function

Function Encode(oAppDrivr As clsAppDriver, ByVal Value As String) As String
   Dim lors As Recordset
   Dim lsSQL As String
   
   If Len(Trim(Value)) < 0 Then
      Encode = Value
   Else
   
      lsSQL = "SELECT HEX(ENCODE(REVERSE(" & strParm(Value) & "), " & strParm("PETMGR") & "))"
      Set lors = oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
         
      Encode = lors(0)
   End If
End Function

Function Decode(oAppDrivr As clsAppDriver, ByVal Value As String) As String
   Dim lors As Recordset
   Dim lsSQL As String
   
   'Without cast it will return a binary string
   'so use cast to convert it into a string
   If Len(Value) < 10 And IsNumeric(Value) Then
      Decode = Value
   Else
      lsSQL = "SELECT CAST(REVERSE(DECODE(UNHEX(" & strParm(Value) & "), " & strParm("PETMGR") & ")) AS CHAR)"
      Set lors = oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
      Decode = lors(0).Value
   End If
End Function

Public Function Dec2Hex(ByVal nmbDec As Long) As String
    Dec2Hex = Format$(Hex(nmbDec), "00")
    Dec2Hex = IIf(Len(Dec2Hex) Mod 2 = 0, Dec2Hex, "0" & Dec2Hex)
End Function

Public Function Hex2Dec(ByVal strNumHex As String) As Long
   Hex2Dec = Val("&H" & strNumHex)
End Function

Public Function Bin2Dec(Num As String) As Long
   Dim n As Integer
   Dim x As Integer
   Dim a As Integer
   n = Len(Num) - 1
   a = n
   Do While n > -1
      x = Mid(Num, ((a + 1) - n), 1)
      Bin2Dec = IIf((x = "1"), Bin2Dec + (2 ^ (n)), Bin2Dec)
      n = n - 1
   Loop
End Function

Public Function Dec2Bin(ByVal Num As Long) As String
   Select Case Num
      Case Is > 1: Dec2Bin = Dec2Bin(Num \ 2) & (Num Mod 2)
      Case 0, 1:   Dec2Bin = Num
      Case Else:   Dec2Bin = -Dec2Bin(-Num)
   End Select
End Function

Public Function PadLeft(ByVal fsStr As String, ByVal fnLen As Integer, ByVal fsPad As String) As String
   If fnLen <= Len(fsStr) Then
      PadLeft = fsStr
   Else
      PadLeft = String(fnLen - Len(fsStr), Left(fsPad, 1)) & fsStr
   End If
End Function

Public Function Multi_Find( _
       ByRef oRs As ADODB.Recordset, _
       sCriteria As String) As Boolean
      
   Dim clone_rs As ADODB.Recordset
   
   If oRs.RecordCount > 0 Then
      Set clone_rs = oRs.Clone
      
      clone_rs.Filter = sCriteria
   
      If clone_rs.EOF Or clone_rs.BOF Then
         oRs.MoveLast
         oRs.MoveNext
      Else
         oRs.Bookmark = clone_rs.Bookmark
         Multi_Find = True
      End If
   
      clone_rs.Close
      Set clone_rs = Nothing
   End If
End Function

Public Function Random(ByVal from As Long, ByVal upto As Long) As Integer
   Randomize
   Random = Int((upto - from + 1) * Rnd + from)
End Function

Public Function isValidApproveCode( _
    ByVal fsSysReq As String, _
    ByVal fsBranch As String, _
    ByVal fsIssuee As String, _
    ByVal fsDatexx As String, _
    ByVal fsMiscxx As String, _
    ByVal fsCdeGvn As String _
    ) As Boolean
    
   Dim loCode As clsCodeGen
   Set loCode = New clsCodeGen
   
   loCode.System = fsSysReq
   loCode.Branch = fsBranch
   loCode.IssuedBy = fsIssuee
   loCode.DateRequested = fsDatexx
   loCode.MiscInfo = fsMiscxx
    
   If loCode.Encode Then
      Debug.Print loCode.Result
      If fsBranch <> "" Then
         If loCode.Equal(loCode.Result, fsCdeGvn) >= 0 Then
             isValidApproveCode = True
         End If
      Else
         If loCode.Equalx(loCode.Result, fsCdeGvn) >= 0 Then
             isValidApproveCode = True
         End If
      End If
   End If
End Function

Function isValidEmail(ByVal strCheck As String) As Boolean
   Dim bCheck As Boolean
   Dim sDomnType As String
   Dim sDomnName As String
   Dim I As Integer

   Const sInvalidChr As String = "!#$%^&*()=+{}[]|\;:'/?>,< "

   bCheck = Not InStr(1, strCheck, Chr(34)) > 0 'Check to see if there is a double quote
   If Not bCheck Then GoTo ExitFunction

   bCheck = Not InStr(1, strCheck, "..") > 0 'Check to see if there are consecutive dots
   If Not bCheck Then GoTo ExitFunction

   ' Check for invalid characters.
   If Len(strCheck) > Len(sInvalidChr) Then
       For I = 1 To Len(sInvalidChr)
           If InStr(strCheck, Mid(sInvalidChr, I, 1)) > 0 Then
               bCheck = False
               GoTo ExitFunction
           End If
       Next
   Else
       For I = 1 To Len(strCheck)
           If InStr(sInvalidChr, Mid(strCheck, I, 1)) > 0 Then
               bCheck = False
               GoTo ExitFunction
           End If
       Next
   End If
   
   'Check for an @ symbol
   If InStr(1, strCheck, "@") > 1 Then
       bCheck = Len(Left(strCheck, InStr(1, strCheck, "@") - 1)) > 0
   Else
       bCheck = False
   End If
   If Not bCheck Then GoTo ExitFunction
   
   'Check to see if there are too many @'s
   strCheck = Right(strCheck, Len(strCheck) - InStr(1, strCheck, "@"))
   bCheck = Not InStr(1, strCheck, "@") > 0
   If Not bCheck Then GoTo ExitFunction
   
   'Check Domain TYpe
   sDomnType = Right(strCheck, Len(strCheck) - InStr(1, strCheck, "."))
   bCheck = Len(sDomnType) > 0 And InStr(1, strCheck, ".") < Len(strCheck)
   If Not bCheck Then GoTo ExitFunction
   
   strCheck = Left(strCheck, Len(strCheck) - Len(sDomnType) - 1)
   Do Until InStr(1, strCheck, ".") <= 1
       If Len(strCheck) >= InStr(1, strCheck, ".") Then
           strCheck = Left(strCheck, Len(strCheck) - (InStr(1, strCheck, ".") - 1))
       Else
           bCheck = False
           GoTo ExitFunction
       End If
   Loop
   If strCheck = "." Or Len(strCheck) = 0 Then bCheck = False

ExitFunction:
   isValidEmail = bCheck
End Function

'kalyptus - 2018.04.03 02:30pm
Public Function HexToString(ByVal HexToStr As String) As String
   Dim strTemp   As String
   Dim strReturn As String
   Dim I As Long
    For I = 1 To Len(HexToStr) Step 2
        strTemp = Chr$(Val("&H" & Mid$(HexToStr, I, 2)))
        strReturn = strReturn & strTemp
    Next I
    HexToString = strReturn
End Function

'kalyptus - 2018.04.03 02:30pm
Public Function StringToHex(ByVal StrToHex As String) As String
   Dim strTemp   As String
   Dim strReturn As String
   Dim I         As Long
    For I = 1 To Len(StrToHex)
        strTemp = Hex$(Asc(Mid$(StrToHex, I, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & strTemp
    Next I
    StringToHex = strReturn
End Function

'mac 2024.04.04 02:14pm
Public Sub OnlinePointsEntry(ByVal fsProdctID As String, _
                              ByVal fsUserIDxx As String, _
                              ByVal fsGCardNox As String, _
                              ByVal fsSourceCD As String, _
                              ByVal fsReferNox As String, _
                              ByVal fsDigital As String, _
                              ByVal fdTransact As String)
   Dim lsCommand As String
   
   'Delete the destination file of GCard peeping
   If FileExists("D:\GGC_Java_Systems\temp\res.tmp") Then
      Kill "D:\GGC_Java_Systems\temp\res.tmp"
   End If
   
   lsCommand = fsProdctID & " " & fsUserIDxx & " " & fsGCardNox & " " & fsDigital & " " & fsReferNox & " " & fsSourceCD & " " & fdTransact
   Debug.Print lsCommand
   If RMJExecute("D:\GGC_Java_Systems\gcard-online-points-entry.bat " & lsCommand) = 0 Then
      MsgBox "Transaction G-Card points encode successfully.", vbInformation, "Notice"
   Else
      MsgBox "Unable to encode G-Card points.", vbCritical, "Notice"
   End If
endProc:
   Exit Sub
End Sub

Public Function FileRead(ByVal lsFileName As String) As String
    Dim lsValue As String
    
    If Not FileExists(lsFileName) Then
        FileRead = ""
        Exit Function
    End If
    
    Open lsFileName For Input As #1
    Line Input #1, lsValue
    Close #1

    FileRead = lsValue
End Function

Public Function FileWrite(ByVal lsFileName As String, ByVal lsValue As String)
   Open lsFileName For Output As #1
   Print #1, lsValue
   Close #1
End Function

Public Function FileExists(ByVal sFileName As String) As Boolean
   Dim intReturn As Integer

   On Error GoTo FileExists_Error
   intReturn = GetAttr(sFileName)
   FileExists = True
    
Exit Function
FileExists_Error:
    FileExists = False
End Function

Public Function RMJExecute(ByVal fsPath As String) As Integer
   Dim lTaskID As Long, lPID As Long, lExitCode As Long, sAppDir As String

   lTaskID = Shell(fsPath, vbHide)
   lPID = OpenProcess(PROCESS_ALL_ACCESS, True, lTaskID)

   If lPID Then
      'WAIT FOR PROCESS TO finish
      'Note, you must now enter a value in the form and click close.
      Call WaitForSingleObject(lPID, INFINITE)

      'Get EXIT PROCESS
      If GetExitCodeProcess(lPID, lExitCode) Then
          RMJExecute = lExitCode
      Else
          RMJExecute = -3
      End If
   Else
      RMJExecute = -4
   End If
   
   lTaskID = CloseHandle(lPID)
End Function

Public Function RMJConnected()
   Dim hInet As Long, hUrl As Long, Flags As Long, url As Variant
  
   hInet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_PRECONFIG, _
                        vbNullString, vbNullString, 0&)

   If hInet Then
      Flags = INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_NO_CACHE_WRITE _
               Or INTERNET_FLAG_RELOAD
      
      hUrl = InternetOpenUrl(hInet, Environ$("GUANZON_WEB_SERVER"), _
      vbNullString, 0, Flags, 0)
      
      If hUrl Then
         RMJConnected = True
         Debug.Print "Your computer is connected to Guanzon"
         Call InternetCloseHandle(hUrl)
      Else
         RMJConnected = False
         Debug.Print "Your computer is not connected to Guanzon"
      End If
   End If
   
   Call InternetCloseHandle(hInet)
End Function

