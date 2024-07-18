Attribute VB_Name = "modRMJ"
'Option Explicit
'
''-------------------------------------------------------------------------------------'
''  Execute
''-------------------------------------------------------------------------------------'
'Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
'Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
'
'Private Const INFINITE = &HFFFFFFFF       '  Infinite timeout
'Private Const SYNCHRONIZE = &H100000
'Private Const STILL_ACTIVE = 0
'Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
'Private Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
''-------------------------------------------------------------------------------------'
''  Execute
''-------------------------------------------------------------------------------------'
'
''-------------------------------------------------------------------------------------'
''  Internet Connection
''-------------------------------------------------------------------------------------'
'Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" _
'   (ByVal hInet As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, _
'      ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
'
'Private Declare Function InternetCloseHandle Lib "wininet.dll" _
'   (ByVal hInet As Long) As Long
'
'Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
'   (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As _
'      String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
'
'Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
'Private Const INTERNET_FLAG_RELOAD = &H80000000
'Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
'Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
''-------------------------------------------------------------------------------------'
''  Internet Connection
''-------------------------------------------------------------------------------------'
'
'Public Function RMJExecute(ByVal fsPath As String) As Integer
'   Dim lTaskID As Long, lPID As Long, lExitCode As Long, sAppDir As String
'
'   lTaskID = Shell(fsPath, vbHide)
'   lPID = OpenProcess(PROCESS_ALL_ACCESS, True, lTaskID)
'
'   If lPID Then
'      'WAIT FOR PROCESS TO finish
'      'Note, you must now enter a value in the form and click close.
'      Call WaitForSingleObject(lPID, INFINITE)
'
'      'Get EXIT PROCESS
'      If GetExitCodeProcess(lPID, lExitCode) Then
'          RMJExecute = lExitCode
'      Else
'          RMJExecute = -3
'      End If
'   Else
'      RMJExecute = -4
'   End If
'
'   lTaskID = CloseHandle(lPID)
'End Function
'
'Public Function RMJConnected()
'   Dim hInet As Long, hUrl As Long, Flags As Long, url As Variant
'
'   hInet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_PRECONFIG, _
'                        vbNullString, vbNullString, 0&)
'
'   If hInet Then
'      Flags = INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_NO_CACHE_WRITE _
'               Or INTERNET_FLAG_RELOAD
'
'      hUrl = InternetOpenUrl(hInet, Environ$("GUANZON_WEB_SERVER"), _
'      vbNullString, 0, Flags, 0)
'
'      If hUrl Then
'         RMJConnected = True
'         Debug.Print "Your computer is connected to Guanzon"
'         Call InternetCloseHandle(hUrl)
'      Else
'         RMJConnected = False
'         Debug.Print "Your computer is not connected to Guanzon"
'      End If
'   End If
'
'   Call InternetCloseHandle(hInet)
'End Function
'
''Public Function FileRead(ByVal lsFileName As String) As String
''    Dim lsValue As String
''
''    If Not FileExists(lsFileName) Then
''        FileRead = ""
''        Exit Function
''    End If
''
''    Open lsFileName For Input As #1
''    Line Input #1, lsValue
''    Close #1
''
''    FileRead = lsValue
''End Function
''
''Public Function FileWrite(ByVal lsFileName As String, ByVal lsValue As String)
''   Open lsFileName For Output As #1
''   Print #1, lsValue
''   Close #1
''End Function
''
''Public Function FileExists(ByVal sFileName As String) As Boolean
''   Dim intReturn As Integer
''
''   On Error GoTo FileExists_Error
''   intReturn = GetAttr(sFileName)
''   FileExists = True
''
''Exit Function
''FileExists_Error:
''    FileExists = False
''End Function
'
