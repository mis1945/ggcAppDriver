VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmCodeApproval 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCodeApproval.frx":0000
   ScaleHeight     =   4350
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtIssuee 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2865
      TabIndex        =   3
      Top             =   2520
      Width           =   3750
   End
   Begin VB.TextBox txtApprovalCode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2865
      TabIndex        =   1
      Top             =   2100
      Width           =   1710
   End
   Begin xrControl.xrButton xrButton 
      Height          =   495
      Index           =   0
      Left            =   4245
      TabIndex        =   4
      Top             =   3330
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   873
      Caption         =   "&OK"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15720398
      ForeColor       =   4194304
      BackColorDown   =   16051167
      BorderColorFocus=   7883077
      BorderColorHover=   16379363
   End
   Begin xrControl.xrButton xrButton 
      Height          =   495
      Index           =   1
      Left            =   5505
      TabIndex        =   5
      Top             =   3330
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   873
      Caption         =   "&Cancel"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15720398
      ForeColor       =   4194304
      BackColorDown   =   16051167
      BorderColorFocus=   7883077
      BorderColorHover=   16379363
   End
   Begin VB.Label lblFormTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SYSTEM CODE APPROVAL"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00ECE1D0&
      Height          =   360
      Index           =   1
      Left            =   165
      TabIndex        =   9
      Top             =   330
      Width           =   4965
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GGC - SEG '14"
      ForeColor       =   &H000080FF&
      Height          =   195
      Index           =   3
      Left            =   5640
      TabIndex        =   8
      Top             =   4050
      Width           =   1065
   End
   Begin VB.Label lblfield 
      BackStyle       =   0  'Transparent
      Caption         =   "AUTHORIZED PERSONNEL."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   3405
      TabIndex        =   7
      Top             =   1530
      Width           =   2625
   End
   Begin VB.Label lblfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   2100
      TabIndex        =   0
      Top             =   2160
      Width           =   570
   End
   Begin VB.Label lblfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Issuee"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   2100
      TabIndex        =   2
      Top             =   2595
      Width           =   705
   End
   Begin VB.Label lblFormTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SYSTEM CODE APPROVAL"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   0
      Left            =   150
      TabIndex        =   6
      Top             =   330
      Width           =   4515
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblfield 
      BackStyle       =   0  'Transparent
      Caption         =   "    System Approval is required by the object.   Seek  assistance  from an"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Index           =   2
      Left            =   1560
      TabIndex        =   10
      Top             =   1260
      Width           =   4980
   End
End
Attribute VB_Name = "frmCodeApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pbCancel As Boolean
Private psAppPath As String
Private poMod As New clsMainModules
Private p_oAppDrivr As clsAppDriver

Private p_sUserIDxx As String
Private p_sIssueexx As String
Private p_cIssueexx As String
Private p_sCodeAprv As String

Property Get UserID() As String
   UserID = p_sUserIDxx
End Property

Property Get Issuee() As String
   Issuee = p_sIssueexx
End Property

Property Get IssueeType() As String
   IssueeType = p_cIssueexx
End Property

Property Get CodeApproval() As String
   CodeApproval = p_sCodeAprv
End Property

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Let AppPath(ByVal Value As String)
   psAppPath = Value
End Property

Property Get Cancel() As Boolean
   Cancel = pbCancel
End Property

Private Sub Form_Initialize()
   pbCancel = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         poMod.SetNextFocus
      Case vbKeyUp
         poMod.SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set poMod = Nothing
End Sub

Private Sub txtApprovalCode_Validate(Cancel As Boolean)
   p_sCodeAprv = txtApprovalCode.Text
End Sub

Private Sub txtIssuee_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyF3, vbKeyReturn
      Call getIssuee(txtIssuee, KeyCode = vbKeyF3)
   End Select
End Sub

Private Sub xrButton_Click(Index As Integer)
   pbCancel = Index = 1
   Me.Hide
End Sub

Private Sub Form_Load()

'   ShockwaveFlash1.Movie = psAppPath & "\Images\hand_key.swf"
'   ShockwaveFlash1.Play
End Sub

Private Sub getIssuee(ByVal fsValue As String, ByVal fbSearch As Boolean)
   Dim lsSQL As String
   Dim lors As Recordset
   Dim lasSelected() As String
   Dim lsConditn As String
   
   If fsValue = "" And fbSearch = False Then
      p_sUserIDxx = ""
      p_sIssueexx = ""
      p_cIssueexx = ""
      GoTo endProc
   End If
                   
   lsSQL = "SELECT" & _
               "  a.sUserIDxx" & _
               ", CONCAT(sLastName, ', ', sFrstName) sIssueexx" & _
               ", b.sEmpLevID" & _
               ", b.sDeptIDxx" & _
               ", b.sPositnID" & _
               ", b.sEmployID" & _
          " FROM xxxSysUser a" & _
               ", Employee_Master001 b" & _
                  " LEFT JOIN Client_Master c ON b.sEmployID = c.sClientID" & _
          " WHERE a.sEmployNo = b.sEmployID" & _
               " AND b.cRecdStat = '1'" & _
          " GROUP BY sEmployID"
'          " UNION " & _
'          " SELECT" & _
'                  "  sDeptIDxx sUserIDxx" & _
'                  ", sDeptName sIssueexx" & _
'                  ", '' sEmpLevID" & _
'                  ", sDeptIDxx" & _
'          " FROM Department " & _
'            " WHERE sDeptIDxx IN('021','022','034','025','027','026')"

   If fbSearch Then
      If fsValue <> "" Then
         lasSelected = poMod.GetSplitedName(fsValue)
         lsConditn = "c.sLastName LIKE " & poMod.strParm(lasSelected(0) & "%") & _
                " AND c.sFrstName LIKE " & poMod.strParm(IIf(UBound(lasSelected) > 0, lasSelected(1), "") & "%") & _
                " AND c.sMiddName LIKE " & poMod.strParm(IIf(UBound(lasSelected) > 1, lasSelected(2), "") & "%")
      End If
   Else
      lasSelected = poMod.GetSplitedName(fsValue)
      lsConditn = "c.sLastName = " & poMod.strParm(lasSelected(0) & "%") & _
             " AND c.sFrstName = " & poMod.strParm(IIf(UBound(lasSelected) > 0, lasSelected(1), "") & "%") & _
             " AND c.sMiddName = " & poMod.strParm(IIf(UBound(lasSelected) > 1, lasSelected(2), "") & "%")
   End If
   
   If lsConditn <> "" Then
      lsSQL = poMod.AddCondition(lsSQL, lsConditn)
   End If
   
   Debug.Print lsSQL
   
   Set lors = p_oAppDrivr.Connection.Execute(lsSQL, , adCmdText)
   
   
   If lors.RecordCount = 0 Then
      p_sUserIDxx = ""
      p_sIssueexx = ""
      p_cIssueexx = ""
      GoTo endProc
   ElseIf lors.RecordCount = 1 Then
      p_sUserIDxx = lors("sUserIDxx")
      p_sIssueexx = lors("sIssueexx")
      
      If lors("sEmpLevID") = "4" Then
         p_cIssueexx = "0"
      ElseIf lors("sEmpLevID") = "5" Or lors("sEmployID") = "M00112000440" Then
         p_cIssueexx = "9" 'general manager/grace padlan
      Else
         Select Case lors("sDeptIDxx")
         Case "021"   'Human Capital Management
            p_cIssueexx = "1"
         Case "022"   'Credit Support Services
            p_cIssueexx = "2"
         Case "034"   'Compliance Management
            p_cIssueexx = "3"
         Case "025"   'Marketing & Promotions
            p_cIssueexx = "4"
         Case "027"    'After Sales Management
            p_cIssueexx = "5"
         Case "035"   'Telemarketing
            p_cIssueexx = "6"
         Case "024"   'Supply Chain Management
            p_cIssueexx = "7"
         Case "026"   'Management Information Systems
            p_cIssueexx = "X"
         Case "015"  'Sales;CI/Collector
            If lors("sPositnID") = "091" Or _
               lors("sPositnID") = "299" Or _
               lors("sPositnID") = "298" Or _
               lors("sPositnID") = "056" Then
               p_cIssueexx = "8"
            End If
         Case Else
            p_cIssueexx = ""
         End Select
      End If
   Else
      lsSQL = poMod.KwikBrowse(p_oAppDrivr, lors, "sUserIDxx»sIssueexx", "ID»Issuee")
      If lsSQL <> "" Then
         lasMaster = Split(lsSQL, "»")
         
         p_sUserIDxx = lasMaster(0)
         p_sIssueexx = lasMaster(1)
         
         If lasMaster(2) = "4" Then
            p_cIssueexx = "0"
         ElseIf lasMaster(2) = "5" Or lasMaster(5) = "M00112000440" Then
            p_cIssueexx = "9" 'general manager/grace padlan
         Else
            Select Case lasMaster(3)
            Case "021"   'Human Capital Management
               p_cIssueexx = "1"
            Case "022"   'Credit Support Services
               p_cIssueexx = "2"
            Case "034"   'Compliance Management
               p_cIssueexx = "3"
            Case "025"   'Marketing & Promotions
               p_cIssueexx = "4"
            Case "027"   'After Sales Management
               p_cIssueexx = "5"
            Case "035"   'Telemarketing
               p_cIssueexx = "6"
            Case "024"   'Supply Chain Management
               p_cIssueexx = "7"
            Case "026"
               p_cIssueexx = "X"
            Case "015"  'Sales;CI/Collector
               If lors("sPositnID") = "091" Or _
                  lors("sPositnID") = "299" Or _
                  lors("sPositnID") = "298" Or _
                  lors("sPositnID") = "056" Then
                  p_cIssueexx = "8"
               End If
            Case Else
               p_cIssueexx = ""
            End Select
         End If
      Else
         p_sUserIDxx = ""
         p_sIssueexx = ""
         p_cIssueexx = ""
      End If
   End If

endProc:
   txtIssuee = p_sIssueexx
End Sub
