VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmModules 
   BorderStyle     =   0  'None
   Caption         =   "Modules"
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1860
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   540
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   3281
      BorderStyle     =   1
      Begin VB.ComboBox cmbDivision 
         Height          =   315
         ItemData        =   "frmModules.frx":0000
         Left            =   1095
         List            =   "frmModules.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   690
         Width           =   1965
      End
      Begin VB.ComboBox cmbProduct 
         Height          =   315
         ItemData        =   "frmModules.frx":0028
         Left            =   1095
         List            =   "frmModules.frx":0035
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1020
         Width           =   1965
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1095
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   150
         Width           =   1965
      End
      Begin VB.TextBox txtField 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   3
         Left            =   1095
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   1350
         Width           =   3120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   13
         Top             =   750
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product ID"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   11
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Field Name"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   1410
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Module ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   210
         Width           =   885
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   315
         Left            =   1170
         Tag             =   "et0;ht2"
         Top             =   225
         Width           =   1965
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   0
      Left            =   3720
      TabIndex        =   4
      Top             =   2625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Cancel"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmModules.frx":0056
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   1
      Left            =   2940
      TabIndex        =   5
      Top             =   2625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Browse"
      AccessKey       =   "B"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmModules.frx":07D0
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   2
      Left            =   2160
      TabIndex        =   6
      Top             =   2625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Save"
      AccessKey       =   "S"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmModules.frx":0F4A
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   3
      Left            =   1380
      TabIndex        =   7
      Top             =   2625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Update"
      AccessKey       =   "U"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmModules.frx":16C4
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   4
      Left            =   600
      TabIndex        =   8
      Top             =   2625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&New"
      AccessKey       =   "N"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmModules.frx":1E3E
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   5
      Left            =   3720
      TabIndex        =   9
      Top             =   2625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Close"
      AccessKey       =   "C"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmModules.frx":25B8
      PicturePos      =   1
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   705
      Index           =   6
      Left            =   2160
      TabIndex        =   10
      Top             =   2625
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1244
      Caption         =   "&Delete"
      AccessKey       =   "D"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmModules.frx":2D32
      PicturePos      =   1
   End
End
Attribute VB_Name = "frmModules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmFields"

Private WithEvents oDriver As clsFormDriver
Attribute oDriver.VB_VarHelpID = -1
Private oSkin As clsFormSkin
Private bLoaded As Boolean

Dim pnIndex As Integer

Private Sub cmbDivision_Validate(Cancel As Boolean)
   Select Case cmbDivision.ListIndex
   Case 0
      oDriver.FieldValue(1) = "MC"
   Case 1
      oDriver.FieldValue(1) = "MP"
   Case Else
      oDriver.FieldValue(1) = ""
   End Select
End Sub

Private Sub cmbProduct_Validate(Cancel As Boolean)
   Select Case cmbProduct.ListIndex
   Case 0
      oDriver.FieldValue(2) = "IntegSys"
   Case 1
      oDriver.FieldValue(2) = "Telecom"
   Case 2
      oDriver.FieldValue(2) = "LRTrackr"
   Case Else
   End Select
End Sub

Private Sub cmdButton_Click(Index As Integer)
10       Dim lsOldProc As String
   
20       lsOldProc = "cmdButton_Click"
30       'On Error GoTo errProc
   
40       txtField_LostFocus pnIndex
50       Select Case Index
   Case 0
60          oDriver.RecordCancelUpdate
70       Case 1
80          oDriver.BrowseRecord
90       Case 2
100         oDriver.RecordSave
110      Case 3
120         oDriver.RecordUpdate
130      Case 4
140         oDriver.RecordNew
150      Case 5
160         Unload Me
170      Case 6
180         oDriver.RecordDelete
190      End Select

endProc:
200      Exit Sub
errProc:
210      ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
10       Dim lsOldProc As String
   
20       lsOldProc = "Form_Activate"
30       'On Error GoTo errProc
   
40        oApp.MenuName = Me.Tag
50        Me.ZOrder 0
    
60       If bLoaded = False Then
70          oDriver.RecordNew
80          oDriver.DisableTextbox 0
90          bLoaded = True
100      End If

endProc:
110      Exit Sub
errProc:
120      ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Load()
10       Dim lsOldProc As String
   
20       lsOldProc = "Form_Load"
30       'On Error GoTo errProc
   
40       CenterChildForm mdiMain, Me
   
50       bLoaded = False
   
60       Set oDriver = New clsFormDriver
70       Set oDriver.AppDriver = oApp
80       Set oDriver.MainForm = Me
   
90       Set oSkin = New clsFormSkin
100      Set oSkin.AppDriver = oApp
110      Set oSkin.Form = Me
120      oSkin.ApplySkin
   
130      oDriver.RecQuery = "SELECT * FROM System_Support_Module"
140      oDriver.BrowseQuery = "SELECT" _
                           & "  sModuleID" _
                           & ", cDivision" _
                           & ", sProdctID" _
                           & ", sModuleNm" _
                        & " FROM System_Support_Module" _
                        & " WHERE cRecdStat = " & strParm(xeRecStateActive) _
                        & " ORDER BY sModuleID"
   
150      oDriver.InitRecForm
   
160      oDriver.BrowseFTitle(0) = "Code"
170      oDriver.BrowseFTitle(1) = "Division"
171      oDriver.BrowseFTitle(1) = "Product"
172      oDriver.BrowseFTitle(1) = "Description"
   
180      oDriver.FieldFormat(0) = "@@@@@@@@@@"
190      oDriver.FieldSize(0) = Len(oDriver.FieldFormat(0))
200      oDriver.FieldStart = 3

endProc:
210      Exit Sub
errProc:
220      ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
10       Set oDriver = Nothing
20       Set oSkin = Nothing
End Sub

Private Sub oDriver_DisableOtherControl()
10       cmbDivision.Enabled = False
20       cmbProduct.Enabled = False
End Sub

Private Sub oDriver_EnableOtherControl()
10       oDriver.DisableTextbox 0
20       cmbDivision.Enabled = True
30       cmbProduct.Enabled = True
End Sub

Private Sub oDriver_InitValue()
10       Dim lsOldProc As String
   
20       lsOldProc = "oDriver_InitValue"
30       'On Error GoTo errProc
   
40       If oDriver.SetValue(0, GetNextCode("System_Support_Module", "sModuleID", True, oApp.Connection, True, oApp.BranchCode)) = False Then Exit Sub
50       oDriver.FieldReference(0) = True
60       oDriver.FieldValue(1) = "MC"
70       oDriver.FieldValue(2) = "IntegSys"
80       oDriver.FieldValue(4) = 1

90          cmbDivision.ListIndex = 0
100         cmbProduct.ListIndex = 0
End Sub

Private Sub oDriver_WillSave(Cancel As Boolean)
10       If oDriver.FieldValue(3) = "" Then
20          MsgBox "Invalid Fields Name detected!!!", vbCritical, "Warning"
30          txtField(3).SetFocus
40          Cancel = True
50       End If

   Select Case cmbDivision.ListIndex
   Case 0
      oDriver.FieldValue(1) = "MC"
   Case 1
      oDriver.FieldValue(1) = "MP"
   Case Else
      oDriver.FieldValue(1) = ""
   End Select
   
   Select Case cmbProduct.ListIndex
   Case 0
      oDriver.FieldValue(2) = "IntegSys"
   Case 1
      oDriver.FieldValue(2) = "Telecom"
   Case 2
      oDriver.FieldValue(2) = "LRTrackr"
   Case Else
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
10       oDriver.ColumnIndex = Index
20       With txtField(Index)
30          .BackColor = oApp.getColor("HT1")
40       End With
End Sub

Private Sub txtField_LostFocus(Index As Integer)
10       With txtField(Index)
20          .BackColor = oApp.getColor("EB")
30       End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
10       Dim lsOldProc As String
   
20       lsOldProc = "txtField_Validate"
30       'On Error GoTo errProc
   
40       txtField(Index).Text = txtField(Index).Text
50       Cancel = Not oDriver.ValidateField(Index)

endProc:
60       Exit Sub
errProc:
70       ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & Cancel _
                       & " )", True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
10       Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
20          Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
30             SetNextFocus
40          Case vbKeyUp
50             SetPreviousFocus
60          End Select
70       End Select
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
10       With oApp
20          .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
30          If bEnd Then
40             .xShowError
50             End
60          Else
70             With Err
80                .Raise .Number, .Source, .Description
90             End With
100         End If
110      End With
End Sub


