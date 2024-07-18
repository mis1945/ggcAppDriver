VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash9c.ocx"
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmDayApproval 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDayApproval.frx":0000
   ScaleHeight     =   2895
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmdApproved 
      Height          =   315
      ItemData        =   "frmDayApproval.frx":24DB
      Left            =   2055
      List            =   "frmDayApproval.frx":24EE
      TabIndex        =   13
      Top             =   1845
      Width           =   2490
   End
   Begin VB.TextBox txtApprovalCode 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2055
      TabIndex        =   1
      Top             =   1530
      Width           =   2490
   End
   Begin xrControl.xrButton xrButton 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   2370
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&OK"
      AccessKey       =   "O"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      Height          =   375
      Index           =   1
      Left            =   3165
      TabIndex        =   4
      Top             =   2370
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
      BackColor       =   15720398
      ForeColor       =   4194304
      BackColorDown   =   16051167
      BorderColorFocus=   7883077
      BorderColorHover=   16379363
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   1260
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1245
      _cx             =   2196
      _cy             =   2222
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
      Profile         =   0   'False
      ProfileAddress  =   ""
      ProfilePort     =   0
      AllowNetworking =   "all"
      AllowFullScreen =   "false"
   End
   Begin VB.Label lblFormTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Day to Day"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00ECE1D0&
      Height          =   285
      Index           =   2
      Left            =   -195
      TabIndex        =   11
      Top             =   105
      Width           =   5265
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFormTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Approval"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00ECE1D0&
      Height          =   285
      Index           =   1
      Left            =   -465
      TabIndex        =   10
      Top             =   465
      Width           =   5685
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GGC - SEG '14"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   7
      Top             =   2610
      Width           =   1065
   End
   Begin VB.Label lblfield 
      BackStyle       =   0  'Transparent
      Caption         =   "Area Supervisor"
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
      Left            =   1710
      TabIndex        =   6
      Top             =   1230
      Width           =   1830
   End
   Begin VB.Label lblfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   1530
      TabIndex        =   0
      Top             =   1583
      Width           =   375
   End
   Begin VB.Label lblfield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Issuee"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1530
      TabIndex        =   2
      Top             =   1905
      Width           =   465
   End
   Begin VB.Label lblFormTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Approval"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   -255
      TabIndex        =   5
      Top             =   450
      Width           =   5235
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFormTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Day to Day"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   -195
      TabIndex        =   9
      Top             =   90
      Width           =   5265
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblfield 
      BackStyle       =   0  'Transparent
      Caption         =   "    System Approval is required by the object.   Seek  assistance  from  your                                 ."
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Index           =   2
      Left            =   1710
      TabIndex        =   12
      Top             =   840
      Width           =   2700
   End
End
Attribute VB_Name = "frmDayApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pbCancel As Boolean
Private psAppPath As String
Private poMod As New clsMainModules

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

Private Sub xrButton_Click(Index As Integer)
   pbCancel = Index = 1
   Me.Hide
End Sub

Private Sub Form_Load()
   cmdApproved.ListIndex = 0

   ShockwaveFlash1.Movie = psAppPath & "\Images\hand_key.swf"
   ShockwaveFlash1.Play
End Sub

