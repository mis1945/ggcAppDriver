VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmProgress 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   2670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6675
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmProgress.frx":0000
   ScaleHeight     =   2670
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox shpProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H007A3A14&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   0
      Left            =   2625
      ScaleHeight     =   150
      ScaleWidth      =   3885
      TabIndex        =   5
      Top             =   930
      Width           =   3885
   End
   Begin VB.PictureBox shpProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H007A3A14&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   150
      Index           =   1
      Left            =   2625
      ScaleHeight     =   150
      ScaleWidth      =   3885
      TabIndex        =   4
      Top             =   1545
      Width           =   3885
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   3885
      TabIndex        =   1
      Top             =   2130
      Width           =   1125
   End
   Begin MSComCtl2.Animation aniPiston 
      Height          =   1815
      Left            =   105
      TabIndex        =   0
      Top             =   645
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   3201
      _Version        =   393216
      BackColor       =   1842204
      FullWidth       =   160
      FullHeight      =   121
   End
   Begin VB.Label lblRemarks 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   0
      Left            =   2610
      TabIndex        =   7
      Top             =   1110
      Width           =   3915
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00800000&
      Height          =   180
      Left            =   2610
      Top             =   915
      Width           =   3915
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   180
      Left            =   2610
      Top             =   1530
      Width           =   3915
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2640
      X2              =   6555
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   2640
      X2              =   6510
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Label lblProcess 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Processing..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Top             =   180
      Width           =   1815
   End
   Begin VB.Label lblProcess 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Processing..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   0
      Left            =   2610
      TabIndex        =   3
      Top             =   225
      Width           =   1815
   End
   Begin VB.Label lblRemarks 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   1
      Left            =   2610
      TabIndex        =   6
      Top             =   1725
      Width           =   3915
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
                           
' Used to support captionless drag
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const HWND_TOPMOST = -&H1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private p_nPriMaxValue As Long
Private p_nSecMaxValue As Long
Private p_bCancelled As Boolean

Private pnPriInterval As Long
Private pnSecInterval As Long
Private pnPriProgress As Long
Private pnSecProgress As Long

Private pnCtr As Long
Private Const MaxProgress = 3915

Private Sub Form_Load()
   pnPriProgress = 0
   pnSecProgress = 0
   
   shpProgress(0).Width = 0
   shpProgress(1).Width = 0

   aniPiston.Open App.Path & "\piston.avi"
   aniPiston.Play
End Sub

Function MoveProgress()
   pnSecProgress = pnSecProgress + 1
   
   shpProgress(0).Width = Fix(pnSecProgress / p_nSecMaxValue * MaxProgress)
   DoEvents
   
   If pnSecProgress = p_nSecMaxValue Then
      If pnPriProgress < p_nPriMaxValue Then
         pnPriProgress = pnPriProgress + 1
         DoEvents

         shpProgress(1).Width = Fix(pnPriProgress / p_nPriMaxValue * MaxProgress)
         DoEvents
      End If
      pnSecProgress = 0
   End If
   
   DoEvents
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   ' Automatically allow user to drag using any portion of form, not just titlebar,
   '  when user depresses left mousebutton. Useful for captionless forms.
   If Button = vbLeftButton Then
      DoEvents
      Call ReleaseCapture
      DoEvents
      Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
      DoEvents
   End If
End Sub

Property Get SecondaryMaxValue() As Long
    SecondaryMaxValue = p_nSecMaxValue
End Property

Property Let SecondaryMaxValue(ByVal Value As Long)
   p_nSecMaxValue = Value
   pnSecProgress = 0
End Property

Property Get PrimaryMaxValue() As Long
   PrimaryMaxValue = p_nPriMaxValue
End Property

Property Let PrimaryMaxValue(ByVal Value As Long)
   p_nPriMaxValue = Value
   pnPriProgress = 0
End Property

Property Let ProgressStatus(ByVal Value As String)
   lblProcess(0).Caption = Value
   lblProcess(1).Caption = Value
End Property

Property Let PrimaryRemarks(ByVal Value As String)
   lblRemarks(1).Caption = Value
End Property

Property Let SecondaryRemarks(ByVal Value As String)
   lblRemarks(0).Caption = Value
End Property

Property Get Cancelled() As Boolean
   Cancelled = p_bCancelled
End Property

Private Sub cmdCancel_Click()
   p_bCancelled = True
   Me.Hide
End Sub
