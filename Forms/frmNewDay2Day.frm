VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmNewDay2Day 
   BorderStyle     =   0  'None
   Caption         =   "Day-to-day Transaction Opening"
   ClientHeight    =   2925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmNewDay2Day.frx":0000
   ScaleHeight     =   2925
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin xrControl.xrButton cmdButton 
      Height          =   360
      Index           =   1
      Left            =   3255
      TabIndex        =   4
      Top             =   2145
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   635
      Caption         =   "Ok"
      AccessKey       =   "Ok"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtApproval 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2385
      TabIndex        =   3
      Top             =   1665
      Width           =   2115
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   2385
      TabIndex        =   2
      Top             =   1335
      Width           =   2115
   End
   Begin xrControl.xrButton cmdButton 
      Cancel          =   -1  'True
      Height          =   360
      Index           =   0
      Left            =   1995
      TabIndex        =   5
      Top             =   2145
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   635
      Caption         =   "Cancel"
      AccessKey       =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GMC - SEG '14"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   2550
      Width           =   1080
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter the date and approval code of transaction that you like to encode."
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   1170
      TabIndex        =   7
      Top             =   810
      Width           =   3045
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Day-To-Day Transaction Entry"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Index           =   0
      Left            =   300
      TabIndex        =   6
      Top             =   345
      Width           =   4020
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Approval Code"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1170
      TabIndex        =   1
      Top             =   1725
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1875
      TabIndex        =   0
      Top             =   1395
      Width           =   345
   End
End
Attribute VB_Name = "frmNewDay2Day"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pbCancel As Boolean
Private pbFocus As Boolean
Private poMod As New clsMainModules

Property Get Cancel() As Boolean
   Cancel = pbCancel
End Property

Private Sub cmdButton_Click(Index As Integer)
   If Index = 1 Then
      If Not IsDate(txtDate.Text) Then
         MsgBox "Invalid Day-To-Day Transaction Date detected!" & vbCrLf & _
                "Please enter a valid date and try again...", vbInformation + vbOKOnly, "Day-To-Day Date Validation"
         GoTo endProc
      End If
   End If
   
   pbCancel = Index = 0
   Me.Hide

endProc:
   Exit Sub
End Sub

Private Sub Form_Initialize()
   pbCancel = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      If KeyCode <> vbKeyReturn And pbFocus Then Exit Sub
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         poMod.SetNextFocus
      Case vbKeyUp
         poMod.SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub Form_Terminate()
   Set poMod = Nothing
End Sub
