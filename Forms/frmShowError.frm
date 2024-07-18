VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmShowError 
   BackColor       =   &H003C682F&
   BorderStyle     =   0  'None
   Caption         =   "BugTracker - ShowError"
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   8535
   Icon            =   "frmShowError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmShowError.frx":030A
   ScaleHeight     =   3885
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox ErrorInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1515
      Left            =   1575
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmShowError.frx":E5EC
      Top             =   1140
      Width           =   6810
   End
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   135
      TabIndex        =   6
      Top             =   2970
      Width           =   8250
   End
   Begin VB.TextBox ErrorDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   1575
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   495
      Width           =   6810
   End
   Begin VB.TextBox ErrorNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1575
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   90
      Width           =   1545
   End
   Begin xrControl.xrButton xrButton1 
      Height          =   615
      Index           =   0
      Left            =   7470
      TabIndex        =   8
      Top             =   3120
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   1085
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
      Picture         =   "frmShowError.frx":E5F2
      BackColor       =   15720398
      ForeColor       =   4194304
      BackColorDown   =   16051167
      BorderColorFocus=   7883077
      BorderColorHover=   16379363
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmShowError.frx":ED6C
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   135
      TabIndex        =   7
      Top             =   3150
      Width           =   7170
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Other Information:"
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
      Height          =   375
      Left            =   45
      TabIndex        =   2
      Top             =   1170
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
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
      Height          =   375
      Left            =   45
      TabIndex        =   1
      Top             =   540
      Width           =   1455
   End
   Begin VB.Label lblErrorNo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Error No:"
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
      Height          =   285
      Left            =   45
      TabIndex        =   0
      Top             =   135
      Width           =   1455
   End
End
Attribute VB_Name = "frmShowError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Private Sub xrButton1_Click()
'    Unload Me
'End Sub

Private Sub xrButton1_Click(Index As Integer)
   Unload Me
End Sub
