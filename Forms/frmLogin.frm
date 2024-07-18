VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9000
   ControlBox      =   0   'False
   ForeColor       =   &H8000000D&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDate 
      Height          =   350
      Left            =   5900
      TabIndex        =   2
      Top             =   3225
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   5900
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2800
      Width           =   2460
   End
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00FFFFFF&
      Height          =   350
      Left            =   5900
      TabIndex        =   0
      Top             =   2400
      Width           =   2460
   End
   Begin xrControl.xrButton xrButton 
      Height          =   375
      Index           =   0
      Left            =   6100
      TabIndex        =   4
      Top             =   4580
      Width           =   1260
      _ExtentX        =   2223
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
      BackColor       =   33023
      ForeColor       =   16777215
      BackColorDown   =   16051167
      BorderColorFocus=   7883077
      BorderColorHover=   16379363
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4020
      Width           =   3315
   End
   Begin xrControl.xrButton xrButton 
      Height          =   375
      Index           =   1
      Left            =   7500
      TabIndex        =   5
      Top             =   4580
      Width           =   1260
      _ExtentX        =   2223
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
      BackColor       =   33023
      ForeColor       =   16777215
      BackColorDown   =   16051167
      BorderColorFocus=   7883077
      BorderColorHover=   16379363
   End
   Begin xrControl.xrButton xrButton 
      Height          =   285
      Index           =   2
      Left            =   8640
      TabIndex        =   10
      Top             =   45
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      Caption         =   " X"
      AccessKey       =   " X"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      ForeColor       =   4210752
      BackColorDown   =   16051167
      BorderColorFocus=   7883077
      BorderColorHover=   16379363
   End
   Begin VB.Image imgDate 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   5475
      Picture         =   "frmLogin.frx":1D98F
      Stretch         =   -1  'True
      Top             =   3225
      Width           =   345
   End
   Begin VB.Label lblFaxNo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fax No: (075) 522 9275"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   90
      TabIndex        =   9
      Top             =   5580
      Width           =   4785
   End
   Begin VB.Label lblTelNo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tel No: (075) 522 1085; 522 1097"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   90
      TabIndex        =   8
      Top             =   5310
      Width           =   4785
   End
   Begin VB.Label lblAddress 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Guanzon Bldg., Perez Blvd., Dagupan City"
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   90
      TabIndex        =   7
      Top             =   5025
      Width           =   4785
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Guanzon Merchandising Corporation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   90
      TabIndex        =   6
      Top             =   4710
      Width           =   4785
   End
End
Attribute VB_Name = "frmLogin"
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

Private Sub Combo1_GotFocus()
   pbFocus = True
End Sub

Private Sub Combo1_LostFocus()
   pbFocus = False
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

Private Sub xrButton_Click(Index As Integer)
   If IsDate(txtDate.Text) Then
      pbCancel = Index = 1
      Me.Hide
   Else
      MsgBox "Invalid System Date detected", vbOKOnly, "Warning"
   End If
End Sub
