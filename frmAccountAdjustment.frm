VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Begin VB.Form frmAccountAdjustment 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin xrControl.xrFrame xrFrame1 
      Height          =   4200
      Index           =   1
      Left            =   1620
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   7408
      BackColor       =   12632256
      BorderStyle     =   1
      Begin xrControl.xrFrame xrFrame2 
         Height          =   3480
         Left            =   150
         Tag             =   "wt0;fb0"
         Top             =   510
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   6138
         BackColor       =   12632256
         ClipControls    =   0   'False
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   1425
            TabIndex        =   6
            Top             =   945
            Width           =   5265
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
            Index           =   2
            Left            =   1425
            TabIndex        =   5
            Top             =   615
            Width           =   2640
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   1425
            TabIndex        =   4
            Text            =   "December 20, 2010"
            Top             =   180
            Width           =   1500
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Index           =   9
            Left            =   4215
            TabIndex        =   3
            Tag             =   "ht0;ft0"
            Top             =   2730
            Width           =   2445
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Index           =   7
            Left            =   1425
            TabIndex        =   2
            Top             =   2175
            Width           =   2445
         End
         Begin VB.TextBox txtField 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Index           =   8
            Left            =   1425
            TabIndex        =   1
            Top             =   2715
            Width           =   2445
         End
         Begin VB.TextBox txtField 
            Appearance      =   0  'Flat
            Height          =   645
            Index           =   6
            Left            =   1425
            MultiLine       =   -1  'True
            TabIndex        =   0
            Top             =   1275
            Width           =   5265
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ADDRESS:"
            Height          =   195
            Index           =   2
            Left            =   585
            TabIndex        =   13
            Top             =   1005
            Width           =   825
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ACCNT NAME:"
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
            Left            =   135
            TabIndex        =   12
            Top             =   645
            Width           =   1275
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DATE:"
            Height          =   195
            Index           =   1
            Left            =   930
            TabIndex        =   11
            Top             =   210
            Width           =   480
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "A. BALANCE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   8
            Left            =   4230
            TabIndex        =   10
            Top             =   2370
            Width           =   2445
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AMOUNT IN:"
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
            Index           =   3
            Left            =   285
            TabIndex        =   9
            Top             =   2205
            Width           =   1125
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AMOUNT OUT:"
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
            Index           =   4
            Left            =   90
            TabIndex        =   8
            Top             =   2775
            Width           =   1320
         End
         Begin VB.Label lblField 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "REMARKS:"
            Height          =   195
            Index           =   5
            Left            =   570
            TabIndex        =   7
            Top             =   1305
            Width           =   840
         End
      End
      Begin VB.Label lblTransNox 
         Caption         =   "0000-00000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1620
         TabIndex        =   15
         Tag             =   "eb0;wt0"
         Top             =   180
         Width           =   1410
      End
      Begin VB.Label lblField 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transaction No."
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
         Index           =   6
         Left            =   210
         TabIndex        =   14
         Top             =   180
         Width           =   1350
      End
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   4
      Left            =   90
      TabIndex        =   18
      Top             =   1815
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmAccountAdjustment.frx":0000
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   90
      TabIndex        =   16
      Top             =   555
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmAccountAdjustment.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   90
      TabIndex        =   19
      Top             =   1815
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmAccountAdjustment.frx":0EF4
   End
   Begin xrControl.xrButton cmdButton 
      CausesValidation=   0   'False
      Height          =   600
      Index           =   3
      Left            =   90
      TabIndex        =   17
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "Searc&h"
      AccessKey       =   "h"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmAccountAdjustment.frx":166E
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   90
      TabIndex        =   20
      Top             =   1185
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmAccountAdjustment.frx":1DE8
   End
End
Attribute VB_Name = "frmAccountAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const pxeMODULENAME = "frmAccountAdjustment"

Private WithEvents oTrans As clsAdjustment
Attribute oTrans.VB_VarHelpID = -1
Private oSkin As clsFormSkin

Dim pbSearch As Boolean
Dim pnCtr As Integer, pnIndex As Integer
Dim psAccountType As String

Property Let AccountType(ByVal value As String)
   psAccountType = UCase(value)
End Property

Private Sub cmdButton_Click(Index As Integer)
   Dim lsOldProc As String
   Dim lnRep As Integer

   lsOldProc = "cmdButton_Click"
   On Error GoTo errProc

   Select Case Index
   Case 0 'New
      oTrans.NewTransaction
      clearFields
      initButton xeModeAddNew
      
      txtField(1).SetFocus
   Case 1 'Close
      Unload Me
   Case 2 'Save
      If oTrans.SaveTransaction Then
         If MsgBox("Transaction Updated Successfully!!!" & vbCrLf & vbCrLf & _
               "Post Transaction?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
            If oTrans.PostTransaction Then
               MsgBox "Transaction Posted Successfully!!!", vbInformation, "Notice"
            Else
               MsgBox "Unable to Post Transaction!!!" & vbCrLf & _
                        "Please Inform GGC SEG/SSG about this matter!!!", vbCritical, "Warning"
            End If
         End If
         Call cmdButton_Click(0) 'new
      Else
         MsgBox "Unable to Update Transaction!!!", vbCritical, "Warning"
      End If
   Case 3 'Search
      If pnIndex = 3 Or pnIndex = 9 Or pnIndex = 10 Then oTrans.Master(pnIndex) = txtField(pnIndex).Text
      txtField(pnIndex).SetFocus
   Case 4 'Cancel
      lnRep = MsgBox("Transaction is in Update Mode!!!" & vbCrLf & _
                     "Do you want to Cancel Transaction!!!", vbYesNo + vbQuestion, "Confirm")

      If lnRep = vbYes Then
         oTrans.NewTransaction
         clearFields
         initButton xeModeReady
      End If
   End Select

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index & " )", True
End Sub

Private Sub Form_Activate()
   oApp.MenuName = Me.Tag
   Me.ZOrder 0
End Sub

Private Sub Form_Load()
   Dim lsOldProc As String
   
   lsOldProc = "Form_Load"
   'On Error GoTo errProc

   CenterChildForm mdiMain, Me

   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   oSkin.ApplySkin xeFormTransaction

   Set oTrans = New clsAdjustment
   Set oTrans.AppDriver = oApp
   oTrans.AccountType = psAccountType
   oTrans.InitTransaction
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oTrans = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
      Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
         SetNextFocus
      Case vbKeyUp
         SetPreviousFocus
      End Select
   End Select
End Sub

Private Sub oTrans_MasterRetrieved(ByVal Index As Variant, ByVal value As Variant)
   Select Case Index
   Case 9
      txtField(Index).Text = Format(value, "#,##0.00")
   Case Else
      If Index > 9 Then Exit Sub
      txtField(Index).Text = value
   End Select
End Sub

Private Sub txtField_GotFocus(Index As Integer)
   With txtField(Index)
      If Index = 1 Then .Text = Format(.Text, "MM/DD/YYYY")
      .SelStart = 0
      .SelLength = Len(.Text)
      .BackColor = oApp.getColor("HT1")
   End With

   pnIndex = Index
End Sub

Private Sub clearFields()
   lblTransNox = Format(oTrans.Master("sTransNox"), "@@@@-@@@@@@@@")
   For pnCtr = 1 To txtField.Count
      Select Case pnCtr
      Case 1
         txtField(pnCtr).Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
      Case 7, 8, 9
         txtField(pnCtr).Text = Format(oTrans.Master(pnCtr), "#,##0.00")
      Case Else
         txtField(pnCtr).Text = IIf(IsNull(oTrans.Master(pnCtr)), "", oTrans.Master(pnCtr))
      End Select
   Next
End Sub

Private Sub initButton(lnStat As Integer)
   Dim lbShow As Boolean

   lbShow = IIf(lnStat = 0, False, True)
   cmdButton(0).Visible = Not lbShow
   cmdButton(1).Visible = Not lbShow

   cmdButton(2).Visible = lbShow
   cmdButton(3).Visible = lbShow
   cmdButton(4).Visible = lbShow
End Sub

Private Sub txtField_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_KeyDown"
   On Error GoTo errProc
   
   If KeyCode = vbKeyF3 Or KeyCode = vbKeyReturn Then
      With txtField(Index)
         If Index = 2 Then
            If KeyCode = vbKeyF3 Then
               oTrans.SearchMaster Index, .Text
               If .Text <> "" Then SetNextFocus
            Else
               If .Text <> "" Then oTrans.SearchMaster Index, .Text
            End If
         End If
      End With
      KeyCode = 0
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " _
                       & "  " & Index _
                       & ", " & KeyCode _
                       & ", " & Shift _
                       & " )", True
End Sub

Private Sub txtField_LostFocus(Index As Integer)
   With txtField(Index)
      If Index = 11 Then
         .BackColor = oApp.getColor("ft0")
      Else
         .BackColor = oApp.getColor("EB")
      End If
   End With
End Sub

Private Sub txtField_Validate(Index As Integer, Cancel As Boolean)
   Dim lsOldProc As String
   
   lsOldProc = "txtField_Validate"
   On Error GoTo errProc

   With txtField(Index)
      oTrans.Master(Index) = .Text
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & Index _
                       & ", " & Cancel & " )", True
End Sub

Private Sub ShowError(ByVal lsProcName As String, Optional bEnd As Boolean = False)
   With oApp
      .xLogError Err.Number, Err.Description, pxeMODULENAME, lsProcName, Erl
      If bEnd Then
         .xShowError
         End
      Else
         With Err
            .Raise .Number, .Source, .Description
         End With
      End If
   End With
End Sub

   

