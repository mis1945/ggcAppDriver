VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmVehicleGasTaggingReg 
   BorderStyle     =   0  'None
   Caption         =   "Repair Tagging"
   ClientHeight    =   7860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   13335
   ShowInTaskbar   =   0   'False
   Begin xrControl.xrFrame xrFrame2 
      Height          =   5550
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   1665
      Width           =   11700
      _ExtentX        =   20638
      _ExtentY        =   9790
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtMaster 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   1635
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   60
         Width           =   2355
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5040
         Left            =   75
         TabIndex        =   4
         Top             =   435
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   8890
         _Version        =   393216
      End
      Begin VB.Shape Shape3 
         Height          =   360
         Index           =   0
         Left            =   8490
         Top             =   45
         Width           =   2160
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8535
         TabIndex        =   16
         Tag             =   "eb0;et0"
         Top             =   75
         Width           =   2070
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Transaction No"
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
         Index           =   5
         Left            =   225
         TabIndex        =   14
         Tag             =   "wt0;fb0"
         Top             =   90
         Width           =   1335
      End
   End
   Begin xrControl.xrFrame xrFrame1 
      Height          =   1080
      Left            =   90
      Tag             =   "wt0;fb0"
      Top             =   555
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   1905
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.TextBox txtPlate 
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
         Height          =   330
         Left            =   6225
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   570
         Width           =   4980
      End
      Begin VB.TextBox txtSupplier 
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
         Height          =   330
         Index           =   2
         Left            =   6225
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   210
         Width           =   4980
      End
      Begin VB.TextBox txtDateThru 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1590
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   540
         Width           =   2355
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1590
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   195
         Width           =   2355
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Plate No"
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
         Index           =   3
         Left            =   5100
         TabIndex        =   11
         Tag             =   "wt0;fb0"
         Top             =   570
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Supplier"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5100
         TabIndex        =   10
         Tag             =   "wt0;fb0"
         Top             =   240
         Width           =   1020
      End
      Begin VB.Line Line1 
         X1              =   4635
         X2              =   4635
         Y1              =   165
         Y2              =   870
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Date Thr:"
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
         Index           =   1
         Left            =   525
         TabIndex        =   9
         Tag             =   "wt0;fb0"
         Top             =   570
         Width           =   990
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Date Fr:"
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
         Index           =   0
         Left            =   525
         TabIndex        =   8
         Tag             =   "wt0;fb0"
         Top             =   240
         Width           =   990
      End
      Begin VB.Shape Shape1 
         Height          =   870
         Left            =   135
         Top             =   105
         Width           =   11355
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   2
      Left            =   11985
      TabIndex        =   3
      Top             =   1845
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
      Picture         =   "frmVehicleGasTaggingReg.frx":0000
   End
   Begin xrControl.xrFrame xrFrame3 
      Height          =   525
      Left            =   135
      Tag             =   "wt0;fb0"
      Top             =   7230
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   926
      BackColor       =   12632256
      ClipControls    =   0   'False
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   8460
         TabIndex        =   13
         Tag             =   "wt0;fb0"
         Top             =   15
         Width           =   2250
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "TOTAL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   7395
         TabIndex        =   12
         Tag             =   "wt0;fb0"
         Top             =   120
         Width           =   990
      End
      Begin VB.Shape Shape2 
         Height          =   870
         Left            =   9855
         Top             =   3225
         Width           =   10605
      End
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   1
      Left            =   11985
      TabIndex        =   1
      Top             =   585
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
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
      Picture         =   "frmVehicleGasTaggingReg.frx":077A
   End
   Begin xrControl.xrButton cmdButton 
      Height          =   600
      Index           =   0
      Left            =   11985
      TabIndex        =   2
      Top             =   1215
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1058
      Caption         =   "&Confirm"
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
      Picture         =   "frmVehicleGasTaggingReg.frx":0EF4
   End
End
Attribute VB_Name = "frmVehicleGasTaggingReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'sheryl 12/01/2014 4:34 pm
'start creating this form

Option Explicit

Private Const pxeMODULENAME = "frmVehicleGasTaggingReg"
Private p_oappdriver As clsAppDriver
Private oSkin As clsFormSkin
Private oTrans As clsVehicleGasTagging

Dim lsSQL As String
Dim loRec As Recordset

Private Sub InitGrid()
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      .Cols = 9
      
      .TextMatrix(0, 1) = "Trans No"
      .TextMatrix(0, 2) = "Date"
      .TextMatrix(0, 3) = "Plate"
      .TextMatrix(0, 4) = "Company"
      .TextMatrix(0, 5) = "Refer No"
      .TextMatrix(0, 6) = "Liter/s"
      .TextMatrix(0, 7) = "Paid Amt"
      .TextMatrix(0, 8) = "Tag"
      
      .Row = 0
      
      .ColWidth(0) = 500
      .ColWidth(1) = 1400
      .ColWidth(2) = 1300
      .ColWidth(3) = 1000
      .ColWidth(4) = 3070
      .ColWidth(5) = 1200
      .ColWidth(6) = 1200
      .ColWidth(7) = 1200
      .ColWidth(8) = 500
      
      For lnCtr = 1 To .Cols - 1
         .Col = lnCtr
         .CellAlignment = 3
         .CellFontBold = True
      Next
      
   End With
      
End Sub

Private Sub cmdButton_Click(Index As Integer)
   Dim lnCtr As Integer
   
   With MSFlexGrid1
      Select Case Index
      Case 0 'Confirm
         If txtMaster.Text <> "" Then
            If oTrans.postTransaction(oTrans.Master("sTransNox")) = True Then
               MsgBox "Transaction POST Successfully!", vbInformation, "Info"
               Label2.Caption = "POSTED"
            End If
         End If
      Case 1 'Browse
         If oTrans.SearchTransaction() Then
            LoadRecord
         End If
      Case 2 'Close
         Unload Me
      End Select
   End With
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
   
   Set oTrans = New clsVehicleGasTagging
   Set oTrans.AppDriver = oApp
   
   oTrans.InitTransaction
   oTrans.NewTransaction
   
   Set oSkin = New clsFormSkin
   Set oSkin.AppDriver = oApp
   Set oSkin.Form = Me
   
   oSkin.ApplySkin xeFormTransMaintenance
   txtDateFrom.Text = Format(DateAdd("d", 1, DateAdd("m", -1, oApp.ServerDate)), "MMMM DD, YYYY")
   txtDateThru.Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   
   Call InitGrid
   Call InitField
   
endProc:
   Exit Sub
errProc:
   ShowError lsOldProc & "( " & " )", True

End Sub

Private Sub InitField()
     
   With MSFlexGrid1
      .TextMatrix(1, 1) = ""
      .TextMatrix(1, 2) = ""
      .TextMatrix(1, 3) = ""
      .TextMatrix(1, 4) = ""
      .TextMatrix(1, 5) = "0.00"
      .TextMatrix(1, 6) = "0.00"
      .TextMatrix(1, 7) = "0.00"
      .TextMatrix(1, 8) = "No"
      
      .Row = 1
      .Col = 1
      .ColSel = .Cols - 1
   End With
   
   txtPlate.Text = ""
   txtSupplier(2).Text = ""
   txtMaster.Text = oTrans.Master(0)
   lblTotal.Caption = "0.00"
   lblTotal.ForeColor = &HFF&
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oSkin = Nothing
   Set oTrans = Nothing
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

Private Sub LoadRecord()
   Dim lnCtr As Integer

   With MSFlexGrid1
      .Rows = oTrans.ItemCount + 1

      For lnCtr = 1 To oTrans.ItemCount
         .TextMatrix(lnCtr, 0) = lnCtr
         .TextMatrix(lnCtr, 1) = oTrans.Detail(lnCtr - 1, ("sTransNox"))
         .TextMatrix(lnCtr, 2) = oTrans.Detail(lnCtr - 1, ("dTransact"))
         .TextMatrix(lnCtr, 3) = oTrans.Detail(lnCtr - 1, ("sPlateNox"))
         .TextMatrix(lnCtr, 4) = oTrans.Detail(lnCtr - 1, ("sCompnyNm"))
         .TextMatrix(lnCtr, 5) = oTrans.Detail(lnCtr - 1, ("sReferNox"))
         .TextMatrix(lnCtr, 6) = oTrans.Detail(lnCtr - 1, ("nNoLiters"))
         .TextMatrix(lnCtr, 7) = oTrans.Detail(lnCtr - 1, ("nPaidAmtx"))
         .TextMatrix(lnCtr, 8) = "Yes"
       Next
   End With
   
      txtSupplier(2).Text = oTrans.Master(2)
      txtMaster.Text = oTrans.Master(0)
      lblTotal.Caption = Format(oTrans.Master(4), "#,##0.00")
      Label2.Caption = Format(TransStat(oTrans.Master("cTranStat")), ">")
End Sub

Property Get DateFrom() As Date
   DateFrom = CDate(txtDateFrom.Text)
End Property

Property Get DateThru() As Date
   DateThru = CDate(txtDateThru.Text)
End Property

Private Sub txtDateFrom_LostFocus()
   If IsDate(CDate(txtDateFrom.Text)) Then
      txtDateFrom.Text = Format(txtDateFrom.Text, "MMMM DD, YYYY")
   End If
End Sub

Private Sub txtDateFrom_Validate(Cancel As Boolean)
   If Not IsDate(txtDateFrom.Text) Then
      txtDateFrom.Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   End If
End Sub

Private Sub txtDateThru_LostFocus()
   If IsDate(CDate(txtDateThru.Text)) Then
      txtDateThru.Text = Format(txtDateThru.Text, "MMMM DD, YYYY")
   End If
End Sub

Private Sub txtDateThru_Validate(Cancel As Boolean)
   If Not IsDate(txtDateThru.Text) Then
      txtDateThru.Text = Format(oApp.ServerDate, "MMMM DD, YYYY")
   End If
End Sub

Private Sub txtPlate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
      If txtPlate.Text <> "" Then
         SearchPlate True
      End If
   End If
End Sub

Private Sub txtSupplier_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

   If KeyCode = vbKeyF3 Then
      oTrans.SearchTransaction txtSupplier(2).Text
      txtSupplier(2).Text = oTrans.Master(2)
   End If
End Sub

