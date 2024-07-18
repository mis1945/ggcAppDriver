VERSION 5.00
Object = "{34A378CB-112C-461B-94E8-02D25370A1CE}#8.1#0"; "xrControl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmLookUp1 
   BorderStyle     =   0  'None
   Caption         =   "Look Up Table"
   ClientHeight    =   7275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLookUp1.frx":0000
   ScaleHeight     =   7275
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   240
      Left            =   105
      TabIndex        =   13
      Top             =   6900
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00253315&
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Tag             =   "et0;eb0"
      Top             =   1650
      Width           =   3795
   End
   Begin VB.ComboBox cmbSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5790
      TabIndex        =   1
      Tag             =   "et0;eb0"
      Text            =   "Sort Key"
      Top             =   1650
      Width           =   1920
   End
   Begin xrControl.xrButton xrButton1 
      Height          =   330
      Index           =   2
      Left            =   6705
      TabIndex        =   12
      Top             =   1200
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   582
      Caption         =   "F3-Searc&h"
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
   End
   Begin xrControl.xrButton xrButton1 
      Height          =   330
      Index           =   1
      Left            =   6705
      TabIndex        =   11
      Top             =   840
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   582
      Caption         =   "ESC-&Cancel"
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
   End
   Begin xrControl.xrButton xrButton1 
      Default         =   -1  'True
      Height          =   330
      Index           =   0
      Left            =   6705
      TabIndex        =   10
      Top             =   480
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   582
      Caption         =   "F5-&Load"
      AccessKey       =   "L"
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
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4890
      Left            =   60
      TabIndex        =   9
      Tag             =   "et0;eb0;et0;fb0"
      Top             =   2280
      Visible         =   0   'False
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   8625
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   8388608
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GMC-Software Engineering Group"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   5
      Left            =   1920
      TabIndex        =   8
      Tag             =   "hb1"
      Top             =   1170
      Width           =   2115
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Tag             =   "hb2"
      Top             =   435
      Width           =   1875
   End
   Begin VB.Image Image4 
      Height          =   1455
      Left            =   120
      Picture         =   "frmLookUp1.frx":361B
      Top             =   555
      Width           =   1500
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000005&
      Height          =   1515
      Left            =   90
      Top             =   525
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   2
      Left            =   1935
      TabIndex        =   6
      Tag             =   "hb1"
      Top             =   1455
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fields"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   3
      Left            =   5790
      TabIndex        =   5
      Tag             =   "hb1"
      Top             =   1455
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2003 and beyond"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   4
      Left            =   1920
      TabIndex        =   4
      Tag             =   "hb1"
      Top             =   960
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   1905
      TabIndex        =   3
      Tag             =   "1-1"
      Top             =   465
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 2.00"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   6
      Left            =   1920
      TabIndex        =   2
      Tag             =   "hb1"
      Top             =   750
      Width           =   765
   End
End
Attribute VB_Name = "frmLookUp1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Rex S. Adversalo
' XerSys Computing
' Canaoalan, Binmaley, Pangasinan
'
' LookUp (RecordSet) v1.5
'     Display lookup table and allows user to select from a list.
'     Properties:
'        RowSource = Recordset that contains the selection
'        Column    = a string or array of array of field name the will appear on
'                    lookup table
'        ColHead   = a string or array of string of column heading
'        SortKey   = the default column to be use as sort key
'
' Copyright 2002 and beyond
' All Rights Reserved
'
' ººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All rights reserved. No part of this  €€  This Software is Owned by        €
' €  software may be reproduced or trans-  €€                                   €
' €  mitted in any  form or by any means,  €€    GUANZON MERCHANDISING CORP.    €
' €  electronic or mechanical,  including  €€     Guanzon Bldg. Perez Blvd.     €
' €  recording, or by information storage  €€           Dagupan City            €
' €  and retrieval systems, without prior  €€  Tel No. 522-1085 ; 522-0863      €
' €  written permission from the author.   €€                                   €
' ººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ================================================================================
'  XerSys [ 01/29/2007 09:44 am ]
'     This will be the lookup when the Paramater is a SQL Statement. Separate it
'        from the lookup with a recordset parameter to allow a faster disply of
'        the lookup with a recordset paramater.
'     This is the lookup that has search feature
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

Private Const xeColMargin As Integer = 50
Private Const pxeCharWidth As Integer = 110
Private Const xeScrollBar As Integer = 300
Private Const xeMaxItem As Integer = 20
Private Const xeMaxRecd As Integer = 32767

Private WithEvents p_oLookup As ADODB.Recordset
Attribute p_oLookup.VB_VarHelpID = -1
Private p_oAppDrivr As clsAppDriver
Private p_oSkin As clsFormSkin
Private p_oMod As clsMainModules

Private p_sSQLQuery As String
Private p_asFldName() As String
Private p_asColName() As String
Private p_asColHead() As String
Private p_asColPict() As String
Private p_acColType() As String
Private p_anColWdth() As Integer
Private p_sColHead As String
Private p_sColName As Variant
Private p_sFldName As String
Private p_sColPict As Variant
Private p_nSearch As Integer
Private p_bSearch As Boolean

Private p_bSelected As Boolean
Private p_bRowSource As Boolean
Private p_bDisplayd As Boolean

Private pnCtr As Integer
Private pnInterval As Integer
Private pnProgress As Integer
Private pbProgress As Boolean
Private pbFocus As Boolean

Property Set AppDriver(oAppDriver As clsAppDriver)
10       Set p_oAppDrivr = oAppDriver
End Property

Property Let SQLSource(Source As String)
10       p_sSQLQuery = Source
End Property

Property Let FldTitle(Title As String)
   ' the column heading of the lookup
   ' the heading item/s must correspond to the order of the column
   '     of the rec source. This will be the only visible column
   '     description that identifies its content
   
10       p_sColHead = Title
End Property

Property Let FldName(Name As String)
   ' added this property to customize the # of column and the order
   '     of column to be displayed
10       p_sColName = Name
End Property

Property Let FldCriteria(Value As String)
   ' added this property to implement the runtime filtering of recordset
10       p_sFldName = Value
End Property

Property Let FldFormat(Format As String)
   ' added this property to allow field formating
   
10       p_sColPict = Format
End Property

Property Get SelectedItem() As Variant
   ' the selected item
   
10       If p_bSelected Then
20          SelectedItem = getSelectedItem()
30       Else
40          SelectedItem = Empty
50       End If
End Property

Private Sub cmbSearch_GotFocus()
10       pbFocus = True
20       With cmbSearch
30          .Tag = .ListIndex
40       End With
End Sub

Private Sub cmbSearch_LostFocus()
   ' this will allow the user to modify the search key
10       With cmbSearch
20          If .ListIndex = -1 Or .ListIndex = .Tag Then Exit Sub
30       End With
   
40       If p_bDisplayd Then SortList
50       pbFocus = False
End Sub

Private Sub Form_Activate()
10       initLookUp
End Sub

Private Sub Form_Initialize()
10       Set p_oSkin = New clsFormSkin
20       Set p_oMod = New clsMainModules
   
30       p_bSelected = False
40       p_bDisplayd = False
50       p_bSearch = False
60       p_nSearch = 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
10       Select Case KeyCode
   Case vbKeyReturn, vbKeyUp, vbKeyDown
20          If pbFocus Then Exit Sub
30          If KeyCode <> vbKeyReturn Then Exit Sub
40          Select Case KeyCode
      Case vbKeyReturn, vbKeyDown
50             p_oMod.SetNextFocus
60          Case vbKeyUp
70             p_oMod.SetPreviousFocus
80          End Select
90       Case vbKeyF5
100         Call xrButton1_Click(0)
110      Case vbKeyF3
120         Call xrButton1_Click(2)
130      Case vbKeyEscape
140         Call xrButton1_Click(1)
150      End Select
End Sub

Private Sub Form_Load()
10       Set p_oSkin.Form = Me
20       Set p_oSkin.AppDriver = p_oAppDrivr
30       p_oSkin.ApplySkin xeFormQuickSearch
End Sub

Private Sub Form_Unload(Cancel As Integer)
10       Set p_oLookup = Nothing
20       Set p_oMod = Nothing
30       Set p_oSkin = Nothing
End Sub

' retrieves the table and set the field property
Private Sub initLookUp()
10       Dim lsProcName As String
20       Dim lnTotWidth As Long
30       Dim lnAlignment As Long
   
40       lsProcName = "initLookUp"
50       On Error GoTo errProc
   
60       Set p_oLookup = New Recordset
70       Debug.Print p_sSQLQuery
80       p_oLookup.Open p_oMod.AddCondition(p_sSQLQuery, "0 = 1"), _
                  p_oAppDrivr.Connection, , , adCmdText
   
   ' check if client passed a field filter
90       If p_sColName <> "" Then
100         p_asColName = Split(p_sColName, "»", , vbTextCompare)
110      Else
      ' if not include all fields in the lookup
120         ReDim p_asColName(p_oLookup.Fields.Count - 1) As String
130         For pnCtr = 0 To UBound(p_asColName)
140            p_asColName(pnCtr) = p_oLookup.Fields(pnCtr).Name
150         Next
160      End If
   
170      If p_sColHead <> Empty Then
180         p_asColHead = Split(p_sColHead, "»", -1, vbTextCompare)
190      Else
200         ReDim p_asColHead(UBound(p_asColName)) As String
210         For pnCtr = 0 To UBound(p_asColName)
220            p_asColHead(pnCtr) = p_asColName(pnCtr)
230         Next
240      End If
   
   ' after retrieving the field name, create a field criteria
   ' to be used in creating sql statement at runtime
250      If p_sFldName <> Empty Then
260         p_asFldName = Split(p_sFldName, "»", , vbTextCompare)
270      Else
280         ReDim p_asFldName(UBound(p_asColName)) As String
290         For pnCtr = 0 To UBound(p_asColName)
300            p_asFldName(pnCtr) = p_asColName(pnCtr)
310         Next
320      End If

   ' after retrieving the column, set the type and the width
330      ReDim p_acColType(UBound(p_asColName))
340      ReDim p_asColPict(UBound(p_asColName))
350      ReDim p_anColWdth(UBound(p_asColName))
360      For pnCtr = 0 To UBound(p_asColName)
370         p_anColWdth(pnCtr) = Len(p_asColHead(pnCtr))
380         p_asColPict(pnCtr) = "@"
      
390         Select Case p_oLookup(p_asColName(pnCtr)).Type
      Case 129, 130, 202, 200    ' string
400            p_acColType(pnCtr) = "s"
410         Case 2, 3, 11, 17, 72      ' numeric without decimal point
420            p_acColType(pnCtr) = "n"
430         Case 4, 5, 6, 131          ' numeric with decimal point
440            p_acColType(pnCtr) = "l"
450         Case 135                   ' datetime
460            p_acColType(pnCtr) = "d"
470         End Select
480      Next
490      If p_sColPict <> Empty Then p_asColPict = Split(p_sColPict, "»", , vbTextCompare)
   
500      cmbSearch.Clear
510      For pnCtr = LBound(p_asColHead) To UBound(p_asColHead)
520         If p_asColHead(pnCtr) = "" Then Exit For
530         cmbSearch.AddItem (p_asColHead(pnCtr))
      
      ' assign the length of the headers as the max width of the columns
540         p_anColWdth(pnCtr) = Len(Trim(p_asColHead(pnCtr)))
550      Next
560      cmbSearch.ListIndex = IIf(UBound(p_asColHead) > 1, 1, 0)
   
570      With MSFlexGrid1
580         .Cols = UBound(p_asColName) + 1
590         .Rows = 2
600         .Row = 0
      
610         For pnCtr = LBound(p_asColName) To UBound(p_asColName)
         ' get the appropraite alignment of each field
620            If p_acColType(pnCtr) = "s" Then
630               lnAlignment = flexAlignLeftTop
640            Else
650               lnAlignment = flexAlignRightTop
660            End If
         
670            .Col = pnCtr
680            .CellAlignment = lvwColumnLeft
690            .CellFontBold = True
         
700            .TextMatrix(0, pnCtr) = p_asColHead(pnCtr)
710            .ColAlignment(pnCtr) = lnAlignment
720         Next
730      End With
   
endProc:

740      Exit Sub
errProc:
750      ShowError lsProcName & "( " & " )"
End Sub

' assigns the contents of the recordset to the grid
Public Function LoadList() As Boolean
10       Dim lvValue As Variant
20       Dim lnAlignment As Integer
30       Dim lanColWidth() As Long
40       Dim lnTotWidth As Long
50       Dim lsProcName As String
   
60       lsProcName = "LoadList"
70       LoadList = False
80       On Error GoTo errProc
   
   ' reformat the flexgrid
90       With MSFlexGrid1
      ' check if there's a record to display
100         If p_oLookup.RecordCount = 0 Then
110            .Visible = False
         
120            p_bDisplayd = False
130            GoTo endProc
140         End If
   
150         If p_oLookup.RecordCount > xeMaxRecd Then
160            MsgBox "Search Record Result Exceeds The Maximum Allowable Record Display!!!" & _
               vbCrLf & "Please Limit Your Selection by Specifying More Detailed Info!!!", vbCritical, "Warning"
170            GoTo endProc
180         End If
      
190         .Visible = True
200         .Rows = p_oLookup.RecordCount + 1
210         p_oLookup.MoveFirst
      
220         showProgress .Rows + 1
230         .Row = 0
      
240         ReDim lanColWidth(UBound(p_asColName))
250         Do Until p_oLookup.EOF
260            .Row = .Row + 1
270            For pnCtr = 0 To UBound(p_asColName)
280               lvValue = p_oLookup(p_asColName(pnCtr))
290               If IsNull(p_oLookup(p_asColName(pnCtr))) Then lvValue = setDefaultValue(pnCtr)
300               .TextMatrix(.Row, pnCtr) = Format(lvValue, p_asColPict(pnCtr))
            
310               If Len(Trim(.TextMatrix(.Row, pnCtr))) > lanColWidth(pnCtr) Then
320                  lanColWidth(pnCtr) = Len(Trim(.TextMatrix(.Row, pnCtr)))
330               End If
340            Next
         
350            p_oLookup.MoveNext
360         Loop

      ' after fetching all record to the grid, adjust the column width
370         lnTotWidth = 0
380         For pnCtr = 0 To .Cols - 1
390            If lanColWidth(pnCtr) > p_anColWdth(pnCtr) Then p_anColWdth(pnCtr) = lanColWidth(pnCtr)
400            .ColWidth(pnCtr) = p_anColWdth(pnCtr) * pxeCharWidth
410            lnTotWidth = lnTotWidth + p_anColWdth(pnCtr)
420         Next
      
430         If .Rows > xeMaxItem Then
440            If (lnTotWidth * pxeCharWidth) < .Width - xeScrollBar Then
450               For pnCtr = 0 To .Cols - 1
460                  .ColWidth(pnCtr) = p_anColWdth(pnCtr) / lnTotWidth * (.Width - xeScrollBar) - xeColMargin
470               Next
480            End If
490         Else
500            If (lnTotWidth * pxeCharWidth) < .Width - xeColMargin Then
510               For pnCtr = 0 To .Cols - 1
520                  .ColWidth(pnCtr) = p_anColWdth(pnCtr) / lnTotWidth * .Width - xeColMargin
530               Next
540            End If
550         End If
      
560         .Row = 1
570         .Col = 0
580         .ColSel = .Cols - 1
590      End With
600      hideProgress
610      p_bDisplayd = True
620      LoadList = True
   
endProc:

630      Exit Function
errProc:
640      ShowError lsProcName & "( " & " )"
End Function

Private Function setDefaultValue(lnCol As Integer) As Variant
10       Select Case p_acColType(lnCol)
   Case "n"
20          setDefaultValue = 0
30       Case "l"
40          setDefaultValue = 0#
50       Case "d"
60          setDefaultValue = "01/01/1900"
70       Case Else
80          setDefaultValue = ""
90       End Select
End Function

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
   'Remarks: This procedure only exists to trap a delete key, which irritatingly,
   '         does not trigger a KeyPress event
   '
10       Dim lsSearchOn As String          'current string to search on

20       On Error Resume Next
   
30       If p_bDisplayd = False Then Exit Sub
   
   'Check if we're dealing with a Delete key
40       If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Or _
          KeyCode = vbKeyPageDown Or KeyCode = vbKeyPageUp Then
50          MSFlexGrid1.SetFocus
60          Exit Sub
70       ElseIf KeyCode <> vbKeyDelete Then
80          Exit Sub
90       End If
   
   'The delete key was pressed; decide what to search on
100      lsSearchOn = ResultingText(KeyCode)
110      SearchOn lsSearchOn
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
   'Remarks: 1. When the user types something into the text portion, move to the
   '            first list entry which begins with the displayed text
   '         2. Not all keys trigger this event. In particular -
   '              <Delete> - triggers KeyDown by not KeyPress
   '              <BackSpace> - triggers KeyPress by not KeyDown
   '         3. This code was originally in the change() event, but confusing inter
   '            actions kept occurring (list index was being set to -1 by WINDOWS)
   '
10       Dim lsSearchOn As String             'current string to search on

20       On Error Resume Next
   
30       If p_bDisplayd = False Then Exit Sub
40       If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then Exit Sub
   
   'A content-modifying key was entered; decide what to search on
50       lsSearchOn = ResultingText(KeyAscii)
60       If SearchOn(lsSearchOn) = False Then KeyAscii = 0
End Sub

Private Function ResultingText(iKeyAscii%) As String
   'Purpose: Works out the text string that results from an original string
   '         comprising the specified elements, following addition of <KeyAscii>
   '         at <iSelStart>
   '
   'Returns: Resulting text string
   
10       Dim sLeft As String             ' string element
20       Dim sSel As String              ' selected string element
30       Dim sRight As String            ' string element
40       Dim sResult As String           ' what well return
   
50       On Error Resume Next
   
60       With txtSearch
70          sLeft = Left$(.Text, .SelStart)         ' SelStart is 0-based
80          sSel = Mid$(.Text, .SelStart + 1, .SelLength)
90          sRight = Mid$(.Text, .SelStart + .SelLength + 1)
100      End With
   
   ' now check if text field is cleared
110      If sLeft = "" And sRight = "" Then
      ' clear result to allow new search criteria
120         MSFlexGrid1.Visible = False
130         p_bDisplayd = False
140      End If
   
150      Select Case iKeyAscii
      Case vbKeyBack             'Backspace Key
160            If Len(sSel) = 0 Then   'Nothing selected
170               sResult = MinusRightChar(sLeft) & sRight  'Del first char on the left
180            Else                    'Selection exists
190               sResult = sLeft & sRight   'Delete selected text only
200            End If
         
210         Case vbKeyDelete           'Delete key
220            If Len(sSel) = 0 Then   'Nothing selected
230               sResult = sLeft & MinusLeftChar(sRight)    'Del first char on the right
240            Else
250               sResult = sLeft & sRight    'Delete selected text only
260            End If
         
270         Case Else         'an ordinary character
280            sResult = sLeft & Chr$(iKeyAscii) & sRight
290      End Select
300      ResultingText = sResult
End Function

Private Function MinusLeftChar(ByVal sGiven As String) As String

   'Purpose: Returns <sGiven> with the leftmost character removed, or "" if
   '         <sGiven> was empty.
   '
   'Returns: The trimmed string
   '
   'Remarks: Just a safe wrapper for Mid$()
10       On Error Resume Next
   
20       If Len(sGiven) = 0 Then
30          MinusLeftChar = ""
40       Else
50          MinusLeftChar = Mid$(sGiven, 2)
60       End If
End Function

Private Function MinusRightChar(ByVal sGiven As String) As String

   'Purpose: Returns <sGiven> with the rightmost character removed, or "" if
   '         <sGiven> was empty.
   '
   'Returns: The trimmed string
   '
   'Remarks: Just a safe wrapper for Left$()
10       On Error Resume Next
   
20       If Len(sGiven) = 0 Then
30          MinusRightChar = ""
40       Else
50          MinusRightChar = Left$(sGiven, Len(sGiven) - 1)
60       End If
End Function

Private Sub MSFlexGrid1_LostFocus()
10       MSFlexGrid1.BackColorSel = &H800000
End Sub

Private Sub MSFlexGrid1_DblClick()
10       With MSFlexGrid1
20          If .MouseRow = 0 Then
30             If .MouseCol <> (cmbSearch.ListIndex) Then
40                cmbSearch.ListIndex = .MouseCol
50                SortList
60             End If
70          Else
80             xrButton1_Click 0
90          End If
100      End With
End Sub

Private Sub MSFlexGrid1_GotFocus()
10       With MSFlexGrid1
20          .HighLight = flexHighlightAlways
30          .BackColorSel = &HB06F00
40       End With
End Sub

Private Function SearchOn(ByVal lsSeek) As Boolean
10       Dim lnCtr As Long
20       Dim lbFound As Boolean
   
30       lbFound = False
40       With MSFlexGrid1
50          For lnCtr = 1 To .Rows
60             If StrComp(Left(.TextMatrix(lnCtr, cmbSearch.ListIndex), Len(lsSeek)), lsSeek, vbTextCompare) >= 0 Then
70                .TopRow = lnCtr
80                .Row = lnCtr
90                .RowSel = lnCtr
100               .ColSel = MSFlexGrid1.Cols - 1
110               lbFound = True
120               Exit For
130            End If
140         Next
150      End With
160      SearchOn = lbFound
End Function

Private Sub SortList()
10       p_oLookup.Sort = p_asColName(cmbSearch.ListIndex)
20       ReLoadList
End Sub

Private Function getSelectedItem() As Variant
10       Dim lvSelected As Variant
20       Dim lsProcName As String
   
30       lsProcName = "getSelectedItem"
40       On Error GoTo errProc
   
50       lvSelected = ""
60       With MSFlexGrid1
70          If .RowSel > 0 Then
80             p_oLookup.MoveFirst
90             p_oLookup.Move .RowSel - 1, adBookmarkFirst
100            For pnCtr = 0 To p_oLookup.Fields.Count - 1
110               Select Case p_oLookup(pnCtr).Type
            Case 2, 3, 11, 17, 72, 4, 5, 6, 131
120                  lvSelected = lvSelected & _
                           IIf(IsNull(p_oLookup(pnCtr)), "", p_oLookup(pnCtr)) & "»"
130               Case 135
140                  lvSelected = lvSelected & _
                              IIf(IsNull(p_oLookup(pnCtr)), "", p_oLookup(pnCtr)) & "»"
150               Case Else
160                  lvSelected = lvSelected & _
                              IIf(IsNull(p_oLookup(pnCtr)), "", p_oLookup(pnCtr)) & "»"
170               End Select
180            Next
190            lvSelected = Left(lvSelected, Len(lvSelected) - 1)

200         End If
210      End With
220      getSelectedItem = lvSelected

endProc:
230      Exit Function
errProc:
240      ShowError lsProcName & "( " & " )"
End Function

Private Sub p_oLookup_MoveComplete(ByVal adReason As EventReasonEnum, ByVal pError As Error, adStatus As EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
10       If Not pbProgress Then Exit Sub
20       DoEvents
30       If Not pRecordset.EOF Then MoveProgress
End Sub

Private Sub txtSearch_GotFocus()
10       pbFocus = True
End Sub

Private Sub txtSearch_LostFocus()
10       pbFocus = False
End Sub

Private Sub xrButton1_Click(Index As Integer)
10       Select Case Index
   Case 0
20          If MSFlexGrid1.RowSel < 1 Or p_bDisplayd = False Then
30             MsgBox "Nothing to Load!", vbInformation, "Warning"
40             p_bSelected = False
50             Exit Sub
60          End If
70          p_bSelected = True
80          Me.Hide
90       Case 1
100         p_bSelected = False
110         Me.Hide
120      Case 2
130         getList
140         p_oLookup.Sort = p_asColName(cmbSearch.ListIndex)
150         LoadList
      
160         txtSearch.SetFocus
170      End Select
End Sub

Private Sub showProgress(ByVal lnMaxLength As Long)
10       pnInterval = 1
20       pnProgress = 1
30       If lnMaxLength > 32767 Then
40          pnInterval = Int(lnMaxLength / 32767)
50          ProgressBar1.Max = 32767
60       Else
70          ProgressBar1.Max = lnMaxLength
80       End If
   
90       pbProgress = True
100      ProgressBar1.Visible = True
End Sub

Private Sub MoveProgress()
10       pnProgress = pnProgress + 1
20       DoEvents
30       ProgressBar1.Value = Int(pnProgress / pnInterval)
40       DoEvents
End Sub

Private Sub hideProgress()
10       pbProgress = False
20       ProgressBar1.Visible = False
End Sub

Private Sub ShowError(ByVal lsProcName As String)
10        With p_oAppDrivr
20           .xLogError Err.Number, Err.Description, "frmLookUp1", lsProcName, Erl
30        End With
40        With Err
50           .Raise .Number, .Source, .Description
60        End With
End Sub

Private Sub getList()
10       Dim lsProcName As String
20       Dim lsSQL As String
   
30       lsProcName = "getList"
40       On Error GoTo errProc
   
50       If p_sSQLQuery <> Empty Then
60          lsSQL = p_sSQLQuery
70       Else
80          lsSQL = p_oLookup.Source
90       End If
   
100      If txtSearch.Text <> Empty Then
110         lsSQL = p_oMod.AddCondition(lsSQL, p_asFldName(cmbSearch.ListIndex) & " LIKE " & p_oMod.strParm(Trim(txtSearch) & "%"))
120      End If
   
130      If p_oLookup.State = adStateOpen Then p_oLookup.Close
140      p_oLookup.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, , adCmdText
   
endProc:
150      Exit Sub
errProc:
160      ShowError lsProcName & "( " & " )"
End Sub

Private Sub ReLoadList()
10       Dim lvValue As Variant
20       Dim lnCol As Long
30       Dim lsProcName As String
   
40       lsProcName = "ReLoadList"
50       On Error GoTo errProc
   
60       With MSFlexGrid1
70          If p_oLookup.RecordCount = 0 Then
80             .Rows = 2
90             GoTo endProc
100         End If
      
110         p_oLookup.MoveFirst
120         .Rows = p_oLookup.RecordCount + 1
      
130         showProgress .Rows + 1
140         pnCtr = 0
150         p_bDisplayd = True
160         Do Until p_oLookup.EOF
170            pnCtr = pnCtr + 1
180            For lnCol = 0 To UBound(p_asColName)
190               lvValue = p_oLookup(p_asColName(lnCol))
200               If IsNull(p_oLookup(p_asColName(lnCol))) Then lvValue = Empty
210               .TextMatrix(pnCtr, lnCol) = Format(lvValue, p_asColPict(lnCol))
220            Next
         
230            p_oLookup.MoveNext
240         Loop
250         hideProgress
260      End With

endProc:
270      Exit Sub
errProc:
280      ShowError lsProcName & "( " & " )"
End Sub

