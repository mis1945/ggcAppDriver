VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Rex S. Adversalo
' XerSys Computing
' Canaoalan, Binmaley, Pangasinan
'
' Copyright 2003-2004 and beyond
' All Rights Reserved
'
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                 €€  Tel No. 522-1085 ; 522-0863      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
'
' ==========================================================================================
'  XerSys [ 06/15/2004 11:50 am ]
'     Start revising this object. i adopt the idea of the form driver of my dear friend
'  Marlon Sayson (with his consent). Included in here are new procedure format.
' ------------------------------------------------------------------------------------------
'  XerSys [ 12/19/2006 09:20 am ]
'     Create a second version of this object
'        -  Refer to GCC SEGs Notes for detailed changes of this object
' ------------------------------------------------------------------------------------------
'  Kalyptus [ 12/20/2006 11:20 am ]
'     Inserts BugTracking
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€

Option Explicit

'-----------------------------------------------------------
' Private variables here
'-----------------------------------------------------------

Private p_oAppDrivr As clsAppDriver
Private p_oMod As New clsMainModules

Private p_oRecMain As ADODB.Recordset
Private p_oMainForm As Object

Private p_axTextFld() As FieldProperty
Private p_asAHFName() As String
Private p_asAHTitle() As String
Private p_asAHFPict() As String
Private p_asAHValue() As String
Private p_abAHFRefx() As Boolean

Private p_sQuerySQL As String
Private p_sBrowsSQL As String
Private p_nCurIndex As Integer
Private p_nFldCount As Integer
Private p_nFldStart As Integer
Private p_bFldChnge As Boolean
Private p_bRecExist As Boolean
Private p_bShowMsg As Boolean
Private p_nEditMode As xeEditMode

Private pnCtr As Long

'-----------------------------------------------------------
'Class Events Here
'-----------------------------------------------------------
' Triggered before the record is deleted. Can be used in validating record
Public Event WillDelete(Cancel As Boolean)

' Triggered before the record is deleted. Can be used to override the delete
' module of this object...
Public Event Delete(Deleted As Boolean)

' Triggered after the record is deleted. Can be used in performing
' other process after deleting record...
Public Event DeleteComplete()

' Triggered before the record is saved. Can be used in validating record
Public Event WillSave(Cancel As Boolean)

' Triggered before the record is save. Can be used to override the save module
Public Event Save(Saved As Boolean)

' Triggered after the record is saved. Can be used in performing other
' process after saving record...
Public Event SaveComplete()

' Triggered after enabling the TextBox for entry. Can be used to enable other controls
Public Event EnableOtherControl()

' Triggered after disabling the TextBox for entry. Can be used to disable other controls
Public Event DisableOtherControl()

' Triggered after emptying the TextBox for entry. Can be used to set the
' initial value of other controls
Public Event InitValue()

' Triggered after loading the data for the textbox. Can be used to set the value
' of the the other data entry controls
Public Event LoadOtherData()

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Set MainForm(oForm As Object)
   Set p_oMainForm = oForm
End Property

Property Get MainForm() As Object
   Set MainForm = p_oMainForm
End Property

Property Let RecQuery(Value As String)
   p_sQuerySQL = Value
End Property

Property Get RecQuery() As String
   RecQuery = p_sQuerySQL
End Property

Property Let BrowseQuery(Value As String)
   p_sBrowsSQL = Value
End Property

Property Get BrowseQuery() As String
   BrowseQuery = p_sBrowsSQL
End Property

Property Let BrowseColumn(ByVal Index As Integer, ByVal Value As String)
   p_asAHFName(Index) = Value
End Property

Property Get BrowseColumn(ByVal Index As Integer) As String
   BrowseColumn = p_asAHFName(Index)
End Property

Property Let BrowseFTitle(ByVal Index As Integer, ByVal Value As String)
   p_asAHTitle(Index) = Value
End Property

Property Get BrowseFTitle(ByVal Index As Integer) As String
   BrowseFTitle = p_asAHTitle(Index)
End Property

Property Let BrowseFFormat(ByVal Index As Integer, ByVal Value As String)
   p_asAHFPict(Index) = Value
End Property

Property Get BrowseFFormat(ByVal Index As Integer) As String
   BrowseFFormat = p_asAHFPict(Index)
End Property

Property Let BrowseFReference(ByVal Index As Integer, ByVal Value As Boolean)
   p_abAHFRefx(Index) = Value
End Property

Property Get BrowseFReference(ByVal Index As Integer) As Boolean
   BrowseFReference = p_abAHFRefx(Index)
End Property

Property Let DisplayConfirmation(Value As Boolean)
   p_bShowMsg = Value
End Property

Property Get EditMode() As xeEditMode
   EditMode = p_nEditMode
End Property

Property Let LookupQuery(ByVal Index As Integer, ByVal Value As String)
   p_axTextFld(Index).Query = Value

   With p_oMainForm
      If Value = "" Then
         .txtField(Index).MaxLength = FieldSize(Index)
      Else
         .txtField(Index).MaxLength = 0
      End If
   End With
End Property

Property Get LookupQuery(ByVal Index As Integer) As String
   LookupQuery = p_axTextFld(Index).Query
End Property

Property Get LookupColumn(ByVal Index As Integer) As String
   LookupColumn = p_axTextFld(Index).QColumn
End Property

Property Let LookupColumn(ByVal Index As Integer, Value As String)
   p_axTextFld(Index).QColumn = Value
End Property

Property Let LookupTitle(ByVal Index As Integer, Value As Variant)
   p_axTextFld(Index).QTitle = Value
End Property

Property Get LookupTitle(ByVal Index As Integer) As Variant
   LookupTitle = p_axTextFld(Index).QTitle
End Property

Property Let LookupFormat(ByVal Index As Integer, ByVal Value As Variant)
   p_axTextFld(Index).QFormat = Value
End Property

Property Get LookupFormat(ByVal Index As Integer) As Variant
   LookupFormat = p_axTextFld(Index).QFormat
End Property

Property Let LookupReference(ByVal Index As Integer, ByVal Value As Variant)
   p_axTextFld(Index).QReference = Value
End Property

Property Get LookupReference(ByVal Index As Integer) As Variant
   LookupReference = p_axTextFld(Index).QReference
End Property

Property Let LookupValue(ByVal Index As Integer, Value As Variant)
   p_asAHValue(Index) = Value
End Property

Property Get LookupValue(ByVal Index As Integer) As Variant
   LookupValue = p_asAHValue(Index)
End Property

Property Let FieldName(ByVal Index As Integer, ByVal Value As String)
   p_axTextFld(Index).Name = Value
End Property

Property Get FieldName(ByVal Index As Integer) As String
   FieldName = p_axTextFld(Index).Name
End Property

Property Let FieldValue(ByVal Index As Integer, ByVal Value As Variant)
   p_axTextFld(Index).Value = Value
End Property

Property Get FieldValue(ByVal Index As Integer) As Variant
   FieldValue = p_axTextFld(Index).Value
End Property

Property Let FieldType(ByVal Index As Integer, ByVal Value As String)
   p_axTextFld(Index).Type = Value
End Property

Property Get FieldType(ByVal Index As Integer) As String
   FieldType = p_axTextFld(Index).Type
End Property

Property Let FieldSize(ByVal Index As Integer, ByVal Value As Integer)
   p_axTextFld(Index).Size = Value
   p_oMainForm.txtField(Index).MaxLength = Value
End Property

Property Get FieldSize(ByVal Index As Integer) As Integer
   FieldSize = p_axTextFld(Index).Size
End Property

Property Let FieldHelp(ByVal Index As Integer, ByVal Value As String)
   p_axTextFld(Index).Help = Value
End Property

Property Get FieldHelp(ByVal Index As Integer) As String
   FieldHelp = p_axTextFld(Index).Help
End Property

Property Let FieldReference(ByVal Index As Integer, ByVal Value As Boolean)
   p_axTextFld(Index).Reference = Value
End Property

Property Get FieldReference(ByVal Index As Integer) As Boolean
   FieldReference = p_axTextFld(Index).Reference
End Property

Property Let FieldEmpty(ByVal Index As Integer, ByVal Value As Boolean)
   p_axTextFld(Index).Empty = Value
End Property

Property Get FieldEmpty(ByVal Index As Integer) As Boolean
   FieldEmpty = p_axTextFld(Index).Empty
End Property

Property Let FieldInit(ByVal Index As Integer, ByVal Value As Boolean)
   p_axTextFld(Index).Initialize = Value
End Property

Property Get FieldInit(ByVal Index As Integer) As Boolean
   FieldInit = p_axTextFld(Index).Initialize
End Property

Property Let FieldFormat(ByVal Index As Integer, ByVal Value As String)
   p_axTextFld(Index).Format = Value
End Property

Property Get FieldFormat(ByVal Index As Integer) As String
   FieldFormat = p_axTextFld(Index).Format
End Property

Property Let ColumnIndex(ByVal Value As Integer)
   p_nCurIndex = Value
   With MainForm
      If .txtField(Value).Enabled And .txtField(Value).Visible Then
         .txtField(Value).SelStart = 0
         .txtField(Value).SelLength = Len(.txtField(Value).Text)
      End If
   End With
End Property

Property Get ColumnIndex() As Integer
   ColumnIndex = p_nCurIndex
End Property

Property Get FieldCount() As Integer
   FieldCount = p_nFldCount
End Property

Property Let FieldStart(nStart As Integer)
   p_nFldStart = nStart
End Property

Property Get FieldStart() As Integer
   FieldStart = p_nFldStart
End Property

Property Get RecStatus() As Integer
   RecStatus = p_nEditMode
End Property

' Show the command button for record process (cmdButton)
'  Index = -1 all record process button will be enabled; record entry button disabled
'        = >= 0 cmdbutton with that index will be enabled
Sub showButton(ByVal Index As Integer)
   Dim loButton As Control
   Dim lsProcName As String
   Dim lbShow As Boolean
   
   lsProcName = "ShowButton"
   On Error GoTo errProc
   
   With MainForm
      If p_nEditMode = xeModeReady Then lbShow = True
      If Index < 0 Then
         For Each loButton In .cmdButton
            Select Case UCase(loButton.Caption)
            Case "&NEW", "&UPDATE", "&DELETE", "&CLOSE"
               loButton.Visible = lbShow
            Case "&SAVE", "&CANCEL", "SEARC&H"
               loButton.Visible = Not lbShow
            End Select
         Next
      Else
         .cmdButton(Index).Visible = True
      End If
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Index & " )"
End Sub

' hide the command button for record process (cmdButton)
'  Index = see showbutton
Sub hideButton(ByVal Index As Integer)
   Dim loButton As Control
   Dim lsProcName As String
   Dim lbShow As Boolean
   
   lsProcName = "HideButton"
   On Error GoTo errProc
   
   With MainForm
      If p_nEditMode = xeModeReady Then lbShow = True
      If Index < 0 Then
         For Each loButton In .cmdButton
            Select Case UCase(loButton.Caption)
            Case "&NEW", "&UPDATE", "&DELETE", "&CLOSE"
               loButton.Visible = lbShow
            Case "&SAVE", "&CANCEL", "SEARC&H"
               loButton.Visible = Not lbShow
            End Select
         Next
      Else
         .cmdButton(Index).Visible = False
      End If
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Index & " )"
End Sub

' Enables the textbox entry withing the MainForm.
'  Index = -1 Enables all textbox entry (txtfield)
'        = any no. Enables the textbox entry of that index
Sub EnableTextbox(ByVal Index As Integer)
   Dim loText As TextBox
   Dim lsProcName As String
   
   lsProcName = "EnableTextBox"
   On Error GoTo errProc
   
   With MainForm
      If Index < 0 Then
         For Each loText In .txtField
           loText.Enabled = True
         Next
         RaiseEvent EnableOtherControl
      Else
         .txtField(Index).Enabled = True
      End If
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Index & " )"
End Sub

' Disables the textboxes within MainForm
'  Index = see EnableTextbox
Sub DisableTextbox(ByVal Index As Integer)
   Dim loText As TextBox
   Dim lsProcName As String
   
   lsProcName = "DisableTextBox"
   On Error GoTo errProc
   
   With MainForm
      If Index < 0 Then
         For Each loText In .txtField
           loText.Enabled = False
         Next
         RaiseEvent DisableOtherControl
      Else
         .txtField(Index).Enabled = False
      End If
   End With
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Index & " )"
End Sub

' Triggers Validate Event
' Performs formatting and validation of the record entry
Function ValidateField(ByVal Index As Integer) As Boolean
   Dim lors As ADODB.Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lasRef() As String
   Dim vValue As Variant
   
   lsProcName = "ValidateField"
   ValidateField = False
   On Error GoTo errProc
   
   ' Perform textfields that needs validation here!
   With MainForm
      .txtField(Index).Text = Trim(.txtField(Index).Text)

      ' If data encoded is valid/changed then perform the following
      chkFieldStat Index
      If p_bFldChnge Then
         If Index < FieldStart Then
            .txtField(Index).Text = Format(.txtField(Index), FieldFormat(Index))
            FieldValue(Index) = .txtField(Index).Text
         Else
            ' check if field has a lookup table
            lsSQL = LookupQuery(Index)
            If lsSQL = Empty Then
               ValidateField = SetValue(Index, .txtField(Index).Text)
               GoTo endProc
            End If
            Debug.Print vValue
            vValue = .txtField(Index).Text

            If Trim(vValue) = "" Then
               ValidateField = SetValue(Index, .txtField(Index).Text)
               GoTo endProc
            End If

            ' Search for the existence of the TextBox Entry
            lasRef = Split(LookupReference(Index), "»")
            lsSQL = lasRef(1) & " = " & p_oMod.strParm(Trim(vValue))
            lsSQL = p_oMod.AddCondition(LookupQuery(Index), lsSQL)

            Debug.Print lsSQL
            Set lors = New Recordset
            lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
            
            If lors.EOF Then
               Debug.Print SetValue(Index, Empty)
               ValidateField = SetValue(Index, Empty)
               GoTo endProc
            End If
            
            'Existing ba?
            If lors.RecordCount = 1 Then
               FieldValue(Index) = lors(0).Value
               .txtField(Index).Text = lors(1).Value
               ValidateField = True
               GoTo endProc
            End If
            vValue = p_oMod.KwikBrowse(p_oAppDrivr, lors, LookupColumn(Index), _
                        LookupTitle(Index), LookupFormat(Index))
                                
            If vValue <> Empty Then
               p_asAHValue = Split(vValue, "»")
               If p_asAHValue(0) = Empty Then
                  ValidateField = SetValue(Index, Empty)
                  GoTo endProc
               End If
            End If
            
            FieldValue(Index) = Trim(p_asAHValue(0))
            .txtField(Index).Text = p_asAHValue(1)
         End If
         ValidateField = True
      Else
         If Not (isFieldEmpty(Index, .txtField(Index)) And FieldEmpty(Index) = False) Then ValidateField = True
      End If
   End With
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & Index & " )"
End Function

' Empty the contents of  the textboxes within MainForm
' Triggers InitValue
Sub InitTextBoxes()
   Dim lsProcName As String
   
   lsProcName = "InitTextBoxes"
   On Error GoTo errProc
   
   ' Initialize first the fields type and format them accordingly
   For pnCtr = 0 To FieldCount - 1
      If FieldInit(pnCtr) Then setDefaultValue pnCtr
   Next
   Field2Control
   p_bRecExist = False
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

' Set the value of the field and the textbox. This also validate the textbox
' and return if the textbox has a valid value
Function SetValue(ByVal Index As Integer, _
                  Value As Variant) As Boolean
   Dim lsProcName As String
   
   lsProcName = "SetValue"
   On Error GoTo errProc
   
   With MainForm
      If IsNull(Value) Then Value = Empty
         
      Select Case FieldType(Index)
      Case "n", "l"
         If IsNumeric(Value) = False Then
            Value = IIf(FieldType(Index) = "l", 0#, 0)
         Else
            SetValue = True
         End If
         
         ' after validating data type, check if allowed empty
         If Value = 0 And FieldEmpty(Index) = False Then
            SetValue = False
            GoTo endProc
         End If
      Case "d", "t", "p"
         If IsDate(Value) = False Then
            Select Case FieldType(Index)
            Case "d"
               Value = Now
            Case "t"
               Value = Time
            Case Else
               Value = Date
            End Select
         Else
            SetValue = True
         End If
      Case "s"
         ' after validating data type, check if allowed empty
         If Value = "" And FieldEmpty(Index) = False Then GoTo endProc
         SetValue = True
      End Select
      .txtField(Index) = Format(Value, FieldFormat(Index))
      If IsNull(Value) Then Value = ""
      FieldValue(Index) = Value
   End With
   
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & Index _
                       & ", " & Value & " )"
End Function

' Initialize Textfield sizes using it's field size.
Sub InitFieldSize()
   Dim lsProcName As String
   Dim loText As TextBox
   
   lsProcName = "InitFieldSize"
   On Error GoTo errProc
   
   With MainForm
      For Each loText In .txtField
         If FieldType(loText.Index) = "s" And LookupQuery(loText.Index) = Empty Then
            loText.MaxLength = FieldSize(loText.Index)
         End If
      Next
   End With

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

' Initializes the field property base on the query string
Function InitRecForm(Optional Query As Variant, _
                     Optional BQuery As Variant) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "InitRecForm"
   InitRecForm = False
   On Error GoTo errProc
   
   If IsMissing(Query) = False Then p_sQuerySQL = Query
   If IsMissing(BQuery) = False Then p_sBrowsSQL = BQuery
   ' if either query string or application driver is not present, abort object
   If p_sQuerySQL = Empty Then GoTo endProc
   If isAppDriverOK = False Then GoTo endProc
   
   If p_sBrowsSQL = Empty Then p_sBrowsSQL = p_sQuerySQL
   
   ' initialize first the browse field
   lsSQL = p_oMod.getEmptyRec(p_sBrowsSQL)
   Set p_oRecMain = New Recordset
   Debug.Print lsSQL
   p_oRecMain.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   ReDim p_asAHFName(p_oRecMain.Fields.Count - 1) As String
   ReDim p_asAHTitle(p_oRecMain.Fields.Count - 1) As String
   ReDim p_asAHFPict(p_oRecMain.Fields.Count - 1) As String
   ReDim p_abAHFRefx(p_oRecMain.Fields.Count - 1) As Boolean
   ReDim p_asAHValue(p_oRecMain.Fields.Count - 1) As String
   
   For pnCtr = 0 To UBound(p_asAHTitle)
      p_asAHFName(pnCtr) = p_oRecMain(pnCtr).Name
      p_asAHTitle(pnCtr) = p_oRecMain(pnCtr).Name
      p_asAHFPict(pnCtr) = "@"
      p_abAHFRefx(pnCtr) = False
   Next
   p_abAHFRefx(0) = True
   p_oRecMain.Close
   
   ' it must retrieve an empty record coz the default is new record
   lsSQL = p_oMod.getEmptyRec(p_sQuerySQL)
   p_oRecMain.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockPessimistic, adCmdText
   Set p_oRecMain.ActiveConnection = Nothing

   'Initialize variables to be use in manipulating the fields
   p_nFldCount = p_oRecMain.Fields.Count
   
   ReDim p_axTextFld(p_nFldCount) As FieldProperty
   
   For pnCtr = 0 To p_nFldCount - 1
      p_axTextFld(pnCtr).Name = p_oRecMain(pnCtr).Name
      p_axTextFld(pnCtr).Value = Null
      p_axTextFld(pnCtr).Size = p_oRecMain(pnCtr).DefinedSize
      p_axTextFld(pnCtr).Help = Empty
      p_axTextFld(pnCtr).Format = "@"
      p_axTextFld(pnCtr).Reference = False
      p_axTextFld(pnCtr).Initialize = True
      p_axTextFld(pnCtr).Empty = True
      p_axTextFld(pnCtr).Query = Empty
      p_axTextFld(pnCtr).QTitle = Empty
      p_axTextFld(pnCtr).QFormat = Empty
      p_axTextFld(pnCtr).QType = Empty
      p_axTextFld(pnCtr).FieldType = p_oRecMain(pnCtr).Type
      
      Select Case p_oRecMain(pnCtr).Type
      Case 129, 130, 202, 200    ' string
         p_axTextFld(pnCtr).Type = "s"
      Case 2, 3, 11, 17, 72      ' numeric without decimal point
         p_axTextFld(pnCtr).Type = "n"
      Case 4, 5, 6, 131          ' numeric with decimal point
         p_axTextFld(pnCtr).Type = "l"
      Case 133                   ' date
         p_axTextFld(pnCtr).Type = "p"
      Case 135                   ' datetime
         p_axTextFld(pnCtr).Type = "d"
      End Select
   Next
   ' always assume that the first rec is the reference field
   p_axTextFld(0).Reference = True
   
   p_nEditMode = xeModeReady
   InitFieldSize
   InitTextBoxes
   hideButton -1
   DisableTextbox -1
   InitRecForm = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & Query _
                       & ", " & BQuery & " )"
End Function

Sub LoadRecord()
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "LoadRecord"
   On Error GoTo errProc
   
   ' Load the data from the selected record
   GetRecord
   With MainForm
      If Not p_oRecMain.EOF Then
         FieldValue(0) = p_oRecMain(0).Value
         .txtField(0).Text = FieldValue(0)
         For pnCtr = 1 To FieldCount - 1
            ' Load the field value to the field recipeint
            FieldValue(pnCtr) = p_oRecMain(pnCtr)
         Next
         Field2Control
         RaiseEvent LoadOtherData
      Else
         InitTextBoxes
      End If
   End With

   ' after loading record disable edit button
   p_nEditMode = xeModeReady
   p_bRecExist = True
   
   DisableTextbox -1
   If p_bShowMsg Then
      hideButton -1
   End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Sub BrowseRecord()
   Dim lsProcName As String
   Dim lsAHValue As String
   Dim lsColumn As String
   Dim lsTitle As String
   Dim lsPict As String
   
   lsProcName = "BrowseRecord"
   On Error GoTo errProc
   
   showRemarks "Browsing Records..."
   
   lsColumn = ""
   lsTitle = ""
   lsPict = ""
   For pnCtr = 0 To UBound(p_asAHTitle)
      If p_asAHTitle(pnCtr) <> "" Then
         lsColumn = lsColumn & "»" & p_asAHFName(pnCtr)
         lsTitle = lsTitle & "»" & p_asAHTitle(pnCtr)
         lsPict = lsPict & "»" & p_asAHFPict(pnCtr)
      End If
   Next
   lsColumn = Mid(lsColumn, 2, Len(lsColumn) - 1)
   lsTitle = Mid(lsTitle, 2, Len(lsTitle) - 1)
   lsPict = Mid(lsPict, 2, Len(lsPict) - 1)
   
   lsAHValue = p_oMod.KwikSearch(p_oAppDrivr, p_sBrowsSQL, lsColumn, lsTitle, lsPict)
   
   If lsAHValue = Empty Then
      showRemarks "No selection was made..."
      GoTo endProc
   End If
   p_asAHValue = Split(lsAHValue, "»")
   LoadRecord
   showRemarks "Record was Loaded.", "Ready"

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

' Adds new record
Sub RecordNew()
   'Initialize all value of the controls here
   InitTextBoxes
   RaiseEvent InitValue
   
   p_nEditMode = xeModeAddNew
   
   EnableTextbox -1
   If p_bShowMsg Then
      ' enable controls for editing
      showButton -1
      
      p_nCurIndex = FieldStart
      With p_oMainForm
         .txtField(p_nCurIndex).SetFocus
      End With
      showRemarks , "New"
   End If
   
End Sub

' Use this to save the records
Function RecordSave() As Boolean
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lbCancel As Boolean
   Dim lnRow As Long
   
   lsProcName = "RecordSave"
   On Error GoTo errProc

   If p_bShowMsg Then
      If MsgBox("Are you sure you want to save this record?", _
            vbYesNo + vbInformation, "Confirm") <> vbYes Then GoTo endProc
      
      showRemarks "Saving Record..."
      
      If HasRights2Save() = False Then
         showRemarks "Record was not saved."
         GoTo endProc
      End If
   End If
  
   If isRefEmpty() = True Then
      MsgBox "Reference Fields can not Contain Empty Value!!!" & vbCrLf & _
                     " Verify your Entry then Try Again!!!", vbCritical, "Warning"
      showRemarks "Record was not saved."
      GoTo endProc
   End If
         
   With p_oAppDrivr
      .BeginTrans
      FieldValue(p_nFldCount - 2) = p_oAppDrivr.UserID
   
      lbCancel = False
      ' record validation must be in this event
      RaiseEvent WillSave(lbCancel)
   
      If lbCancel = True Then
         showRemarks "Saving Record was Aborted."
         GoTo endWithRoll
      End If
 
      If isModify = False Then GoTo RecSaved
 
      lbCancel = False
      RaiseEvent Save(lbCancel)
      
      ' if record is saved by the client then endproc
      If lbCancel Then GoTo RecSaved
  
      ' Save data base on the kind of task
      If p_nEditMode = xeModeAddNew Then p_oRecMain.AddNew
             
      For pnCtr = 0 To p_nFldCount - 3
         If IsNull(p_oRecMain(FieldName(pnCtr))) Or _
               p_oRecMain(FieldName(pnCtr)).OriginalValue <> _
               FieldValue(pnCtr) Then
            If IsNull(FieldValue(pnCtr)) = False Then
               lsSQL = lsSQL & ", " & _
                        FieldName(pnCtr) & " = " & _
                        p_oMod.FieldParam(p_axTextFld(pnCtr).FieldType, _
                        FieldValue(pnCtr))
            End If
         End If
      Next
      If lsSQL <> Empty Then
         If p_nEditMode = xeModeAddNew Then
            lsSQL = "INSERT INTO " & getTableName(p_sQuerySQL) & " SET" & _
                        Mid(lsSQL, 2) & _
                        ", sModified = " & _
                           p_oMod.strParm(p_oAppDrivr.UserID) & _
                        ", dModified = " & _
                           p_oMod.dateParm(p_oAppDrivr.ServerDate)
         Else
            lsSQL = "UPDATE " & getTableName(p_sQuerySQL) & " SET" & _
                        Mid(lsSQL, 2) & _
                        ", sModified = " & _
                           p_oMod.strParm(p_oAppDrivr.UserID) & _
                        ", dModified = " & _
                           p_oMod.dateParm(p_oAppDrivr.ServerDate) & _
                     " WHERE "
            
            For pnCtr = 0 To p_nFldCount - 2
               If FieldReference(pnCtr) Then
                  lsSQL = lsSQL & _
                           FieldName(pnCtr) & " = " & _
                           p_oMod.strParm(p_oRecMain(FieldName(pnCtr))) & _
                           " AND "
               End If
            Next
            
            If Right(Trim(lsSQL), 3) <> "AND" Then GoTo endWithRoll
            lsSQL = Left(lsSQL, Len(Trim(lsSQL)) - 3)
         End If
         
         lnRow = .Execute(lsSQL, getTableName(p_sQuerySQL))
         If lnRow <= 0 Then
            showRemarks "Unable to Save Record."
            GoTo endWithRoll
         End If
      End If
     
RecSaved:
      .CommitTrans
      ' after saving the main table, trigger the savecomplete event
      RaiseEvent SaveComplete
      
      ' show some message that the record was save successfully
      showRemarks "Record was Saved Successfully.", "Ready"
      
      If p_nEditMode = xeModeAddNew Then
         RecordNew
      Else
         p_asAHValue(0) = p_oRecMain(0)
         LoadRecord
         
         p_nEditMode = xeModeReady
         hideButton -1
         DisableTextbox -1
      End If
   End With
   
   RecordSave = True
   
endProc:
   Exit Function
   
endWithRoll:
   p_oAppDrivr.RollBackTrans
   Exit Function
errProc:
   p_oAppDrivr.RollBackTrans
   ShowError lsProcName & "( " & " )"
End Function

'Use this to search lookup fields...
Sub RecordSearch(Optional Value As Variant = "")
   Dim lors As ADODB.Recordset
   Dim lsProcName As String
   Dim lsAHValue As String
   Dim lsLookupSQL As String
   Dim lasAHValue() As String
   Dim lasRef() As String
   
   lsProcName = "RecordSearch"
   On Error GoTo errProc
   
   With p_oMainForm
      If LookupQuery(p_nCurIndex) <> Empty And p_nCurIndex > 0 Then
         lsLookupSQL = LookupQuery(p_nCurIndex)
         If Value <> "" Then
            ' no changes was made then abort search
            If .txtField(p_nCurIndex).Tag = .txtField(p_nCurIndex).Text Then
               GoTo endProc
            End If
            lasRef = Split(LookupReference(p_nCurIndex), "»")
            lsLookupSQL = lasRef(1) & " LIKE " & p_oMod.strParm(Trim(Value) & "%")

            lsLookupSQL = p_oMod.AddCondition(LookupQuery(p_nCurIndex), lsLookupSQL)
         End If
         Set lors = New Recordset
         
         lors.Open lsLookupSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
         
         If lors.EOF Then GoTo endProc
         
         If lors.RecordCount = 1 Then
            FieldValue(p_nCurIndex) = lors(0)
            .txtField(p_nCurIndex).Text = lors(1)
            .txtField(p_nCurIndex).Tag = lors(1)
            GoTo endProc
         End If
         
         lsAHValue = p_oMod.KwikBrowse(p_oAppDrivr, lors, _
                           LookupColumn(p_nCurIndex), LookupTitle(p_nCurIndex), _
                           LookupFormat(p_nCurIndex))

         If Not (lsAHValue = Empty) Then
            lasAHValue = Split(lsAHValue, "»")
            FieldValue(p_nCurIndex) = Trim(lasAHValue(0))
            .txtField(p_nCurIndex).Text = lasAHValue(1)
            .txtField(p_nCurIndex).Tag = lasAHValue(1)
         End If
         .txtField(p_nCurIndex).SetFocus
      End If
   End With

endProc:
   Set lors = Nothing
   Exit Sub
errProc:
   ShowError lsProcName & "( " & Value & " )"
End Sub

' this enable the record for updating
Sub RecordUpdate()
   If p_bRecExist = False Then
      MsgBox "No Record is Loaded to Modify!!!" & vbCrLf & _
            "Verify your Entry then Try Again!!!", vbInformation, "Warning"
      Exit Sub
   End If

   p_nEditMode = xeModeUpdate
   
   EnableTextbox -1
   If p_bShowMsg Then
      ' enable controls for editing
      showButton -1
      
      p_nCurIndex = FieldStart
      With p_oMainForm
         .txtField(p_nCurIndex).SetFocus
      End With
      showRemarks , "Update"
   End If
End Sub

' actual delete module
Function RecordDelete() As Boolean
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lbCancel As Boolean
   Dim lnCtr As Integer
   Dim lsTable As String

   lsProcName = "RecordDelete"
   On Error GoTo errProc
   
   If p_bRecExist = False Then
      MsgBox "No Record is Loaded to Modify!!!" & vbCrLf & _
           "Verify your Entry then Try Again!!!", vbInformation, "Warning"
      GoTo endProc
   End If

   If p_bShowMsg Then
      If MsgBox("Delete this record?", vbOKCancel + vbCritical, "Confirm") <> vbOK Then
         showRemarks "Deletion of Record was Aborted...", "Ready"
         GoTo endProc
      End If
      
      ' show mode
      showRemarks "Deleting Record...", "Delete"
      
      If HasRights2Delete = False Then
         showRemarks "Record was not Deleted...", "Ready"
         GoTo endProc
      End If
   End If
   
   With p_oAppDrivr
      .BeginTrans
      
      lbCancel = False
      RaiseEvent WillDelete(lbCancel)
      If lbCancel Then
         showRemarks "Deletion of Record was Aborted.", "Ready"
         GoTo endWithRoll
      End If
      
      lbCancel = False
      RaiseEvent Delete(lbCancel)
      
      If lbCancel Then GoTo RecDeleted
      
      lnCtr = InStr(1, p_sQuerySQL, "FROM ", vbTextCompare)
      If lnCtr = 0 Then
         lsSQL = p_sQuerySQL
      Else
         lsSQL = Trim(Mid(p_sQuerySQL, lnCtr + 5))
         lnCtr = InStr(1, lsSQL, " ", vbTextCompare)
         If lnCtr > 0 Then
            lsSQL = Trim(Mid(lsSQL, 1, lnCtr))
         End If
      End If
         
      lsTable = lsSQL
      lsSQL = "DELETE FROM " & lsSQL
      For lnCtr = 0 To UBound(p_abAHFRefx)
         If p_abAHFRefx(lnCtr) Then
            lsSQL = p_oMod.AddCondition(lsSQL, p_asAHFName(lnCtr) & " = " & _
                     p_oMod.strParm(p_oRecMain(p_asAHFName(lnCtr))))
         End If
      Next
      
'      If p_oAppDrivr.RegisDelete(lsSQL) = False Then GoTo endWithRoll
'      p_oRecMain.Delete
      
      If p_oAppDrivr.Execute(lsSQL, lsTable) = 0 Then GoTo endWithRoll
      
RecDeleted:
      .CommitTrans
      RaiseEvent DeleteComplete
   
      ' after deleting record, remove record from display
      InitTextBoxes
      p_nEditMode = xeModeReady
      
      ' show some message that the record was deleted successfully
      showRemarks "Record was Deleted Successfully.", "Ready"
   End With
      
   RecordDelete = True
   
endProc:
   Exit Function
endWithRoll:
   p_oAppDrivr.RollBackTrans
   Exit Function
errProc:
   p_oAppDrivr.RollBackTrans
   ShowError lsProcName & "( " & " )"
End Function

Function RecordCancelUpdate() As Boolean
   If p_nEditMode = xeModeAddNew Then
      If p_asAHValue(0) = Empty Then
         InitTextBoxes
      Else
         LoadRecord
      End If
   Else
      LoadRecord
   End If

   p_nEditMode = xeModeReady
   hideButton -1
   DisableTextbox -1
   
   showRemarks "Modification was Aborted.", "Ready"
End Function

Sub ViewUserModify()
   Dim lors As ADODB.Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "ViewUserModify"
   On Error GoTo errProc
   
   lsSQL = "SELECT sUserName" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx = " & p_oMod.strParm(p_oRecMain("sModified"))

   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If lors.EOF Then
      MsgBox "User does not Exist in the User Table!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   MsgBox "    ««« LAST MODIFICATION »»»" & vbCrLf & vbCrLf & _
            "User Name      : " & p_oMod.Decrypt(lors("sUserName"), p_oAppDrivr.Machinex) & vbCrLf & _
            "Date Modified : " & Format(p_oRecMain("dModified"), "MMMM DD, YYYY HH:NN:SS") _
            , vbInformation, "Notice"
         
endProc:
   Set lors = Nothing
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

' retrieve the records based by the reference passed
Private Sub GetRecord()
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "getRecord"
   On Error GoTo errProc
   
   ' before wasting processor time check if there's a valid value
   If p_asAHValue(0) = "" Then GoTo endProc
   
   lsSQL = p_sQuerySQL
   For pnCtr = 0 To UBound(p_asAHFName)
      If p_abAHFRefx(pnCtr) Then
         lsSQL = p_oMod.AddCondition(lsSQL, p_asAHFName(pnCtr) & " = " & p_oMod.strParm(p_asAHValue(pnCtr)))
      End If
   Next

   If p_oRecMain.State = adStateOpen Then
      If Not p_oRecMain.EOF Then
         If p_oRecMain.EditMode <> adEditNone Then p_oRecMain.CancelUpdate
      End If
      p_oRecMain.Close
   End If
   p_oRecMain.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Private Sub Class_Initialize()
   Set p_oRecMain = New Recordset
   
   p_bShowMsg = True
   p_nEditMode = xeModeAddNew
End Sub

Private Sub Class_Terminate()
   Set p_oRecMain = Nothing
   Set p_oMod = Nothing
End Sub

Private Sub chkFieldStat(ByVal Index As Integer)
   If IsNull(FieldValue(Index)) Then
      p_bFldChnge = MainForm.txtField(Index) <> Empty
   Else
      If LookupQuery(Index) <> Empty Then
         p_bFldChnge = Trim(MainForm.txtField(Index)) <> p_asAHValue(1)
      Else
         Select Case p_axTextFld(Index).Type
         Case "s"
            p_bFldChnge = Trim(FieldValue(Index)) <> Trim(MainForm.txtField(Index))
         Case "n", "l"
            If IsNumeric(MainForm.txtField(Index)) Then
               If p_axTextFld(Index).Type = "n" Then
                  p_bFldChnge = FieldValue(Index) <> CLng(MainForm.txtField(Index))
               Else
                  p_bFldChnge = FieldValue(Index) <> CDbl(MainForm.txtField(Index))
               End If
            Else
               MainForm.txtField(Index) = Format(FieldValue(Index), p_axTextFld(Index).Format)
               p_bFldChnge = False
            End If
         Case "p", "d"
            If IsDate(MainForm.txtField(Index)) Then
               If p_axTextFld(Index).Type = "p" Then
                  p_bFldChnge = Format(FieldValue(Index), "yyyy/mm/dd") <> Format(MainForm.txtField(Index), "yyyy/mm/dd")
               Else
                  p_bFldChnge = Format(FieldValue(Index), "yyyy/mm/dd hh:nn:ss") <> Format(MainForm.txtField(Index), "yyyy/mm/dd hh:nn:ss")
               End If
            Else
               MainForm.txtField(Index) = Format(FieldValue(Index), p_axTextFld(Index).Format)
               p_bFldChnge = False
            End If
         End Select
      End If
   End If
   Debug.Print p_bFldChnge
End Sub

Private Sub setDefaultValue(ByVal lnIndex As Integer)
   Select Case FieldType(lnIndex)
   Case "s"    ' string
      FieldValue(lnIndex) = ""
   Case "n"    ' numeric without decimal value
      FieldValue(lnIndex) = 0
   Case "l"    ' numeric with decimal value
      FieldValue(lnIndex) = 0#
   Case "b"    ' binary
   Case "d"    ' datetime
      FieldValue(lnIndex) = Now
   Case "p"    ' date (petsa)
      FieldValue(lnIndex) = Date
   Case "t"
      FieldValue(lnIndex) = Time
   End Select
End Sub

Private Function isAppDriverOK() As Boolean
   isAppDriverOK = True
   If p_oAppDrivr Is Nothing Then
      MsgBox "Application Driver is not yet set!!!" & vbCrLf & _
               "Can not further proceed!!!", vbCritical, "Warning"
      isAppDriverOK = False
   End If
End Function

Private Sub Field2Control()
   Dim lors As ADODB.Recordset
   Dim loText As TextBox
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lasRef() As String
   Dim lnIndex As Integer
   Dim lvValue As Variant
   
   lsProcName = "Field2Control"
   On Error GoTo errProc
   
   Set lors = New Recordset
   
   ' assign the field value to the textbox
   With MainForm
      For Each loText In .txtField
         lnIndex = loText.Index
         lvValue = IIf(IsNull(FieldValue(lnIndex)), Empty, FieldValue(lnIndex))
         
         lsSQL = LookupQuery(lnIndex)
         If lsSQL = Empty Then
            .txtField(lnIndex) = Format(lvValue, FieldFormat(lnIndex))
         Else
            If lvValue = Empty Then
               .txtField(lnIndex) = Empty
            Else
            
               lasRef = Split(LookupReference(lnIndex), "»")
               If lors.State = adStateOpen Then lors.Close
               lsSQL = p_oMod.AddCondition(LookupQuery(lnIndex), lasRef(0) & " = " & p_oMod.strParm(lvValue))
               
               lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
                  
               ' always assumed that the 2nd field contains the definition of the code.
               If Not lors.EOF Then
                  .txtField(lnIndex) = lors(1)
               Else
                  .txtField(lnIndex).Text = String(15, "?")
               End If
               lors.Close
            End If
         End If
      Next
   End With

endProc:
   Set lors = Nothing
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Private Function isRefEmpty() As Boolean
   Dim lsProcName As String
   
   lsProcName = "isRefEmpty"
   isRefEmpty = False
   On Error GoTo errProc
   
   For pnCtr = 0 To FieldCount
      If FieldReference(pnCtr) = True Then
         If isFieldEmpty(pnCtr, FieldValue(pnCtr)) Then
            isRefEmpty = True
            GoTo endProc
         End If
      End If
   Next
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function isFieldEmpty(ByVal Index As Integer, _
                              Value As Variant) As Boolean
   Dim lsProcName As String
   
   lsProcName = "isFieldEmpty"
   isFieldEmpty = True
   On Error GoTo errProc
   
   Select Case FieldType(Index)
   Case "n", "l"
      If Value = 0 Then GoTo endProc
   Case "d", "t", "p"
      If Value = "01/01/1890" Then GoTo endProc
   Case "s"
      If Value = "" Then GoTo endProc
   End Select
   isFieldEmpty = False
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & Index _
                       & ", " & Value & " )"
End Function

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
       .xLogError Err.Number, Err.Description, "clsFormDriver", lsProcName, Erl
    End With
    With Err
'       .Raise .Number, .Source, .Description
    End With
End Sub

Private Function HasRights2Save() As Boolean
   Dim lsProcName As String
   Dim lnUserRght As Integer, lsUserName As String, lsUserIDxx As String
   Dim lasRights() As String, lnRights As Integer
   Dim lnSeek As Integer, lbSeek As Boolean
   
   lsProcName = "HasRights2Save"
   On Error GoTo errProc
   
   If p_oAppDrivr.UserLevel = xeSysMaster Then
      HasRights2Save = True
      GoTo endProc
   End If
   
   With p_oAppDrivr
      ' Structure of the Menu Rights is
      '  0  Menu ID
      '  1  Add Rights
      '  2  Update Rights
      '  3  Cancel Rights
      '  4  Delete Rights
      '  5  Approve Rights
      '  6  Override Add Rights
      '  7  Override Update Rights
      '  8  Override Cancel Rights
      '  9  Override Delete Rights
      '  10 Override Approve Rights
      If .MDIMain.Controls(.MenuName).Tag = "" Then
         ReDim lasRights(5) As String
   
         lasRights(1) = 231
         lasRights(2) = 230
         lasRights(3) = 228
         lasRights(4) = 192
         lasRights(5) = 230
      Else
         lasRights = Split(.MDIMain.Controls(.MenuName).Tag, "»")
         If lasRights(6) <> Empty Then
            lasRights(1) = lasRights(6)
            lasRights(2) = lasRights(7)
            lasRights(3) = lasRights(8)
            lasRights(4) = lasRights(9)
            lasRights(5) = lasRights(10)
         End If
      End If
      
      lbSeek = False
      If p_nEditMode = xeModeAddNew Then
         ' Add Rights
         lnRights = CInt(lasRights(1))
         If (lnRights And .UserLevel) = 0 Then
            lnSeek = MsgBox("User Doesn't Have the Right to Add New Record!!!" & _
                              vbCrLf & "Seek Authorized User's Approval?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
            If lnSeek = vbNo Then GoTo endProc
            lbSeek = True
         End If
      ElseIf p_nEditMode = xeModeUpdate Then
         ' update Rights
         lnRights = CInt(lasRights(2))
         
         If (lnRights And .UserLevel) = 0 Then
            Select Case .UserLevel
            Case xeEncoder
               lbSeek = True
            Case xeSupervisor
               If DateDiff("d", p_oRecMain("dModified"), p_oAppDrivr.ServerDate) > 2 Then lbSeek = True
            Case xeManager
               If DateDiff("d", p_oRecMain("dModified"), p_oAppDrivr.ServerDate) > 7 Then lbSeek = True
            Case xeAudit
               lbSeek = True
            Case xeSysAdmin
               lbSeek = True
            End Select
         End If
      End If
      
      If lbSeek = False Then
         HasRights2Save = True
         GoTo endProc
      End If
      
      lnSeek = MsgBox("User Doesn't Have the Right to Update Record!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
      If lnSeek = vbNo Then GoTo endProc
      
      If p_oMod.GetApproval(p_oAppDrivr, lnUserRght, lsUserIDxx, lsUserName, lasRights(0)) = False Then GoTo endProc
      
      If p_nEditMode = xeModeAddNew Then
         If (lnRights And lnUserRght) = 0 Then
            MsgBox "Approving User Doesn't Have the Right to Add New Record!!!" & _
                  "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
            GoTo endProc
         End If
      ElseIf p_nEditMode = xeModeUpdate Then
         If (lnRights And lnUserRght) = 0 Then
            Select Case lnUserRght
            Case xeSupervisor
               If DateDiff("d", p_oRecMain("dModified"), p_oAppDrivr.ServerDate) <= 2 Then lbSeek = False
            Case xeManager
               If DateDiff("d", p_oRecMain("dModified"), p_oAppDrivr.ServerDate) <= 7 Then lbSeek = False
            Case xeSysAdmin
               If .UserLevel <> xeSysAdmin Then lbSeek = False
            End Select
            
            If lbSeek Then
               MsgBox "Approving User Doesn't Have the Right to Update Record!!!" & _
                     "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
      End If
   End With
   HasRights2Save = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function HasRights2Delete() As Boolean
   Dim lsProcName As String
   Dim lnUserRght As Integer, lsUserName As String, lsUserIDxx As String
   Dim lasRights() As String, lnRights As Integer
   Dim lnSeek As Integer
   
   lsProcName = "HasRights2Delete"
   On Error GoTo errProc
   HasRights2Delete = False
   
   If p_oAppDrivr.UserLevel = xeSysMaster Then
      HasRights2Delete = True
      GoTo endProc
   End If
   
   With p_oAppDrivr
      If .MDIMain.Controls(.MenuName).Tag = "" Then
         ReDim lasRights(5) As String
   
         lasRights(1) = 231
         lasRights(2) = 230
         lasRights(3) = 228
         lasRights(4) = 192
         lasRights(5) = 230
      Else
         lasRights = Split(.MDIMain.Controls(.MenuName).Tag, "»")
         If lasRights(6) <> Empty Then
            lasRights(1) = lasRights(6)
            lasRights(2) = lasRights(7)
            lasRights(3) = lasRights(8)
            lasRights(4) = lasRights(9)
            lasRights(5) = lasRights(10)
         End If
      End If
      
      ' Delete Rights
      lnRights = CInt(lasRights(4))
      If (lnRights And .UserLevel) > 0 Then
         HasRights2Delete = True
         GoTo endProc
      End If
      lnSeek = MsgBox("User Doesn't Have the Right to Delete Transaction!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
      If lnSeek = vbNo Then GoTo endProc
      
      If p_oMod.GetApproval(p_oAppDrivr, lnUserRght, lsUserIDxx, lsUserName, lasRights(0)) = False Then GoTo endProc
   
      If (lnRights And lnUserRght) > 0 Then
         HasRights2Delete = True
         GoTo endProc
      End If
      
      MsgBox "Approving User Doesn't Have the Right to Delete Record!!!" & _
            "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
   End With
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Sub showRemarks(Optional lsRemarks As Variant, _
                        Optional lsMode As Variant)
   ' show remarks and mode in status bar
   With p_oAppDrivr
      If Not (.MDIMain Is Nothing) Then
         If Not IsMissing(lsRemarks) Then .MDIMain.Statusbar1.Panels(1).Text = lsRemarks
         If Not IsMissing(lsMode) Then .MDIMain.Statusbar1.Panels(2).Text = lsMode
      End If
   End With
End Sub

Function isModify() As Boolean
   Dim lsProcName As String
   
   lsProcName = "isModify"
   On Error GoTo errProc
   isModify = True
   
   If p_nEditMode = xeModeAddNew Then GoTo endProc
   
   For pnCtr = 0 To p_nFldCount - 2
      If p_oRecMain(FieldName(pnCtr)).Value <> FieldValue(pnCtr) Or _
         IsNull(p_oRecMain(FieldName(pnCtr)).Value) <> IsNull(FieldValue(pnCtr)) Then GoTo endProc
   Next
   
   isModify = False
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function getTableName(ByVal lsSQLQuery As String) As String
   Dim lsProcName As String
   Dim lnCtr As Long
   lsProcName = "getTableName"
   On Error GoTo errProc
   
   lsSQLQuery = Trim(Mid(lsSQLQuery, InStr(1, lsSQLQuery, "FROM ", vbTextCompare) + 4))
   
   lnCtr = InStr(1, lsSQLQuery, " ", vbTextCompare)
   If lnCtr > 1 Then
      lsSQLQuery = Trim(Left(lsSQLQuery, InStr(1, lsSQLQuery, " ", vbTextCompare)))
   End If
   getTableName = lsSQLQuery
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function
