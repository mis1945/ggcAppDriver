VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormDriverx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Values of HasTextBox
' 0 = NoTextBox
' 1 = HasTextbox
' 4 = picField
'Assumption was there are more than 1 record for each maintenance table

Const MODULENAME = "FormDriverx"

Private p_oFormObj As Object
Private p_oDataObj As Object
Private p_oAppDrivr As Object

Private p_aDataType() As kseDataType
Private p_aHasTextx() As String
Private p_aRequired() As String
Private p_aHelpText() As String
Private p_aFldSizex() As Integer
Private p_sPicFormt() As String

Private m_nColIndex As Integer
Private oFirst As TextBox

Public Event FieldExtract(ByVal Index As Variant, ByVal Value As Variant, Cancel As Boolean)
Public Event FieldsRetrieved(ByVal Index As Variant, ByVal Value As Variant)
Public Event DisableOtherControl()
Public Event EnableOtherControl()

Property Get DataType(Index As Integer) As String
   DataType = p_aDataType(Index)
End Property
Property Let DataType(Index As Integer, Value As String)
   p_aDataType(Index) = Value
End Property

Property Get HasTextBox(Index As Integer) As String
   HasTextBox = p_aHasTextx(Index)
End Property
Property Let HasTextBox(Index As Integer, Value As String)
   p_aHasTextx(Index) = Value
   If Value = "0" Then
       p_aFldSizex(Index) = 0
       p_aHelpText(Index) = ""
       p_sPicFormt(Index) = ""
   End If
End Property

Property Get IsRequired(Index As Integer) As String
   IsRequired = p_aRequired(Index)
End Property
Property Let IsRequired(Index As Integer, Value As String)
   p_aRequired(Index) = Value
End Property

Property Get HelpText(Index As Integer) As String
   HelpText = p_aHelpText(Index)
End Property
   
Property Get FieldSize(Index As Integer) As String
   FieldSize = p_aFldSizex(Index)
End Property
   
'Set the Picture format
Property Get PicFormat(Index As Integer) As String
   PicFormat = p_sPicFormt(Index)
End Property
Property Let PicFormat(Index As Integer, Value As String)
   p_sPicFormt(Index) = Value
End Property

'Current Working Index
Property Let ColumnIndex(Index As Integer)
   m_nColIndex = Index
End Property
Property Get ColumnIndex() As Integer
   ColumnIndex = m_nColIndex
End Property

'Control that will going to receive the first focus
Property Set FirstFocus(oFocus As Object)
   Set oFirst = oFocus
End Property

Sub InitForm(oShadow As Object, oForm As Object, oData As Object)
   Dim nCtr As Integer
   Dim lsProcName As String
   
   lsProcName = MODULENAME & "." & "InitForm(" _
              & " )"
   Debug.Print lsProcName
   'On Error Goto errProc
   
   showRemarks "Initializing objects...", ""
   
   Set p_oFormObj = oForm
   Set p_oDataObj = oData
   Set p_oAppDrivr = oShadow
   Call p_oDataObj.InitRecord(p_oAppDrivr)

   With p_oDataObj
      ReDim p_aDataType(.Columns)
      ReDim p_aHasTextx(.Columns)
      ReDim p_aRequired(.Columns)
      ReDim p_aHelpText(.Columns)
      ReDim p_aFldSizex(.Columns)
      ReDim p_sPicFormt(.Columns)
   
      For nCtr = 0 To .Columns
         p_aDataType(nCtr) = kseString
         p_aHasTextx(nCtr) = "1"
         p_aRequired(nCtr) = "1"
      
         'If not of blob type then set the field size
         If .DataType <> 205 Then
            p_aFldSizex(nCtr) = .DefineSize(nCtr)
         End If
      Next
   End With
   
   showRemarks "Objects Initialized...", "Unknown"
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName
End Sub

Sub NewRecord(Optional bFromFirst As Boolean = False)
   Dim lsProcName As String
   
   lsProcName = MODULENAME & "." & "NewRecord(" _
              & " )"
   Debug.Print lsProcName
   'On Error Goto errProc
   
   showRemarks "Allocating New Record..."
   If p_oDataObj.NewRecord Then
      Call ShowData
      Call hideButton(-1)
      Call EnableTextbox(-1)
      
      If bFromFirst = False Then
         oFirst.SetFocus
      End If
      showRemarks "New Record Allocated...", "New"
   Else
      showRemarks "Allocation failed..."
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsProcName
End Sub

Sub UpdateRecord()
   Dim lsProcName As String
   
   lsProcName = MODULENAME & "." & "UpdateRecord(" _
              & " )"
   Debug.Print lsProcName
   'On Error Goto errProc
   
   showRemarks "Preparing for update mode..."
   If p_oDataObj.UpdateRecord Then
'      Call ShowData
      Call hideButton(-1)
      Call EnableTextbox(-1)
      oFirst.SetFocus
      showRemarks "Update mode set...", "Update"
   Else
      showRemarks "Update mode failed!"
   End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName
End Sub

Sub BrowseRecord()
   Dim lsProcName As String
   
   lsProcName = MODULENAME & "." & "BrowseRecord(" _
              & " )"
   Debug.Print lsProcName
   'On Error Goto errProc

   p_oFormObj.Hide
   If p_oDataObj.BrowseRecord() Then
      p_oFormObj.Show
      Call ShowData
      Call showButton(-1)
      Call DisableTextbox(-1)
   End If
   
endProc:
   Exit Sub
errProc:
   ShowError lsProcName
End Sub

Sub UnSaveRecord()
   Dim oTaskMode As xeEditMode
   Dim lsProcName As String
   
   lsProcName = MODULENAME & "." & "UnsaveRecord(" _
              & " )"
   Debug.Print lsProcName
   'On Error Goto errProc
   
   oTaskMode = p_oDataObj.TaskMode
   showRemarks "Cancelling update..."
   
   p_oAppDrivr.BeginTrans
   If p_oDataObj.UnSaveRecord() Then
      p_oAppDrivr.CommitTrans
      Call ShowData
      Call showButton(-1)
      Call DisableTextbox(-1)
      showRemarks "Cancelling of update successfull!"
   Else
      p_oAppDrivr.RollBackTrans
      showRemarks "Cancelling of update failed!"
   End If
   
'   If oTaskMode = xeModeUpdate Then
'      Call ShowData
'      Call showButton(-1)
'      Call DisableTextbox(-1)
'   End If

endProc:
   Exit Sub
errProc:
   ShowError lsProcName
End Sub

Sub SaveRecord()
   Dim oTaskMode As xeEditMode
   Dim lsProcName As String
   
   lsProcName = MODULENAME & "." & "SaveRecord(" _
              & " )"
   Debug.Print lsProcName
   'On Error Goto errProc
   
   If ValidRecord Then
      oTaskMode = p_oDataObj.TaskMode
      showRemarks "Saving records..."
      p_oAppDrivr.BeginTrans
      If p_oDataObj.SaveRecord() Then
         showRemarks "Records saved!"
         p_oAppDrivr.CommitTrans
         Call ShowData
         Call showButton(-1)
         Call DisableTextbox(-1)
      Else
         p_oAppDrivr.RollBackTrans
         showRemarks "Saving failed..."
      End If
   Else
      MsgBox "Please verify your entry"
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsProcName
End Sub

Sub LookUp()
   Dim sAuto As Variant
   Dim lsProcName As String
   
   lsProcName = MODULENAME & "." & "LookUp(" _
              & " )"
   Debug.Print lsProcName
   'On Error Goto errProc
   
   sCode = p_oFormObj.txtField(ColumnIndex)
   With p_oDataObj
      Select Case .EntryType(ColumnIndex)
      Case 1 'Query
         .LookUp ColumnIndex, sCode, False, False, True
         RaiseEvent FieldsRetrieved(ColumnIndex, .Master(Index))
      Case 3 'Query and AutoLoad
         .LookUp ColumnIndex, sCode, False, False, True
         RaiseEvent FieldsRetrieved(ColumnIndex, .Master(Index))
         For Each sAuto In Split(.FieldsSet(Index), "»")
            RaiseEvent FieldsRetrieved(sAuto, .Master(sAuto))
         Next
      End Select
   End With
   
   p_oFormObj.txtField(ColumnIndex).SetFocus

endProc:
   Exit Sub
errProc:
   ShowError lsProcName
End Sub

Sub DeleteRecord()
   Dim lsProcName As String
   
   lsProcName = MODULENAME & "." & "DeleteRecord(" _
              & " )"
   Debug.Print lsProcName
   'On Error Goto errProc
   showRemarks "Deleting record...", "Delete"
   p_oAppDrivr.BeginTrans
   If p_oDataObj.DeleteRecord() Then
      p_oAppDrivr.CommitTrans
      showRemarks "Record deleted!", "Unknown"
      BrowseRecord
   Else
      p_oAppDrivr.RollBackTrans
      showRemarks "Delete failed!", "Ready"
   End If

endProc:
   Exit Sub
errProc:
   ShowError lsProcName
End Sub

Function ValidData(ByVal Index As Integer, ByVal Value As Variant) As Boolean
   Dim bCancel As Boolean
   Dim lsProcName As String
   Dim sAuto As Variant
   
   lsProcName = MODULENAME & "." & "ValidData(" _
              & "  " & Index _
              & ", " & Value _
              & " )"
   Debug.Print lsProcName
   'On Error Goto errProc
   
   'Make sure that we are at edit or add mode
   With p_oDataObj
      If p_oDataObj.TaskMode = xeModeUpdate Or _
         p_oDataObj.TaskMode = xeModeAddNew Then
         
         ValidData = ValidValue(Index, Value)
         If ValidData Then
            Select Case .EntryType(Index)
            Case 0 'Ordinary
               .Master(Index) = AssignValue(Index)
            Case 1 'Query
               .LookUp Index, Value, False, False
               RaiseEvent FieldsRetrieved(Index, .Master(Index))
            Case 2 'Ordinary and Auto Load
               .Master(Index) = AssignValue(Index)
               For Each sAuto In Split(.FieldsSet(Index), "»")
                  RaiseEvent FieldsRetrieved(sAuto, .Master(sAuto))
               Next
            Case 3 'Query and AutoLoad
               .LookUp Index, Value, False, False
               RaiseEvent FieldsRetrieved(Index, .Master(Index))
               For Each sAuto In Split(.FieldsSet(Index), "»")
                  RaiseEvent FieldsRetrieved(sAuto, .Master(sAuto))
               Next
            End Select
         End If
      End If
   End With

endProc:
   Exit Property
errProc:
   ShowError lsProcName
End Property

Function ValidRecord() As Boolean
   Dim nCtr As Integer
   Dim mystream As New ADODB.Stream
   Dim Cancel As Boolean
   Dim lsProcName As String
   
   Debug.Print "Form.BrowseRecord"
   lsProcName = MODULENAME & "." & "ValidRecord(" _
              & " )"
   Debug.Print lsProcName
   'On Error Goto errProc

   Set mystream = New ADODB.Stream
   mystream.Type = adTypeBinary
   
   With p_oDataObj
      'Perform validation here for those with textboxes
      For nCtr = 0 To .Columns
         If .IsRequired(nCtr) = "1" Then
            If HasTextBox(nCtr) = "1" Then
               If p_oFormObj.txtField(nCtr) = Empty Then
                  p_oFormObj.txtField(nCtr).SetFocus
                  GoTo endProc
               End If
            End If
         End If
         
         If HasTextBox(nCtr) = "0" Then
            'Get the different data in our form that are not stored in our standard textbox
            RaiseEvent FieldExtract(nCtr, Value, Cancel)
            If Cancel = True Then
               ValidRecord = False
               oFirst.SetFocus
               GoTo endProc
            End If
            .Master(nCtr) = Value
         ElseIf HasTextBox(nCtr) = "4" Then
            If Not IsNull(p_oFormObj.picField(nCtr).Picture) Then
               Call SavePicture(p_oFormObj.picField(nCtr).Picture, p_oAppDrivr.AppPath & "\system\xxx.pic")
               mystream.Open
               mystream.LoadFromFile p_oAppDrivr.AppPath & "\system\xxx.pic"
               .Master(nCtr) = mystream.Read
               mystream.Close
            End If
         End If
      
      Next
      'Perform Internal Validation inside the object
      nCtr = .ValidRecord
      
      If nCtr >= 0 Then
         p_oFormObj.txtField(nCtr).SetFocus
      Else
         ValidRecord = True
      End If
   
   End With

   Set mystream = Nothing
endProc:
   Exit Function
errProc:
   ShowError lsProcName
End Function

Sub ShowData()
   Dim nCtr As Integer
   Dim vData1 As Variant
   Dim vData2 As Variant
   Dim mystream As New ADODB.Stream
   Dim lsProcName As String
   
   lsProcName = MODULENAME & "." & "ShowData(" _
              & " )"
   Debug.Print lsProcName
   'On Error Goto errProc

   Set mystream = New ADODB.Stream
   mystream.Type = adTypeBinary
   
   showRemarks "Loading record..."
   With p_oDataObj
      For nCtr = 0 To .Columns
         Select Case HasTextBox(nCtr)
         Case "0"
            RaiseEvent FieldsRetrieved(nCtr, .Master(nCtr))
         Case "1"
            p_oFormObj.txtField(nCtr) = Format(.Master(nCtr), .PicFormat(nCtr))
         Case "4"
            mystream.Open
            mystream.Write .Master(nCtr)
            mystream.SaveToFile p_oAppDrivr.AppPath & "\system\xxx.pic", adSaveCreateOverWrite
            mystream.Close
            Set p_oFormObj.picField(nCtr).Picture = LoadPicture(p_oAppDrivr.AppPath & "\system\xxx.pic")
         End Select
      Next
   End With
   Set mystream = Nothing
   showRemarks "Record loaded!", "Ready"
endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Sub SetPicFormat()
   Dim nCtr As Integer
   For nCtr = 0 To Columns
      Select Case p_aDataType(nCtr)
      Case kseInteger
         p_sPicFormt(nCtr) = "#,##0"
      Case kseFloat, kseMoney
         p_sPicFormat(nCtr) = "#,##0.00"
      Case kseBlob
      Case kseTime
         p_sPicFormat(nCtr) = "HH:MM ampm"
      Case kseDate
         p_sPicFormat(nCtr) = "Mmm. DD, YYYY"
      Case kseDateTime
         p_sPicFormat(nCtr) = "MM/DD/YYYY HH:MM ampm"
      Case Else
         p_sPicFormat(nCtr) = "@"
      End Select
   Next
End Sub

Sub showButton(ByVal Index As Integer)
   Dim loButton As Control
   Dim lsProcName As String
   Dim lbShow As Boolean
   
   lsProcName = "ShowButton"
   'On Error Goto errProc
   
   With p_oFormObj
      If Index < 0 Then
         For Each loButton In .cmdButton
            Select Case UCase(loButton.Caption)
            Case "&UPDATE", "&DELETE", "E&XIT", "CANCEL &TRANS", "&PRINT"
               loButton.Visible = True
            Case "&SAVE", "&CANCEL SAVE", "&LOOKUP"
               loButton.Visible = False
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

Sub hideButton(ByVal Index As Integer)
   Dim loButton As Control
   Dim lsProcName As String
   Dim lbShow As Boolean
   
   lsProcName = "HideButton"
   'On Error Goto errProc
   
   With p_oFormObj
      If Index < 0 Then
         For Each loButton In .cmdButton
            Select Case UCase(loButton.Caption)
            Case "&UPDATE", "&DELETE", "E&XIT", "CANCEL &TRANS", "&PRINT"
               loButton.Visible = False
            Case "&SAVE", "&CANCEL SAVE", "&LOOKUP"
               loButton.Visible = True
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

Sub EnableTextbox(ByVal Index As Integer)
   Dim loText As TextBox
   Dim lsProcName As String
   
   lsProcName = "EnableTextBox"
   'On Error Goto errProc
   
   With p_oFormObj
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

Sub DisableTextbox(ByVal Index As Integer)
   Dim loText As TextBox
   Dim lsProcName As String
   
   lsProcName = "DisableTextBox"
   'On Error Goto errProc
   
   With p_oFormObj
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

Private Sub ShowError(ByVal lsProcName As String)
   With p_oAppDrivr
      .xLogError Err.Number, Err.Description, MODULENAME, lsProcName, Erl
      With Err
         .Raise .Number, .Source, .Description
      End With
   End With
End Sub

Private Function ValidValue(ByVal Index As Integer, ByVal Value As Variant) As Boolean
   If Value = "" Then
      ValidValue = True
   Else
      Select Case p_aDataType(Index)
      'Test for the validity of data according to type
      Case kseDate, kseTime, kseDateTime
         If IsDate(Value) Then
            ValidData = True
         End If
      Case kseInteger, kseFloat, kseMoney
         If IsNumeric(Value) Then
            ValidData = True
         End If
      Case Else
         ValidData = True
      End Select
      
      If ValidData Then
         p_oFormObj.txtField(Index) = Format(Value, .PicFormat(Index))
      End If
   End If

End Function

Private Function AssignValue(ByVal Index As Integer)
   Select Case p_aDataType(Index)
   Case kseDate, kseTime, kseDateTime
      AssignValue = CDate(p_oFormObj.txtField(Index))
   Case kseInteger, kseFloat, kseMoney
      AssignValue = CDbl(p_oFormObj.txtField(Index))
   Case Else
      AssignValue = p_oFormObj.txtField(Index)
   End Select
End Function

Private Sub showRemarks(Optional lsRemarks As Variant, _
                        Optional lsMode As Variant)
   If p_bShowRem = False Then Exit Sub
   
   ' show remarks and mode in status bar
   With p_oAppDrivr
      If Not (.MDIMain Is Nothing) Then
         If Not IsMissing(lsRemarks) Then .MDIMain.Statusbar1.Panels(1).Text = lsRemarks
         If Not IsMissing(lsMode) Then .MDIMain.Statusbar1.Panels(2).Text = lsMode
      End If
   End With
End Sub


