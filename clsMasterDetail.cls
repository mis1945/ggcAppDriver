VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMasterDetail"
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
'  XerSys [ 07/27/2004 02:31 am ]
'     Start creating this object.
'     This object is will be the base class of all transaction objects. This will handle
'        the following:
'           Retrieving transaction
'           Saving transaction
'           Posting transaction
'           Canceling transaction
'           Deleting transaction
'           Searching transaction
' ------------------------------------------------------------------------------------------
'  XerSys [ 12/18/2006 10:18 am ]
'     Create a second version of this object
'        -  Refer to GCC SEGs Notes for detailed changes of this object
' ------------------------------------------------------------------------------------------
'  Kalyptus [ 12/20/2006 11:20 am ]
'     Inserts BugTracking
' ------------------------------------------------------------------------------------------
'  XerSys [ 07/10/2007 04:15 pm ]
'     Convert the mode of updating transaction, use statement rather than recordset
'        to conform to the new replication procedure.
'  XerSys [ 09/21/2011 12:10 pm ]
'     Add special field for destination for transactions that are intended to specific
'        branch which is not the direct recipient of the said transactions.
'  XerSys [ 10/05/2011 10:14 pm ]
'     Incorporate validation in deleting transaction base on user level and date of entry
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private p_oAppDrivr As clsAppDriver
Private p_oMaster As ADODB.Recordset
Private p_oDetail As ADODB.Recordset
Private p_oMod As New clsMainModules

Private p_sBranchCd As String
Private p_sDestinat As String
Private p_sMaster As String
Private p_sDetail As String
Private p_sMasSQL As String
Private p_sLgrSQL As String
Private p_sRefFld As String
Private p_bShowRem As Boolean
Private p_bShowMsg As Boolean

Private p_asLstVal() As String
Private p_asLstSQL() As String
Private p_asLstCol() As String
Private p_asLstCri() As String
Private p_asLstTit() As String
Private p_asLstPic() As String

Private p_asLgrLstSQL() As String
Private p_asLgrLstArg() As String
Private p_asLgrLstCol() As String
Private p_asLgrLstCri() As String
Private p_asLgrLstTit() As String
Private p_asLgrLstPic() As String

Private p_asMasRefFld() As String
Private p_asLgrRefFld() As String

Private pnEditMode As xeEditMode
Private pbRecExist As Boolean
Private pbInitTran As Boolean
Private pbVerifyNo As Boolean
Private pnCtr As Long

Public Event WillDelete(Cancel As Boolean)
Public Event Delete(Deleted As Boolean)
Public Event DeleteOthers(Cancel As Boolean)
Public Event DeleteComplete()

Public Event WillSave(Cancel As Boolean)
Public Event Save(Saved As Boolean)
Public Event SaveOthers(Cancel As Boolean)
Public Event SaveComplete()

Public Event WillCancel(Cancel As Boolean)
Public Event Cancel(Cancelled As Boolean)
Public Event CancelOthers(Cancel As Boolean)
Public Event CancelComplete()

Public Event LoadOthers()
Public Event InitMaster()
Public Event InitDetail(ItemNo As Integer)

Public Event DetailRetrieved(ByVal Index As Integer)
Public Event MasterRetrieved(ByVal Index As Integer)

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Get Branch() As String
   Branch = p_sBranchCd
End Property

Property Let Branch(ByVal Value As String)
   p_sBranchCd = Value
End Property

Property Get Destination() As String
   Destination = p_sDestinat
End Property

Property Let Destination(ByVal Value As String)
   p_sDestinat = Value
End Property

Property Get MasterTable() As String
   MasterTable = p_sMaster
End Property

Property Let MasterTable(Value As String)
   p_sMaster = Value
End Property

Property Get MasterQuery() As String
   MasterQuery = p_sMasSQL
End Property

Property Let MasterQuery(Value As String)
   p_sMasSQL = Value
End Property

Property Get DetailTable() As String
   DetailTable = p_sDetail
End Property

Property Let DetailTable(Value As String)
   p_sDetail = Value
End Property

Property Get DetailQuery() As String
   DetailQuery = p_sLgrSQL
End Property

Property Let DetailQuery(Value As String)
   p_sLgrSQL = Value
End Property

Property Get BrowseDetailQuery(ByVal Index As Integer) As String
   BrowseDetailQuery = p_asLgrLstSQL(Index)
End Property

Property Let BrowseDetailQuery(ByVal Index As Integer, ByVal Value As String)
   p_asLgrLstSQL(Index) = Value
End Property

Property Get BrowseDetailColumn(ByVal Index As Integer) As String
   BrowseDetailColumn = p_asLgrLstCol(Index)
End Property

Property Let BrowseDetailColumn(ByVal Index As Integer, ByVal Value As String)
   p_asLgrLstCol(Index) = Value
End Property

Property Get BrowseDetailCriteria(ByVal Index As Integer) As String
   BrowseDetailCriteria = p_asLgrLstCri(Index)
End Property

Property Let BrowseDetailCriteria(ByVal Index As Integer, ByVal Value As String)
   p_asLgrLstCri(Index) = Value
End Property

Property Get BrowseDetailTitle(ByVal Index As Integer) As String
   BrowseDetailTitle = p_asLgrLstTit(Index)
End Property

Property Let BrowseDetailTitle(ByVal Index As Integer, ByVal Value As String)
   p_asLgrLstTit(Index) = Value
End Property

Property Get BrowseDetailFormat(ByVal Index As Integer) As String
   BrowseDetailFormat = p_asLgrLstPic(Index)
End Property

Property Let BrowseDetailFormat(ByVal Index As Integer, ByVal Value As String)
   p_asLgrLstPic(Index) = Value
End Property

Property Get BrowseQuery(ByVal Index As Integer) As String
   BrowseQuery = p_asLstSQL(Index)
End Property

Property Let BrowseQuery(ByVal Index As Integer, ByVal Value As String)
   p_asLstSQL(Index) = Value
End Property

Property Get BrowseColumn(ByVal Index As Integer) As String
   BrowseColumn = p_asLstCol(Index)
End Property

Property Let BrowseColumn(ByVal Index As Integer, ByVal Value As String)
   p_asLstCol(Index) = Value
End Property

Property Get BrowseCriteria(ByVal Index As Integer) As String
   BrowseCriteria = p_asLstCri(Index)
End Property

Property Let BrowseCriteria(ByVal Index As Integer, ByVal Value As String)
   p_asLstCri(Index) = Value
End Property

Property Get BrowseTitle(ByVal Index As Integer) As String
   BrowseTitle = p_asLstTit(Index)
End Property

Property Let BrowseTitle(ByVal Index As Integer, ByVal Value As String)
   p_asLstTit(Index) = Value
End Property

Property Get BrowseFormat(ByVal Index As Integer) As String
   BrowseFormat = p_asLstPic(Index)
End Property

Property Let BrowseFormat(ByVal Index As Integer, ByVal Value As String)
   p_asLstPic(Index) = Value
End Property

Property Let Reference(Value As String)
   p_sRefFld = Value
End Property

Property Let DisplayMessage(Value As Boolean)
   p_bShowMsg = Value
End Property

Property Get DisplayRemarks() As Boolean
   DisplayRemarks = p_bShowRem
End Property

Property Let DisplayRemarks(Value As Boolean)
   p_bShowRem = Value
End Property

Property Get Detail(ByVal Row As Long, ByVal Index As Variant) As Variant
   Dim lsProcName As String
   
   lsProcName = "(Get)Detail"
   'On Error GoTo errProc

   If pbInitTran = False Then Exit Property
   If Row > (p_oDetail.RecordCount - 1) Then Exit Property
   
   p_oDetail.Move Row, adBookmarkFirst
   If IsNumeric(Index) Then
      If Index = 0 Or p_asLgrLstSQL(Index) = Empty Then
         Detail = p_oDetail(Index).Value
      Else
         Detail = getDetDesc(Row, Index, p_oDetail(Index))
      End If
   Else
      Detail = p_oDetail(Index).Value
   End If
   
endProc:
   Exit Property
errProc:
   ShowError lsProcName & "( " & Row & ", " & Index & " )"
End Property

Property Let Detail(ByVal Row As Long, ByVal Index As Variant, Value As Variant)
   Dim lsProcName As String
   
   lsProcName = "(Let)Detail"
   'On Error GoTo errProc
   
   If pbInitTran = False Then Exit Property
   If Row > (p_oDetail.RecordCount - 1) Then Exit Property
   
   p_oDetail.Move Row, adBookmarkFirst
   If IsNumeric(Index) Then
      If Index = 0 Or p_asLgrLstSQL(Index) = Empty Then
         p_oDetail(Index) = Value
      Else
         getDetCode Row, Index, Value, False
      End If
   Else
      p_oDetail(Index) = Value
   End If

endProc:
   'iMac 2017.05.23
   'RaiseEvent DetailRetrieved(Index)
   Exit Property
errProc:
   ShowError lsProcName & "( " & Row & _
                              ", " & Index & _
                              ", " & Value & " )"
End Property

Property Get DetailOriginal(ByVal Row As Long, ByVal Index As Variant) As Variant
   Dim lsProcName As String
   
   lsProcName = "(Get)DetailOriginal"
   'On Error GoTo errProc
            
   If pbInitTran = False Then Exit Property
   If Row > (p_oDetail.RecordCount - 1) Then Exit Property
   
   p_oDetail.Move Row, adBookmarkFirst
   DetailOriginal = p_oDetail(Index).OriginalValue

endProc:
   Exit Property
errProc:
   ShowError lsProcName & "( " & Row & ", " & Index & " )"
End Property

Property Get DetFldSize(Index As Variant) As Integer
   If pbInitTran = False Then Exit Property
   
   If IsNumeric(Index) Then
      If p_asLgrLstSQL(Index) = Empty Then
         Select Case p_oDetail(Index).Type
         Case 129, 130, 202, 200
            DetFldSize = p_oDetail(Index).DefinedSize
         End Select
      End If
   End If
End Property

Property Get Master(Index As Variant) As Variant
   Dim lsProcName As String
   
   lsProcName = "(Get)Master"
   'On Error GoTo errProc
   
   If pbInitTran = False Then Exit Property
   If p_oMaster.EOF Then Exit Property
   
   If IsNumeric(Index) Then
      If Index = 0 Or p_asLstSQL(Index) = Empty Then
         Master = p_oMaster(Index)
      Else
         Master = p_asLstVal(Index)
      End If
   Else
      Master = p_oMaster(Index)
   End If

endProc:
   Exit Property
errProc:
   ShowError lsProcName & "( " & Index & " )"
End Property

Property Let Master(Index As Variant, Value As Variant)
   Dim lsProcName As String
   
   lsProcName = "(Let)Master"
   'On Error GoTo errProc
   
   If pbInitTran = False Then Exit Property
   If p_oMaster.EOF Then Exit Property
   
   If IsNumeric(Index) Then
      If Index = 0 Or p_asLstSQL(Index) = Empty Then
         p_oMaster(Index) = Value
      Else
         getMasCode Index, Value, False
      End If
   Else
      p_oMaster(Index) = Value
   End If

endProc:
   Exit Property
errProc:
   ShowError lsProcName & "( " & Index & ", " & Value & " )"
End Property

Property Get MasFldSize(Index As Variant) As Integer
   If pbInitTran = False Then Exit Property
   
   If IsNumeric(Index) Then
      If p_asLstSQL(Index) = Empty Then
         Select Case p_oMaster(Index).Type
         Case 129, 130, 202, 200
            MasFldSize = p_oMaster(Index).DefinedSize
         End Select
      End If
   End If
End Property

Property Get MasterValue(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   If p_oMaster.EOF Then Exit Property
   
   MasterValue = p_oMaster(Index)
End Property

Property Get MasterOriginal(Index As Variant) As Variant
   If pbInitTran = False Then Exit Property
   If p_oMaster.EOF Then Exit Property
   
   MasterOriginal = p_oMaster(Index).OriginalValue
End Property

Property Get ItemCount() As Long
   If pbInitTran = False Then Exit Property
   
   ItemCount = p_oDetail.RecordCount
End Property

Property Let EditMode(Value As xeEditMode)
   pnEditMode = Value
End Property

Property Get EditMode() As xeEditMode
   EditMode = pnEditMode
End Property

Property Let VerifyEntryNo(Value As Boolean)
   pbVerifyNo = Value
End Property

Property Let Sort(ByVal fFields As String)
   p_oDetail.Sort = fFields
End Property

Function InitTransaction() As Boolean
   Dim lasFieldRef() As String
   Dim lsProcName As String
   
   lsProcName = "InitTransaction"
   'On Error GoTo errProc

   ' first check the application driver
   If p_oMod.isAppDriverOK(p_oAppDrivr) = False Then
      MsgBox "Invalid Application Driver Detected!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If p_sBranchCd = Empty Then p_sBranchCd = p_oAppDrivr.BranchCode
   
   If p_sMaster = Empty Or p_sMasSQL = Empty Or _
         p_sDetail = Empty Or p_sLgrSQL = Empty Or _
         p_sRefFld = Empty Then
      MsgBox "Invalid Master-Detail Info Detected!!!", vbCritical, "Warning"
      GoTo endProc
   End If
      
   ' then check the reference field
   lasFieldRef = Split(p_sRefFld, "|")
   If UBound(lasFieldRef) <> 1 Then GoTo endProc
   
   p_asMasRefFld = Split(lasFieldRef(0), "»")
   p_asLgrRefFld = Split(lasFieldRef(1), "»")

   If UBound(p_asMasRefFld) <> UBound(p_asLgrRefFld) Then GoTo endProc
   
   ' now initialize the receiving recordset
   clearMaster
   
   ' after initializing the master, set master properties
   ReDim p_asLstVal(p_oMaster.Fields.Count - 1) As String
   ReDim p_asLstSQL(p_oMaster.Fields.Count - 1) As String
   ReDim p_asLstCol(p_oMaster.Fields.Count - 1) As String
   ReDim p_asLstTit(p_oMaster.Fields.Count - 1) As String
   ReDim p_asLstPic(p_oMaster.Fields.Count - 1) As String
   ReDim p_asLstCri(p_oMaster.Fields.Count - 1) As String
   
   ' after initializing the detail, set detail properties
   clearDetail
   ReDim p_asLgrLstSQL(p_oDetail.Fields.Count - 1) As String
   ReDim p_asLgrLstCol(p_oDetail.Fields.Count - 1) As String
   ReDim p_asLgrLstTit(p_oDetail.Fields.Count - 1) As String
   ReDim p_asLgrLstPic(p_oDetail.Fields.Count - 1) As String
   ReDim p_asLgrLstCri(p_oDetail.Fields.Count - 1) As String
   
   InitTransaction = True
   pbInitTran = True
   pnEditMode = xeModeReady
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function NewTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "NewTransaction"
   'On Error GoTo errProc
   NewTransaction = False

   If pbInitTran = False Then GoTo endProc
   
   showRemarks "Adding New Transaction..."
   clearMaster
   For pnCtr = 0 To UBound(p_asLstVal)
      p_asLstVal(pnCtr) = ""
   Next
   
   clearDetail
    
   pnEditMode = xeModeAddNew
   RaiseEvent InitMaster
   RaiseEvent InitDetail(0)
   showRemarks "", "New"
   
   NewTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function OpenTransaction(ByVal sReference As String) As Boolean
   Dim lsReference() As String
   Dim lsProcName As String
   Dim lsSQL As String
   Dim Cancel As Boolean
   Dim lnRecCount As Integer
   
   lsProcName = "OpenTransaction"
   'On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc
   
   lsReference = Split(sReference, "»")
   If UBound(lsReference) <> UBound(p_asMasRefFld) Then GoTo endProc
   
   With p_oAppDrivr
      showRemarks "Opening Transaction..."
   
      If p_oMaster.State = adStateOpen Then
         If Not p_oMaster.EOF Then
            If p_oMaster.EditMode <> adEditNone Then p_oMaster.CancelUpdate
         End If
         p_oMaster.Close
      End If
      
      ' set the select statement for the master
      lsSQL = p_sMasSQL
      For pnCtr = 0 To UBound(lsReference)
         If lsReference(pnCtr) <> Empty Then
            lsSQL = p_oMod.AddCondition(lsSQL, p_asMasRefFld(pnCtr) & _
                     " = " & p_oMod.strParm(lsReference(pnCtr)))
         End If
      Next
      
      p_oMaster.Open lsSQL, .Connection, , adLockOptimistic, adCmdText
      Debug.Print lsSQL
      Set p_oDetail = New Recordset
      ' set the select statement for the ledger
      lsSQL = p_sLgrSQL
      For pnCtr = 0 To UBound(lsReference)
         lsSQL = p_oMod.AddCondition(lsSQL, p_asLgrRefFld(pnCtr) & _
                  " = " & p_oMod.strParm(lsReference(pnCtr)))
      Next
      
      Debug.Print lsSQL
      p_oDetail.Open lsSQL, .Connection, adOpenKeyset, adLockBatchOptimistic, adCmdText
      
      
      ' joined table carry their individual table constraint, thus, if
      '  new record is added all constraints must be satisfied. With regards to
      '  this, a constraint violation occur, and update to the record will occur
      '  immediately. Thus an error will occur. Now deactive the connection in
      '  order to set the recordset to batch update form.
      Set p_oDetail.ActiveConnection = Nothing
      
      'Mac 2018-07-31
      '  this is the original location of this code.
      '  ill move it down just for GMC Baguio
      'lnRecCount = p_oDetail.RecordCount
      
      If p_oDetail.EOF Then
         p_oDetail.AddNew
         RaiseEvent InitDetail(0)
      End If
      
      'Mac 2018-07-31
      '  temporary solution
      lnRecCount = p_oDetail.RecordCount
   End With
   
   ' now verify if there is a record retrieved
   If p_oMaster.EOF Then
      MsgBox "No Record Pass the Given Criteria!!!", vbCritical, "Warning"
      showRemarks "Transaction was not Opened..."
      GoTo endProc
   End If
   
   ' after retrieving record, retrieve all description of refence field
   LoadDescription
   
   If pbVerifyNo Then
      If lnRecCount <> CInt(Str(p_oMod.IFNull(p_oMaster("nEntryNox"), 0))) Then
         MsgBox "Transaction Descripancy Detected!" & vbCrLf & vbCrLf & _
                  "Please Inform SEG/SSG of Guanzon Group of Companies!!!", vbCritical, "Warning"
         Debug.Print lsSQL
         GoTo endProc
      
      'Mac 2018-05-05 2:30PM
      'transferred this below to reinitialize the value of nEntryNox
      'in case nEntryNox = 0 and lnRecCount = 0
      'p_oMaster("nEntryNox") = p_oDetail.RecordCount
      End If
      
      p_oMaster("nEntryNox") = p_oDetail.RecordCount
   End If
   
   pbRecExist = True
   pnEditMode = xeModeReady
   
   RaiseEvent LoadOthers
   showRemarks "Transaction was Opened Successfully...", "Ready"
   
   OpenTransaction = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sReference & " )"
End Function

Function DeleteTransaction() As Boolean
   Dim lsMasSQL As String
   Dim lsDetSQL As String
   Dim lsField As String
   Dim lsProcName As String
   Dim lnCol As Long
   Dim lbCancel As Boolean
   
   lsProcName = "DeleteTransaction"
   'On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc
   
   If p_bShowMsg Then
      If MsgBox("Are you sure you want to Delete this Transaction?", _
                     vbQuestion + vbYesNo, "Confirm Cancel") <> vbYes Then GoTo endProc
   
      showRemarks "Deleting Transaction..."
      
      If HasRights2Delete = False Then
         showRemarks "Transaction was not Deleted."
         GoTo endProc
      End If

      ' »»» XerSys - 2011-10-05 10:18 am
      '  Incorporate the date transaction was created or its status in deleting transaction
      If (DateDiff("h", Master("dModified"), p_oAppDrivr.ServerDate) > 24) Then
         If Not p_oAppDrivr.UserLevel = xeEngineer Then
            If p_bShowMsg Then
               MsgBox "Unable to delete " & Master("sTransNox") & vbCrLf & _
                         "Transaction was created from the previous days...", vbCritical, "Warning"
            End If
            showRemarks "Transaction was not Deleted."
            GoTo endProc
         End If
      Else
         If p_oMod.IFNull(p_oAppDrivr.getConfiguration("cRealTime", p_oAppDrivr.BranchCode), "0") <> "0" Then
            If Master("cTranStat") <> xeStateOpen Then
               If p_bShowMsg Then
                  MsgBox "Unable to delete " & Master("sTransNox") & vbCrLf & _
                           "Only unconfirmed or unprinted transactions are allowable for this action!", vbCritical, "Warning"
               End If
               showRemarks "Transaction was not Deleted."
               GoTo endProc
            End If
         End If
      End If
   
      If pnEditMode <> xeModeReady Then
         MsgBox "Deletion of unsave transaction is not allowed!!!" & vbCrLf & vbCrLf & _
                  "Request can not be granted!!!", vbCritical, "Warning"
         showRemarks "Transaction was not Deleted."
         GoTo endProc
      End If
   End If
   
   If pbRecExist = False Then
      MsgBox "No transaction is available to delete!!!", vbCritical, "Warning"
      showRemarks "Transaction was not Deleted."
      GoTo endProc
   End If
   
   If isPosted Then
      MsgBox "Deletion of posted or cancelled transaction is not allowed!!!" & vbCrLf & vbCrLf & _
               "Request not granted!!!", vbCritical, "Warning"
      showRemarks "Transaction was not Deleted."
      GoTo endProc
   End If
   
   With p_oAppDrivr
      If p_bShowMsg Then .BeginTrans
   
      ' raise willdelete for verification of other info before actually deleting it
      lbCancel = False
      RaiseEvent WillDelete(lbCancel)
      If lbCancel Then
         showRemarks "Deleting Transaction was Aborted."
         GoTo endWithRoll
      End If
      
      RaiseEvent Delete(lbCancel)
      ' use to override the deletion of transaction
      If lbCancel Then GoTo DeleteOthers
      
      lsMasSQL = "DELETE FROM " & p_sMaster
      lsDetSQL = "DELETE FROM " & p_sDetail
      For pnCtr = 0 To UBound(p_asMasRefFld)
         lsField = p_asMasRefFld(pnCtr)
         lnCol = InStr(lsField, ".")
         If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
         lsMasSQL = p_oMod.AddCondition(lsMasSQL, lsField & " = " & _
                           p_oMod.strParm(p_oMaster(lsField).OriginalValue))
                           
         lsField = p_asLgrRefFld(pnCtr)
         lnCol = InStr(lsField, ".")
         If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
         lsDetSQL = p_oMod.AddCondition(lsDetSQL, lsField & " = " & _
                           p_oMod.strParm(p_oMaster(lsField).OriginalValue))
      Next
   
      pnCtr = .Execute(lsMasSQL, p_sMaster)
      If pnCtr <= 0 Then
         Err.Raise 32010, "Recordset", "Unable to delete record"
      End If
      
      pnCtr = .Execute(lsDetSQL, p_sDetail)
      If pnCtr <= 0 Then
         Err.Raise 32010, "Recordset", "Unable to delete record"
      End If
   End With
   
DeleteOthers:
   ' now allow user to delete other info
   RaiseEvent DeleteOthers(lbCancel)
   If lbCancel Then
      showRemarks "Transaction was not Deleted."
      GoTo endWithRoll
   End If
   If p_bShowMsg Then p_oAppDrivr.CommitTrans
   
   RaiseEvent DeleteComplete
   showRemarks "Transaction was Deleted Successfully."
   
   NewTransaction
   pbRecExist = False
   pnEditMode = xeModeReady
   
   DeleteTransaction = True
   
endProc:
   Exit Function
endWithRoll:
   If p_bShowMsg Then p_oAppDrivr.RollBackTrans
   GoTo endProc
errProc:
   If p_bShowMsg Then p_oAppDrivr.RollBackTrans
   ShowError lsProcName & "( " & " )"
End Function

Function CancelTransaction() As Boolean
   Dim lsProcName As String
   Dim lsField As String
   Dim lsSQL As String
   Dim lnCol As Long
   Dim lbCancel As Boolean
   
   lsProcName = "CancelTransaction"
   'On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc
   
   If p_bShowMsg Then
      If MsgBox("Are you sure you want to Cancel this Transaction?", _
                     vbQuestion + vbYesNo, "Confirm Cancel") <> vbYes Then GoTo endProc
                     
      showRemarks "Cancelling Transaction..."
      If HasRights2Cancel = False Then
         showRemarks "Transaction was not Cancelled."
         GoTo endProc
      End If
      
      If pnEditMode <> xeModeReady Then
         MsgBox "Cancellation of unsave transaction is not allowed!!!" & vbCrLf & vbCrLf & _
                  "Request can not be granted!!!", vbCritical, "Warning"
         showRemarks "Transaction was not Cancelled."
         GoTo endProc
      End If
   
   End If
   
   If pbRecExist = False Then
      MsgBox "No transaction is available to cancel!!!", vbCritical, "Warning"
      showRemarks "Transaction was not Cancelled."
      GoTo endProc
   End If
   
   With p_oAppDrivr
      If p_bShowMsg Then .BeginTrans
   
      ' raise willcancel for verification of other info before actually cancelling it
      lbCancel = False
      RaiseEvent WillCancel(lbCancel)
      If lbCancel Then
         showRemarks "Transaction was not Cancelled."
         GoTo endWithRoll
      End If
      
      RaiseEvent Cancel(lbCancel)
      If lbCancel Then GoTo CancelOthers
      
      lsSQL = "UPDATE " & p_sMaster & _
               " SET cTranStat = " & xeStateCancelled & _
                  ", sModified = " & p_oMod.strParm(.UserID)
      
      For pnCtr = 0 To UBound(p_asMasRefFld)
         lsField = p_asMasRefFld(pnCtr)
         lnCol = InStr(lsField, ".")
         If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
         lsSQL = p_oMod.AddCondition(lsSQL, lsField & " = " & _
                        p_oMod.strParm(p_oMaster(lsField)))
      Next

      pnCtr = .Execute(lsSQL, p_sMaster)
      If pnCtr <= 0 Then
         Err.Raise 32011, "Recordset", "Unable to update changes"
      End If
   
      p_oMaster("cTranStat") = xeStateCancelled
   End With
   
CancelOthers:
   RaiseEvent CancelOthers(lbCancel)
   If lbCancel Then
      showRemarks "Transaction was not Cancelled."
      GoTo endWithRoll
   End If
   
   If p_bShowMsg Then p_oAppDrivr.CommitTrans

   RaiseEvent CancelComplete
   showRemarks "Transaction was Cancelled Successfully."
   pnEditMode = xeModeReady
   
   CancelTransaction = True

endProc:
   Exit Function
endWithRoll:
   If p_bShowMsg Then p_oAppDrivr.RollBackTrans
   GoTo endProc
errProc:
   If p_bShowMsg Then p_oAppDrivr.RollBackTrans
   ShowError lsProcName & "( " & " )"
End Function

Function SaveTransaction() As Boolean
   Dim loMaster As ADODB.Recordset
   Dim loDetail As ADODB.Recordset
   Dim lsProcName As String
   Dim lsMasSQL As String
   Dim lsDetSQL As String
   Dim lsField As String
   Dim lnCol As Integer
   Dim dModified As Date
   Dim lbCancel As Boolean
   
   lsProcName = "SaveTransaction"
   'On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc
   
   If p_bShowMsg Then
      If pnEditMode = xeModeReady Then
         MsgBox "Saving of unmodified transaction is not allowed!!!", vbCritical, "Warning"
         GoTo endProc
      End If

      If MsgBox("Are you sure you want to Save this Transaction?", _
            vbYesNo + vbInformation, "Confirm") <> vbYes Then GoTo endProc
            
      showRemarks "Saving Transaction..."
      
      If HasRights2Save() = False Then
         showRemarks "Transaction was not saved."
         GoTo endProc
      End If
   End If
   
   If isRefEmpty Then
      MsgBox "Reference Fields can not Contain Empty Value!!!" & vbCrLf & _
               " Verify your Entry then Try Again!!!", vbCritical, "Warning"
      showRemarks "Transaction was not saved."
      GoTo endProc
   End If

   If p_bShowMsg Then p_oAppDrivr.BeginTrans
   
   lbCancel = False
   ' record validation must be in this event
   RaiseEvent WillSave(lbCancel)
   If lbCancel = True Then
      showRemarks "Saving Transaction was Aborted."
      GoTo endWithRoll
   End If

   dModified = p_oAppDrivr.ServerDate
   p_oMaster("sModified") = p_oAppDrivr.UserID
   
   lbCancel = False
   RaiseEvent Save(lbCancel)
   
   ' if record is saved by the client then endproc
   If lbCancel Then GoTo SaveOthers
   
   ' if Entry No needs to be verified, then assign field to master
   If pbVerifyNo Then
      p_oMaster("nEntryNox") = p_oDetail.RecordCount()
   End If
   
   ' detail table must be re-opened to include the actual fields only
   Set loMaster = New Recordset
   Set loDetail = New Recordset
   
   lsMasSQL = "SELECT * FROM " & p_sMaster
   lsDetSQL = "SELECT * FROM " & p_sDetail & " ORDER BY nEntryNox"
   
   If pnEditMode = xeModeAddNew Then
      lsMasSQL = p_oMod.AddCondition(lsMasSQL, "0 = 1")
      lsDetSQL = p_oMod.AddCondition(lsDetSQL, "0 = 1")

      Debug.Print lsMasSQL
      loMaster.Open lsMasSQL, p_oAppDrivr.Connection, , , adCmdText
      loDetail.Open lsDetSQL, p_oAppDrivr.Connection, , , adCmdText
            
      If AddNewTransaction(loMaster, loDetail) = False Then GoTo endWithRoll
   Else
      ' set the retrieval arguments of the recordset
      For pnCtr = 0 To UBound(p_asLgrRefFld)
         lsField = p_asLgrRefFld(pnCtr)
         lnCol = InStr(lsField, ".")
         If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
         lsDetSQL = p_oMod.AddCondition(lsDetSQL, lsField & " = " & _
                     p_oMod.strParm(p_oMaster(lsField).OriginalValue))
      
         lsField = p_asMasRefFld(pnCtr)
         lnCol = InStr(lsField, ".")
         If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
         lsMasSQL = p_oMod.AddCondition(lsMasSQL, lsField & " = " & _
                     p_oMod.strParm(p_oMaster(lsField).OriginalValue))
      Next
      loMaster.Open lsMasSQL, p_oAppDrivr.Connection, , , adCmdText
      Debug.Print lsMasSQL
      loDetail.Open lsDetSQL, p_oAppDrivr.Connection, , , adCmdText
      
      If UpdateOldTransaction(loMaster, loDetail) = False Then GoTo endWithRoll
   End If
   
SaveOthers:
   lbCancel = False
   RaiseEvent SaveOthers(lbCancel)
   If lbCancel Then GoTo endWithRoll
   
   If p_bShowMsg Then p_oAppDrivr.CommitTrans
   
   ' after saving the main table, trigger the savecomplete event
   RaiseEvent SaveComplete
   showRemarks "Transaction was Saved Successfully.", "Ready"
   
   pnEditMode = xeModeReady
   pbRecExist = True

   SaveTransaction = True
   
endProc:
   Set loDetail = Nothing
   Set loMaster = Nothing
   Exit Function
endWithRoll:
   If p_bShowMsg Then p_oAppDrivr.RollBackTrans
   GoTo endProc
errProc:
   MsgBox "ErrProc"
   If p_bShowMsg Then p_oAppDrivr.RollBackTrans
   ShowError lsProcName & "( " & " )"
End Function

Function PostTransaction(ByVal sReference As String) As Boolean
   Dim lors As Recordset
   Dim lsProcName As String
   Dim lasRef() As String
   Dim lsField As String
   Dim lsSQL As String
   Dim lsQry As String
   Dim lnCol As Long
   
   lsProcName = "PostTransaction"
   'On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc
   
   lasRef = Split(sReference, "»")
   If UBound(lasRef) <> UBound(p_asMasRefFld) Then GoTo endProc
   
   showRemarks "Posting Transaction..."
   
   lsSQL = "UPDATE " & p_sMaster & _
            " SET cTranStat = " & xeStatePosted & _
               ", sModified = " & p_oMod.strParm(p_oAppDrivr.UserID) & _
               ", dModified = " & p_oMod.dateParm(p_oAppDrivr.ServerDate)
   
   lsQry = "SELECT cTranStat FROM " & p_sMaster
   
   For pnCtr = 0 To UBound(p_asMasRefFld)
      lsField = p_asMasRefFld(pnCtr)
      lnCol = InStr(lsField, ".")
      If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
      
      lsSQL = p_oMod.AddCondition(lsSQL, lsField & " = " & p_oMod.strParm(lasRef(pnCtr)))
      lsQry = p_oMod.AddCondition(lsQry, lsField & " = " & p_oMod.strParm(lasRef(pnCtr)))
   Next
   
   Set lors = New Recordset
   lors.Open lsQry, p_oAppDrivr.Connection, , , adCmdText
   If lors.EOF Then
'      Err.Raise 32011, "Recordset", "Unable to update changes"
      GoTo endProc
   ElseIf lors("cTranStat") = xeStatePosted Or lors("cTranStat") = xeStateCancelled Then
      MsgBox "Modification of posted or cancelled transaction is not allowed!!!" & vbCrLf & vbCrLf & _
               "Request not granted!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   If Trim(p_sDestinat) <> "" Then
      pnCtr = p_oAppDrivr.Execute(lsSQL, p_sMaster, , p_sDestinat)
   Else
      pnCtr = p_oAppDrivr.Execute(lsSQL, p_sMaster, , Left(sReference, Len(p_sBranchCd)))
   End If
   
   If pnCtr <= 0 Then
      Err.Raise 32011, "Recordset", "Unable to update changes"
      GoTo endProc
   End If
   
   p_oMaster("cTranStat") = xeStatePosted
   
   showRemarks "Transaction was Posted Successfully."
   
   PostTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sReference & " )"
End Function

Function CloseTransaction(ByVal sReference As String) As Boolean
   Dim lors As Recordset
   Dim lasRef() As String
   Dim lsField As String
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lsQry As String
   Dim lnCol As Long
   
   lsProcName = "CloseTransaction"
   'On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc
   
   lasRef = Split(sReference, "»")
   If UBound(lasRef) <> UBound(p_asMasRefFld) Then GoTo endProc
   
   showRemarks "Closing Transaction..."
   
   lsSQL = "UPDATE " & p_sMaster & _
            " SET cTranStat = " & xeStateClosed & _
               ", sModified = " & p_oMod.strParm(p_oAppDrivr.UserID) & _
               ", dModified = " & p_oMod.dateParm(p_oAppDrivr.ServerDate) & _
            " WHERE cTranStat = " & p_oMod.strParm(xeStateOpen)
               
   lsQry = "SELECT cTranStat FROM " & p_sMaster
   
   For pnCtr = 0 To UBound(p_asMasRefFld)
      lsField = p_asMasRefFld(pnCtr)
      lnCol = InStr(lsField, ".")
      If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
      
      lsSQL = p_oMod.AddCondition(lsSQL, lsField & " = " & p_oMod.strParm(lasRef(pnCtr)))
      lsQry = p_oMod.AddCondition(lsQry, lsField & " = " & p_oMod.strParm(lasRef(pnCtr)))
   Next
   
   Set lors = New Recordset
   lors.Open lsQry, p_oAppDrivr.Connection, , , adCmdText
   If lors.EOF Then
      Err.Raise 32011, "Recordset", "Unable to update changes"
      GoTo endProc
   ElseIf lors("cTranStat") = xeStatePosted Or lors("cTranStat") = xeStateCancelled Then
      MsgBox "Modification of posted or cancelled transaction is not allowed!!!" & vbCrLf & vbCrLf & _
               "Request not granted!!!", vbCritical, "Warning"
      GoTo endProc
   ElseIf lors("cTranStat") = xeStateClosed Then
      CloseTransaction = True
      GoTo endProc
   End If
   
   If Trim(p_sDestinat) <> "" Then
      pnCtr = p_oAppDrivr.Execute(lsSQL, p_sMaster, , p_sDestinat)
   Else
      pnCtr = p_oAppDrivr.Execute(lsSQL, p_sMaster, , Left(sReference, Len(p_sBranchCd)))
   End If
   
   If pnCtr = 0 Then
      Err.Raise 32011, "Recordset", "Unable to update changes"
      GoTo endProc
   End If
   
   p_oMaster("cTranStat") = xeStateClosed
   showRemarks "Transaction was Closed Successfully."
   
   CloseTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & sReference & " )"
End Function

Function SearchMaster(ByVal lnIndex As Integer, _
                      Optional lsValue As Variant = "") As Boolean
   Dim lsProcName As String
   
   lsProcName = "SearchMaster"
   'On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc
   If p_oMaster.EOF Then GoTo endProc
   
   If p_asLstSQL(lnIndex) <> Empty Then
      getMasCode lnIndex, lsValue, True
   End If
   SearchMaster = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & lnIndex _
                       & ", " & lsValue & " )"
End Function

Function SearchDetail(ByVal nRow As Integer, _
                      ByVal nIndex As Integer, _
                      Optional lsValue As Variant = "") As Boolean
   Dim lsProcName As String
   
   lsProcName = "SearchDetail"
   'On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc
   If p_oMaster.EOF Then GoTo endProc
   
   If p_asLgrLstSQL(nIndex) <> Empty Then
      p_oDetail.Move nRow, adBookmarkFirst
      getDetCode nRow, nIndex, lsValue, True
   End If
   
   SearchDetail = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & nRow _
                       & ", " & nIndex _
                       & ", " & lsValue & " )"
End Function

Function SearchTransaction(ByVal sRefValue As String, _
                           ByVal sRefField As String, _
                           sDescript As String) As Boolean
   Dim loLookUp As Object
   Dim lors As ADODB.Recordset
   Dim lsProcName As String
   Dim lasValue() As String
   Dim lasField() As String
   Dim lsRef As String
   Dim lsSQL As String
   
   lsProcName = "SearchTransaction"
   'On Error GoTo errProc

   If pbInitTran = False Then GoTo endProc
   
   If sRefValue = Empty Then
      showRemarks "Browsing Transaction...", "Ready"
      Set loLookUp = New frmLookUp1
      Set loLookUp.AppDriver = p_oAppDrivr
      
      Debug.Print p_asLstSQL(0)
      loLookUp.SQLSource = p_asLstSQL(0)
      loLookUp.FldName = p_asLstCol(0)
      loLookUp.FldTitle = p_asLstTit(0)
      loLookUp.FldFormat = p_asLstPic(0)
      loLookUp.FldCriteria = p_asLstCri(0)
      
      loLookUp.Show 1
      lsRef = loLookUp.SelectedItem
   
      If lsRef <> Empty Then
         lasValue = Split(lsRef, "»")
         
         lsRef = ""
         For pnCtr = 0 To UBound(p_asMasRefFld)
            lsRef = lsRef & lasValue(pnCtr) & "»"
         Next
         lsRef = Left(lsRef, Len(lsRef) - 1)
      End If
   Else
      lasValue = Split(sRefValue, "»")
      lasField = Split(sRefField, "»")
      If UBound(lasValue) <> UBound(lasField) Then
         showRemarks "Invalid parameter detected."
         GoTo endProc
      End If

      lsSQL = p_asLstSQL(0)
      Debug.Print lsSQL
      For pnCtr = 0 To UBound(lasField)
         If lasValue(pnCtr) <> Empty Then
            If InStr(1, sDescript, lasField(pnCtr)) > 0 Then
               lsSQL = p_oMod.AddCondition(lsSQL, lasField(pnCtr) & " LIKE " & _
                        p_oMod.strParm(lasValue(pnCtr) & "%"))
            Else
               lsSQL = p_oMod.AddCondition(lsSQL, lasField(pnCtr) & " = " & _
                        p_oMod.strParm(lasValue(pnCtr)))
            End If
         End If
      Next
      
      Set lors = New Recordset
      lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockReadOnly, adCmdText
      Debug.Print lsSQL
      If lors.EOF Then
         showRemarks "No Record Exists."
         GoTo endProc
      End If
      
      If lors.RecordCount > 1 Then
         Set loLookUp = New frmLookUp
         Set loLookUp.AppDriver = p_oAppDrivr
      
         Set loLookUp.RowSource = lors
         loLookUp.FldName = p_asLstCol(0)
         loLookUp.FldTitle = p_asLstTit(0)
         loLookUp.FldFormat = p_asLstPic(0)
         
         loLookUp.Show 1
         lsRef = loLookUp.SelectedItem
         
         If lsRef <> Empty Then
            lasValue = Split(lsRef, "»")
            
            lsRef = ""
            For pnCtr = 0 To UBound(p_asMasRefFld)
               lsRef = lsRef & lasValue(pnCtr) & "»"
            Next
            lsRef = Left(lsRef, Len(lsRef) - 1)
         End If
      Else
         lsRef = ""
         For pnCtr = 0 To UBound(p_asMasRefFld)
            lsRef = lsRef & lors.Fields(pnCtr) & "»"
         Next
         lsRef = Left(lsRef, Len(lsRef) - 1)
      End If
   End If
   
   If lsRef = Empty Then
      showRemarks "No selection was made..."
      GoTo endProc
   End If
   
   SearchTransaction = OpenTransaction(lsRef)
endProc:
   Set loLookUp = Nothing
   Set lors = Nothing

   Exit Function
errProc:
   ShowError lsProcName & "( " & sRefValue _
                       & ", " & sRefField _
                       & ", " & sDescript & " )"
End Function

Function ClearTransaction() As Boolean
   Dim lsProcName As String
   
   lsProcName = "ClearTransaction"
   'On Error GoTo errProc
   ClearTransaction = False

   If pbInitTran = False Then GoTo endProc
   
   clearMaster
   clearDetail
   
   ClearTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function AddDetail() As Boolean
   Dim lsProcName As String
   
   lsProcName = "AddDetail"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   p_oDetail.AddNew
   
   AddDetail = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function DeleteDetail(Index As Long) As Boolean
   Dim lsProcName As String
   
   lsProcName = "DeleteDetail"
   'On Error GoTo errProc
   
   If pbInitTran = False Then GoTo endProc
   If p_oDetail.RecordCount - 1 < Index Then GoTo endProc
   
   p_oDetail.Move Index, adBookmarkFirst
   p_oDetail.Delete adAffectCurrent
   
   DeleteDetail = True

endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & Index & " )"
End Function

Function GetMaster(ByVal Index As Integer, _
                   Condition As String, _
                   Optional Search As Variant = False) As String
   Dim lors As ADODB.Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lasSelect() As String
   
   lsProcName = "GetMaster"
   'On Error GoTo errProc

   If p_oMaster.Fields.Count - 1 < Index Then GoTo endProc
   lsSQL = p_oMod.AddCondition(p_asLstSQL(Index), Condition)
   Set lors = New Recordset
   
   Debug.Print lsSQL
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText
   
   If lors.EOF Then GoTo endProc
   If lors.RecordCount = 1 Then
      If Search = True Then
         GetMaster = p_oMod.KwikBrowse(p_oAppDrivr, lors, p_asLstCol(Index), p_asLstTit(Index), _
                     p_asLstPic(Index), p_asLstCri(Index))
      Else
         For pnCtr = 0 To lors.Fields.Count - 1
            GetMaster = GetMaster & IIf(IsNull(lors(pnCtr)), "", lors(pnCtr)) & "»"
         Next
         GetMaster = Left(GetMaster, Len(GetMaster) - 1)
      End If
   Else
      GetMaster = p_oMod.KwikBrowse(p_oAppDrivr, lors, p_asLstCol(Index), p_asLstTit(Index), _
                     p_asLstPic(Index), p_asLstCri(Index))
   End If
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & Index _
                       & ", " & Condition _
                       & ", " & Search & " )"
End Function

Function GetDetail(ByVal Row As Integer, _
                   Condition As String, _
                   Optional Search As Variant = False) As String
   Dim lors As ADODB.Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lasSelect() As String

   lsProcName = "GetDetail"
   'On Error GoTo errProc
   If p_oDetail.RecordCount - 1 < Row Then GoTo endProc
   If LCase(Left(Trim(Condition), 6)) = "having" Then
      lsSQL = p_asLgrLstSQL(0) & Condition
   Else
      lsSQL = p_oMod.AddCondition(p_asLgrLstSQL(0), Condition)
   End If

   Set lors = New Recordset
   Debug.Print lsSQL
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenKeyset, adLockOptimistic, adCmdText

   If lors.EOF Then GoTo endProc
   If lors.RecordCount = 1 Then
      For pnCtr = 0 To lors.Fields.Count - 1
         GetDetail = GetDetail & IIf(IsNull(lors(pnCtr)), "", lors(pnCtr)) & "»"
      Next
      GetDetail = Left(GetDetail, Len(GetDetail) - 1)
   Else
      GetDetail = p_oMod.KwikBrowse(p_oAppDrivr, lors, p_asLgrLstCol(0), p_asLgrLstTit(0), p_asLgrLstPic(0), p_asLgrLstCri(0))
   End If

endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & Row _
                       & ", " & Condition _
                       & ", " & Search & " )"
End Function


Function GetNextTrans(ByVal Field As String, _
                      ByVal Condition As String) As String
   Dim lors As ADODB.Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lnCode As Long
   Dim lnLen As Long

   lsProcName = "GetNextTrans"
   'On Error GoTo errProc
   
   lsSQL = "SELECT" & Field & _
      " FROM " & p_sMaster & _
      " WHERE " & Condition & _
      " ORDER BY " & Field & " DESC" & _
      " LIMIT 1"
      
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

   lsSQL = IIf(lors.EOF, Empty, lors(0))
   lnLen = lors(0).DefinedSize
   lnCode = 1
   If lsSQL <> Empty Then
      If Left(lsSQL, 2) = Format(Date, "YY") Then
         lnCode = CLng(Mid(lsSQL, 3)) + 1
      End If
   End If
   
   GetNextTrans = Format(Date, "yy") & Format(lnCode, String(lnLen - 2, "0"))

endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & Field _
                       & ", " & Condition & " )"
End Function

Sub clearMaster()
   Dim lsProcName As String
   
   lsProcName = "clearMaster"
   'On Error GoTo errProc
   
'   If p_oMaster.State = adStateOpen Then
'      If p_oMaster.EditMode <> adEditNone Then p_oMaster.CancelUpdate
'      p_oMaster.Close
'   End If

   Set p_oMaster = New Recordset
   Debug.Print p_sMasSQL
   p_oMaster.Open p_oMod.AddCondition(p_sMasSQL, "0 = 1"), p_oAppDrivr.Connection, adOpenStatic, _
                       adLockOptimistic, adCmdText
   p_oMaster.AddNew

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Sub clearDetail()
   Dim lsProcName As String
   
   lsProcName = "clearDetail"
   'On Error GoTo errProc
   
'   If p_oDetail.State = adStateOpen Then
'      If p_oDetail.EOF Then p_oDetail.Move 0, adBookmarkFirst
'      If p_oDetail.EditMode <> adEditNone Then p_oDetail.CancelUpdate
'      p_oDetail.Close
'   End If
   Set p_oDetail = New Recordset
   Debug.Print p_sLgrSQL
   p_oDetail.Open p_oMod.AddCondition(p_sLgrSQL, "0 = 1"), p_oAppDrivr.Connection, adOpenStatic, _
                      adLockBatchOptimistic, adCmdText
   Set p_oDetail.ActiveConnection = Nothing
   p_oDetail.AddNew

endProc:
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Function isDetailModify() As Boolean
   Dim lors As ADODB.Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lsField As String
   Dim lsColumn As String
   Dim lnCol As Integer
   Dim lnRow As Long
   Dim lvValue1 As Variant
   Dim lvValue2 As Variant
   
   lsProcName = "isDetailModify"
   'On Error GoTo errProc
   isDetailModify = True
   
   If pnEditMode = xeModeAddNew Then GoTo endProc
   
   ' set the retrieval arguments of the recordset
   lsSQL = p_sLgrSQL
   For pnCtr = 0 To UBound(p_asLgrRefFld)
      lsField = p_asLgrRefFld(pnCtr)
      lsColumn = lsField
      lnCol = InStr(lsColumn, ".")
      If lnCol > 0 Then lsColumn = Mid(lsColumn, lnCol + 1)
      lsSQL = p_oMod.AddCondition(lsSQL, lsField & " = " & _
                  p_oMod.strParm(p_oMaster(lsColumn).Value))
   Next
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
   lnRow = 0
   If p_oDetail.RecordCount > 0 Then p_oDetail.MoveFirst
   lsField = p_asLgrRefFld(0)
   lnCol = InStr(lsField, ".")
   If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
   For pnCtr = 0 To p_oDetail.RecordCount - 1
      If IsNull(p_oDetail(lsField)) Then Exit For
      If p_oDetail(lsField) = Empty Then Exit For
      lnRow = pnCtr + 1
      p_oDetail.MoveNext
   Next
   
   If lors.RecordCount <> lnRow Then GoTo endProc
   
   p_oDetail.MoveFirst
   For pnCtr = 0 To lnRow - 1
      For lnCol = 0 To lors.Fields.Count - 1
         If p_oDetail(lnCol).Value <> p_oMod.IFNull(lors(lnCol).Value, "") Then GoTo endProc
      Next
      lors.MoveNext
      p_oDetail.MoveNext
   Next
   
   isDetailModify = False
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function isMasterModify() As Boolean
   Dim lors As ADODB.Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   Dim lsField As String
   Dim lnCol As Integer
   Dim lnRow As Long
   
   lsProcName = "isMasterModify"
   'On Error GoTo errProc
   isMasterModify = True
   
   If pnEditMode = xeModeAddNew Then GoTo endProc
   
   ' set the retrieval arguments of the recordset
   lsSQL = p_sMasSQL
   For pnCtr = 0 To UBound(p_asMasRefFld)
      lsField = p_asMasRefFld(pnCtr)
      lnCol = InStr(lsField, ".")
      If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
      lsSQL = p_oMod.AddCondition(lsSQL, p_asMasRefFld(pnCtr) & " = " & _
                  p_oMod.strParm(p_oMaster(lsField).Value))
   Next
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, adOpenStatic, adLockOptimistic, adCmdText
   
'   Kalyptus(2008-04-22)
'   Added this code to allow saving of a new record.
'   Situation: Update MCSales and change the status of MC Giveaway to
'     Replace. Without this lines of code that would not be possible since
'     the edit mode is xeModeReady and this module assumes that all updates
'     has a record...
'   ==================================================
   If lors.EOF Then
      pnEditMode = xeModeAddNew
      isMasterModify = True
      GoTo endProc
   End If
      
   For lnCol = 0 To lors.Fields.Count - 1
      If p_oMaster(lnCol) <> p_oMod.IFNull(lors(lnCol), "") Then GoTo endProc
   Next
      
   isMasterModify = False
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function isPosted() As Boolean
   isPosted = True
   If Not (p_oMaster("cTranStat") = xeStatePosted Or _
         p_oMaster("cTranStat") = xeStateCancelled) Then isPosted = False
End Function

Sub ViewUserModify()
   Dim lors As ADODB.Recordset
   Dim lsProcName As String
   Dim lsSQL As String
   
   lsProcName = "ViewUserModify"
   'On Error GoTo errProc
   
   lsSQL = "SELECT sUserName" & _
            " FROM xxxSysUser" & _
            " WHERE sUserIDxx = " & p_oMod.strParm(Master("sModified"))

   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
   
   If lors.EOF Then
      MsgBox "User does not Exist in the User Table!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   MsgBox "    ««« LAST MODIFICATION »»»" & vbCrLf & vbCrLf & _
            "User Name      : " & p_oMod.Decrypt(lors("sUserName"), p_oAppDrivr.Machinex) & vbCrLf & _
            "Date Modified : " & Format(Master("dModified"), "MMMM DD, YYYY HH:NN:SS") _
            , vbInformation, "Notice"
         
endProc:
   Set lors = Nothing
   Exit Sub
errProc:
   ShowError lsProcName & "( " & " )"
End Sub

Private Function isRefEmpty() As Boolean
   Dim lsField As String
   Dim lnCol As String
   
   isRefEmpty = True
   For pnCtr = 0 To UBound(p_asMasRefFld)
      lsField = p_asLgrRefFld(pnCtr)
      lnCol = InStr(lsField, ".")
      If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
      If p_oMaster(lsField) = Empty Then Exit Function
   Next
   isRefEmpty = False
End Function

Private Sub ShowError(ByVal lsProcName As String)
    With p_oAppDrivr
       .xLogError Err.Number, Err.Description, "clsMasterDetail", lsProcName, Erl
    End With
    With Err
       .Raise .Number, .Source, .Description
    End With
End Sub

Private Sub Class_Initialize()
   Set p_oMaster = New Recordset
   p_oMaster.CursorLocation = adUseClient
   Set p_oDetail = New Recordset
   p_oMaster.CursorLocation = adUseClient

   p_bShowMsg = True
   p_bShowRem = True
End Sub

Private Sub Class_Terminate()
   Set p_oMaster = Nothing
   Set p_oDetail = Nothing
   Set p_oMod = Nothing
End Sub

Private Sub getMasCode(ByVal lnIndex As Integer, _
                       ByVal lsValue As String, _
                       ByVal lbSearch As Boolean)
   Dim lors As ADODB.Recordset
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lasRef() As String

   lsProcName = "getMasCode"
   'On Error GoTo errProc

   If lsValue = p_asLstVal(lnIndex) And lbSearch = False Then GoTo endProc
   
   If p_asLstCri(lnIndex) = Empty Then
      If p_asLstCol(lnIndex) = Empty Then
         GoTo endProc
      Else
         p_asLstCri(lnIndex) = p_asLstCol(lnIndex)
      End If
   End If
      
   lasRef = Split(p_asLstCri(lnIndex), "»")
                  
   lsSQL = p_asLstSQL(lnIndex)
   
   If lsValue <> "" Then
      lsSQL = p_oMod.AddCondition(lsSQL, lasRef(1) & _
            IIf(lbSearch, " Like " & p_oMod.strParm(lsValue & "%"), _
               " = " & p_oMod.strParm(lsValue)))
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
   
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
                     
   If lors.EOF Then GoTo endWithClear
   If lors.RecordCount = 1 Then
      p_oMaster(lnIndex) = lors(0)
      p_asLstVal(lnIndex) = lors(1)
      RaiseEvent MasterRetrieved(lnIndex)
      GoTo endProc
   End If
   
   lsSQL = p_oMod.KwikBrowse(p_oAppDrivr, lors, p_asLstCol(lnIndex), p_asLstTit(lnIndex), p_asLstPic(lnIndex), p_asLstCri(lnIndex))
   
   If lsSQL = "" Then GoTo endProc
   lasRef = Split(lsSQL, "»")
   p_oMaster(lnIndex) = lasRef(0)
   p_asLstVal(lnIndex) = lasRef(1)
   RaiseEvent MasterRetrieved(lnIndex)
   
endProc:
   Set lors = Nothing
   Exit Sub
endWithClear:
   p_oMaster(lnIndex) = ""
   p_asLstVal(lnIndex) = ""
   RaiseEvent MasterRetrieved(lnIndex)
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & lnIndex _
                       & ", " & lsValue _
                       & ", " & lbSearch & " )"
End Sub

Private Sub getMasDesc(ByVal lnIndex As Integer, _
                       ByVal lsValue As String)
   Dim lors As ADODB.Recordset
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lasRef() As String
   Dim lnCtr As Integer

   lsProcName = "getMasDesc"
   'On Error GoTo errProc
      
   If lsValue = "" Then GoTo endProc
   
   If p_asLstCri(lnIndex) = Empty Then
      If p_asLstCol(lnIndex) = Empty Then
         GoTo endProc
      Else
         p_asLstCri(lnIndex) = p_asLstCol(lnIndex)
      End If
   End If
      
   p_asLstVal(lnIndex) = ""
      
   lasRef = Split(p_asLstCri(lnIndex), "»")
                  
   lsSQL = p_asLstSQL(lnIndex)
   lsSQL = p_oMod.AddCondition(lsSQL, lasRef(0) & " = " & p_oMod.strParm(lsValue))
   
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
                     
   If lors.EOF Then GoTo endProc
   p_asLstVal(lnIndex) = lors(1)
   
endProc:
   Set lors = Nothing
   Exit Sub
errProc:
   ShowError lsProcName & "( " & lnIndex _
                       & ", " & lsValue & " )"
End Sub

Private Function getDetDesc(ByVal lnRow As Integer, _
                            ByVal lnIndex As Variant, _
                            ByVal lsValue As String) As String
   Dim lors As ADODB.Recordset
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lasRef() As String
   Dim lnCtr As Integer

   lsProcName = "getDetDesc"
   'On Error GoTo errProc
   
   If lsValue = "" Then GoTo endProc
   
   If p_asLgrLstCri(lnIndex) = Empty Then
      If p_asLgrLstCol(lnIndex) = Empty Then
         GoTo endProc
      Else
         p_asLgrLstCri(lnIndex) = p_asLgrLstCol(lnIndex)
      End If
   End If
      
   lasRef = Split(p_asLgrLstCri(lnIndex), "»")
                  
   lsSQL = p_asLgrLstSQL(lnIndex)
   lsSQL = p_oMod.AddCondition(lsSQL, lasRef(0) & " = " & p_oMod.strParm(lsValue))
   
   Set lors = New Recordset
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText

   If lors.EOF Then GoTo endProc
   getDetDesc = lors(1)
   
endProc:
   Set lors = Nothing
   Exit Function
errProc:
   ShowError lsProcName & "( " & lnRow _
                       & ", " & lnIndex _
                       & ", " & lsValue & " )"
End Function

Private Sub getDetCode(ByVal lnRow As Integer, _
                       ByVal lnIndex As Integer, _
                       ByVal lsValue As String, _
                       ByVal lbSearch As Boolean)
   Dim lors As ADODB.Recordset
   Dim lsSQL As String
   Dim lsProcName As String
   Dim lasRef() As String

   lsProcName = "getDetCode"
   'On Error GoTo errProc
   
   If p_asLgrLstCri(lnIndex) = Empty Then
      If p_asLgrLstCol(lnIndex) = Empty Then
         GoTo endProc
      Else
         p_asLgrLstCri(lnIndex) = p_asLgrLstCol(lnIndex)
      End If
   End If
      
   p_oDetail.Move lnRow, adBookmarkFirst
   lasRef = Split(p_asLgrLstCri(lnIndex), "»")
                  
   lsSQL = p_asLgrLstSQL(lnIndex)
   
   If lsValue <> "" Then
      lsSQL = p_oMod.AddCondition(lsSQL, lasRef(1) & _
            IIf(lbSearch, " Like " & p_oMod.strParm(lsValue & "%"), _
               " = " & p_oMod.strParm(lsValue)))
   ElseIf lbSearch = False Then
      GoTo endWithClear
   End If
   
   Set lors = New Recordset
   Debug.Print lsSQL
   lors.Open lsSQL, p_oAppDrivr.Connection, , , adCmdText
                     
   If lors.EOF Then GoTo endWithClear
   If lors.RecordCount = 1 Then
      p_oDetail(lnIndex) = lors(0)
      RaiseEvent DetailRetrieved(lnIndex)
      GoTo endProc
   End If
   
   lsSQL = p_oMod.KwikBrowse(p_oAppDrivr, lors, p_asLgrLstCol(lnIndex), p_asLgrLstTit(lnIndex), p_asLgrLstPic(lnIndex))
   
   If lsSQL = "" Then GoTo endProc
   lasRef = Split(lsSQL, "»")
   p_oDetail(lnIndex) = lasRef(0)
   RaiseEvent DetailRetrieved(lnIndex)
   
endProc:
   Set lors = Nothing
   Exit Sub
endWithClear:
   p_oDetail(lnIndex) = ""
   RaiseEvent DetailRetrieved(lnIndex)
   GoTo endProc
errProc:
   ShowError lsProcName & "( " & lnRow _
                       & ", " & lnIndex _
                       & ", " & lsValue _
                       & ", " & lbSearch & " )"
End Sub

Private Sub LoadDescription()
   For pnCtr = 1 To p_oMaster.Fields.Count - 1
      If p_asLstSQL(pnCtr) <> Empty Then
         If IsNull(p_oMaster(pnCtr)) = False Then
            getMasDesc pnCtr, p_oMaster(pnCtr)
         End If
      End If
   Next
End Sub

Private Function HasRights2Save() As Boolean
   Dim lsProcName As String
   Dim lnUserRght As Integer, lsUserName As String, lsUserIDxx As String
   Dim lasRights() As String, lnRights As Integer
   Dim lnSeek As Integer, lbSeek As Boolean
   Dim lasMenu() As String
   Dim loObject As Object
   
   lsProcName = "HasRights2Save"
   'On Error GoTo errProc
   
   If p_oAppDrivr.UserLevel = xeSysMaster Then
      HasRights2Save = True
      GoTo endProc
   End If
   
   If isPosted Then
      MsgBox "Modification of posted or cancelled transaction is not allowed!!!" & vbCrLf & vbCrLf & _
               "Request not granted!!!", vbCritical, "Warning"
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
      
      'kalyptus - 2018.05.28 11:57am
      'Add ability to handle array in our menus...
      lasMenu = Split(.MenuName, "»")
      If (UBound(lasMenu) = 1) Then
         Set loObject = .MDIMain.Controls(lasMenu(0)).Item(lasMenu(1))
      Else
         Set loObject = .MDIMain.Controls(lasMenu(0))
      End If
      
      'If .MDIMain.Controls(.MenuName).Tag = "" Then
      If loObject.Tag = "" Then
         ReDim lasRights(5) As String
   
         lasRights(1) = 231
         lasRights(2) = 230
         lasRights(3) = 228
         lasRights(4) = 192
         lasRights(5) = 230
      Else
         lasRights = Split(loObject.Tag, "»") 'mac 2022.11.09
         'lasRights = Split(.MDIMain.Controls(.MenuName).Tag, "»")
         If lasRights(6) <> Empty Then
            lasRights(1) = lasRights(6)
            lasRights(2) = lasRights(7)
            lasRights(3) = lasRights(8)
            lasRights(4) = lasRights(9)
            lasRights(5) = lasRights(10)
         End If
      End If
      
      lbSeek = False
      If pnEditMode = xeModeAddNew Then
         ' Add Rights
         lnRights = CInt(lasRights(1))
         If (lnRights And .UserLevel) = 0 Then
            lnSeek = MsgBox("User Doesn't Have the Right to Create New Transaction!!!" & _
                              vbCrLf & "Seek Authorized User's Approval?", _
                              vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
            If lnSeek = vbNo Then GoTo endProc
            lbSeek = True
         End If
      ElseIf pnEditMode = xeModeUpdate Then
         ' update Rights
         lnRights = CInt(lasRights(2))
         
         If (lnRights And .UserLevel) = 0 Then
            Select Case .UserLevel
            Case xeEncoder
               lbSeek = True
            Case xeSupervisor
               If DateDiff("d", p_oMaster("dModified"), p_oAppDrivr.ServerDate) > 2 Then lbSeek = True
            Case xeManager
               If DateDiff("d", p_oMaster("dModified"), p_oAppDrivr.ServerDate) > 7 Then lbSeek = True
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
      
      lnSeek = MsgBox("User Doesn't Have the Right to Update Transaction!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
      If lnSeek = vbNo Then GoTo endProc
      
      If p_oMod.GetApproval(p_oAppDrivr, lnUserRght, lsUserIDxx, lsUserName, lasRights(0)) = False Then GoTo endProc
      
      If pnEditMode = xeModeAddNew Then
         If (lnRights And lnUserRght) = 0 Then
            MsgBox "Approving User Doesn't Have the Right to Create New Transaction!!!" & _
                  "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
            GoTo endProc
         End If
      ElseIf pnEditMode = xeModeUpdate Then
         If (lnRights And lnUserRght) = 0 Then
            Select Case lnUserRght
            Case xeSupervisor
               If DateDiff("d", p_oMaster("dModified"), p_oAppDrivr.ServerDate) <= 2 Then lbSeek = False
            Case xeManager
               If DateDiff("d", p_oMaster("dModified"), p_oAppDrivr.ServerDate) <= 7 Then lbSeek = False
            Case xeSysAdmin
               If .UserLevel <> xeSysAdmin Then lbSeek = False
            End Select
            
            If lbSeek Then
               MsgBox "Approving User Doesn't Have the Right to Update Transaction!!!" & _
                     "Seek Other Users Assistance then Try Again", vbCritical, "Warning"
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

Private Function HasRights2Cancel() As Boolean
   Dim lsProcName As String
   Dim lnUserRght As Integer, lsUserName As String, lsUserIDxx As String
   Dim lasRights() As String, lnRights As Integer
   Dim lnSeek As Integer, lbSeek As Boolean
   Dim lasMenu() As String
   Dim loObject As Object
   
   lsProcName = "HasRights2Cancel"
   'On Error GoTo errProc
   
   If p_oAppDrivr.UserLevel = xeSysMaster Then
      HasRights2Cancel = True
      GoTo endProc
   End If
   
   If isPosted Then
      MsgBox "Cancellation of posted or cancelled transaction is not allowed!!!" & vbCrLf & vbCrLf & _
               "Request not granted!!!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   With p_oAppDrivr
      lasMenu = Split(.MenuName, "»")
      If (UBound(lasMenu) = 1) Then
         Set loObject = .MDIMain.Controls(lasMenu(0)).Item(lasMenu(1))
      Else
         Set loObject = .MDIMain.Controls(lasMenu(0))
      End If
      
      'If .MDIMain.Controls(.MenuName).Tag = "" Then
      If loObject.Tag = "" Then
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
      
      ' Cancel Rights
      lnRights = CInt(lasRights(3))
      If (lnRights And .UserLevel) > 0 Then
         HasRights2Cancel = True
         GoTo endProc
      End If
      
      lbSeek = False
      Select Case .UserLevel
      Case xeEncoder
         lbSeek = True
      Case xeSupervisor
         If DateDiff("d", p_oMaster("dModified"), .ServerDate) > 2 Then lbSeek = True
      Case xeManager
         If DateDiff("d", p_oMaster("dModified"), .ServerDate) > 7 Then lbSeek = True
      Case xeSysAdmin
         lbSeek = True
      Case xeAudit
         lbSeek = True
      End Select

      If lbSeek = False Then
         HasRights2Cancel = True
         GoTo endProc
      End If
   
      lnSeek = MsgBox("User Doesn't Have the Right to Cancel Transaction!!!" & _
                        vbCrLf & "Seek Authorized User's Approval?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, "Confirm")
      If lnSeek = vbNo Then GoTo endProc
      
      If p_oMod.GetApproval(p_oAppDrivr, lnUserRght, lsUserIDxx, lsUserName, lasRights(0)) = False Then GoTo endProc
   
      If (lnRights And lnUserRght) = 0 Then
         Select Case lnUserRght
         Case xeSupervisor
            If DateDiff("d", p_oMaster("dModified"), .ServerDate) <= 2 Then lbSeek = False
         Case xeManager
            If DateDiff("d", p_oMaster("dModified"), .ServerDate) <= 6 Then lbSeek = False
         Case xeSysAdmin
            If .UserLevel <> xeSysAdmin Then lbSeek = False
         End Select
         
         If lbSeek Then
            MsgBox "Approving User Doesn't Have the Right to Cancel Transaction!!!" & _
                  "Seek Other User's Assistance then Try Again", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
   End With
   HasRights2Cancel = True

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
   Dim lasMenu() As String
   Dim loObject As Object
   
   lsProcName = "HasRights2Delete"
   'On Error GoTo errProc
   
   If p_oAppDrivr.UserLevel = xeSysMaster Then
      HasRights2Delete = True
      GoTo endProc
   End If
   
   If isPosted Then
      MsgBox "Deletion of posted or cancelled transaction is not allowed!!!" & vbCrLf & vbCrLf & _
            "Request not granted!!!", vbCritical, "Warning"
      GoTo endProc
   End If

   With p_oAppDrivr
      lasMenu = Split(.MenuName, "»")
      If (UBound(lasMenu) = 1) Then
         Set loObject = .MDIMain.Controls(lasMenu(0)).Item(lasMenu(1))
      Else
         Set loObject = .MDIMain.Controls(lasMenu(0))
      End If
      
      If loObject.Tag = "" Then
      'If .MDIMain.Controls(.MenuName).Tag = "" Then
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
      
      MsgBox "Approving User Doesn't Have the Right to Delete Transaction!!!" & _
            "Seek Other Users Assistance then Try Again", vbCritical, "Warning"
   End With
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
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

Private Function AddNewTransaction(loMaster As Recordset, _
         loDetail As Recordset) As Boolean
   Dim lsProcName As String, lsDestinat As String
   Dim lsSQL As String, lsField As String
   Dim lnCtr As Long, lnField As Integer
   
   lsProcName = "AddNewTransaction"
   'On Error GoTo errProc
   
   With loMaster
      For lnCtr = 0 To .Fields.Count - IIf(LCase(.Fields(.Fields.Count - 1).Name) = "dtimestmp", 4, 3)
         lsField = .Fields(lnCtr).Name
         If lsField = "sDestinat" Then lsDestinat = p_oMaster(lsField)
         
         ' check first if field exist in table and in class
         For lnField = 0 To p_oMaster.Fields.Count - 1
            If LCase(p_oMaster(lnField).Name) = LCase(lsField) Then
               lsSQL = lsSQL & ", " & lsField & " = " & _
                        p_oMod.FieldParam(.Fields(lsField).Type, _
                        p_oMaster(lsField))
            End If
         Next
      Next
      
      If p_sDestinat <> "" Then lsDestinat = p_sDestinat
      
      lsSQL = "INSERT INTO " & p_sMaster & " SET" & _
                  Mid(lsSQL, 2) & _
                  ", sModified = " & _
                     p_oMod.strParm(p_oAppDrivr.UserID) & _
                  ", dModified = " & _
                     p_oMod.dateParm(p_oAppDrivr.ServerDate)
   End With
   
   lnCtr = p_oAppDrivr.Execute(lsSQL, p_sMaster, p_sBranchCd, lsDestinat)
   If lnCtr = 0 Then
      MsgBox "Unable to Add Transaction Master!", vbCritical, "Warning"
      GoTo endProc
   End If
   
   With loDetail
      p_oDetail.MoveFirst
      
      Do While p_oDetail.EOF = False
         lsSQL = ""
         lsField = .Fields(0).Name
         If p_oDetail(lsField) = Empty Then Exit Do
         
         For lnCtr = 0 To .Fields.Count - 2
            lsField = .Fields(lnCtr).Name
            
            For lnField = 0 To p_oDetail.Fields.Count - 1
               If LCase(p_oDetail(lnField).Name) = LCase(lsField) Then
                  lsSQL = lsSQL & ", " & lsField & " = " & _
                           p_oMod.FieldParam(.Fields(lsField).Type, _
                           p_oDetail(lsField))
               End If
            Next
         Next
         
         lsSQL = "INSERT INTO " & p_sDetail & " SET" & _
                     Mid(lsSQL, 2) & _
                     ", dModified = " & _
                        p_oMod.dateParm(p_oAppDrivr.ServerDate)
      
         lnCtr = p_oAppDrivr.Execute(lsSQL, p_sDetail, p_sBranchCd, lsDestinat)
         If lnCtr = 0 Then
            MsgBox "Unable to Add Transaction Master!", vbCritical, "Warning"
            GoTo endProc
         End If
         p_oDetail.MoveNext
      Loop
   End With
      
   AddNewTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Private Function UpdateOldTransaction(loMaster As Recordset, _
         loDetail As Recordset) As Boolean
   Dim lsProcName As String
   Dim lsSQL As String, lsField As String
   Dim lsDestinat As String
   Dim lbNewDetail As Boolean, lbModify As Boolean
   Dim lnCol As Integer, lnCtr As Long
   Dim lnField As Integer
   
   lsProcName = "UpdateOldTransaction"
   'On Error GoTo errProc
   
   With loMaster
      If .EOF Then
         MsgBox "Invalid Record Detected!", vbCritical, "Warning"
         GoTo endProc
      End If
         
      For lnCtr = 0 To .Fields.Count - 3
         lsField = .Fields(lnCtr).Name
         ' just hardcode this field name
         If lsField = "sDestinat" Then lsDestinat = p_oMaster(lsField)
         
         For lnField = 0 To p_oMaster.Fields.Count - 1
            If LCase(p_oMaster(lnField).Name) = LCase(lsField) Then
               If IsNull(.Fields(lsField)) Or _
                     .Fields(lsField) <> IFNull(p_oMaster(lsField)) Then
                  If IsNull(p_oMaster(lsField)) = False Then
                     lsSQL = lsSQL & ", " & lsField & " = " & _
                                 p_oMod.FieldParam(.Fields(lsField).Type, _
                                 p_oMaster(lsField))
                  End If
               End If
            End If
         Next
      Next
      
      If p_sDestinat <> "" Then lsDestinat = p_sDestinat
      
      If lsSQL <> Empty Then
         lsSQL = "UPDATE " & p_sMaster & " SET" & _
                     Mid(lsSQL, 2) & _
                     ", dModified = " & _
                        p_oMod.dateParm(p_oAppDrivr.ServerDate) & _
                     " WHERE "
                        
         For lnCtr = 0 To UBound(p_asMasRefFld)
            lsField = p_asMasRefFld(lnCtr)
            lnCol = InStr(lsField, ".")
            If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
            
            lsSQL = lsSQL & " " & lsField & " = " & _
                     p_oMod.strParm((.Fields(lsField))) & _
                     " AND "
         Next
         lsSQL = Left(lsSQL, Len(Trim(lsSQL)) - 3)
         
         lnCtr = p_oAppDrivr.Execute(lsSQL, p_sMaster, p_sBranchCd, lsDestinat)
         If lnCtr = 0 Then
            MsgBox "Unable to Update Transaction Master!", vbCritical, "Warning"
            GoTo endProc
         End If
      End If
   End With
   
   With loDetail
      If .RecordCount > 0 Then .MoveFirst
      p_oDetail.MoveFirst
      
      Do While p_oDetail.EOF = False
         lsSQL = ""
         lsField = .Fields(0).Name
         If p_oDetail.Fields(lsField) = Empty Then Exit Do
         
         If .EOF Then lbNewDetail = True
         
         For lnCtr = 0 To .Fields.Count - 2
            lsField = .Fields(lnCtr).Name
            
            For lnField = 0 To p_oDetail.Fields.Count - 1
               If LCase(p_oDetail(lnField).Name) = LCase(lsField) Then
                  lbModify = lbNewDetail
                  If lbNewDetail = False Then
                     lbModify = p_oMod.IFNull(.Fields(lsField)) <> p_oMod.IFNull(p_oDetail(lsField))
                  End If
                  
                  If lbModify Then
                     lsSQL = lsSQL & ", " & lsField & " = " & _
                              p_oMod.FieldParam(p_oDetail(lsField).Type, _
                              p_oDetail(lsField))
                  End If
               End If
            Next
         Next
         
         If lsSQL <> Empty Then
            If lbNewDetail Then
               lsSQL = "INSERT INTO " & p_sDetail & " SET" & _
                           Mid(lsSQL, 2) & _
                           ", dModified = " & _
                              p_oMod.dateParm(p_oAppDrivr.ServerDate)
            Else
               lsSQL = "UPDATE " & p_sDetail & " SET" & _
                           Mid(lsSQL, 2) & _
                           ", dModified = " & _
                              p_oMod.dateParm(p_oAppDrivr.ServerDate) & _
                        " WHERE "
               
               For lnCtr = 0 To UBound(p_asLgrRefFld)
                  lsField = p_asLgrRefFld(lnCtr)
                  lnCol = InStr(lsField, ".")
                  If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
                   
                  lsSQL = lsSQL & lsField & " = " & _
                           p_oMod.strParm((.Fields(lsField))) & _
                           " AND "
               Next
               ' always assume that detail has entry no field
               lsSQL = lsSQL & " nEntryNox = " & p_oMod.strParm(.Fields("nEntryNox"))
            End If
                           
            lnCtr = p_oAppDrivr.Execute(lsSQL, p_sMaster, p_sBranchCd, lsDestinat)
            If lnCtr = 0 Then
               MsgBox "Unable to Update Transaction Detail!", vbCritical, "Warning"
               GoTo endProc
            End If
         End If
         
         p_oDetail.MoveNext
         If .EOF = False Then .MoveNext
      Loop
      
      Do While .EOF = False
         lsSQL = "DELETE FROM " & p_sDetail & " WHERE "
         
         For lnCtr = 0 To UBound(p_asLgrRefFld)
            lsField = p_asLgrRefFld(lnCtr)
            lnCol = InStr(lsField, ".")
            If lnCol > 0 Then lsField = Mid(lsField, lnCol + 1)
             
            lsSQL = lsSQL & lsField & " = " & _
                     p_oMod.strParm((.Fields(lsField))) & _
                     " AND "
         Next
         
         ' always assume that detail has entry no field
         lsSQL = lsSQL & " nEntryNox = " & p_oMod.strParm(.Fields("nEntryNox"))
         
         lnCtr = p_oAppDrivr.Execute(lsSQL, p_sMaster, p_sBranchCd, lsDestinat)
         If lnCtr = 0 Then
            MsgBox "Unable to Delete Transaction Detail!", vbCritical, "Warning"
            GoTo endProc
         End If
         
         .MoveNext
      Loop
   End With
   
   UpdateOldTransaction = True
   
endProc:
   Exit Function
errProc:
   ShowError lsProcName & "( " & " )"
End Function

Function IFNull(ByVal Value As Variant, Optional ByVal Repl As Variant = "")
   IFNull = IIf(IsNull(Value), Repl, IIf(IsMissing(Value), Repl, Value))
End Function

