VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Receipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type CheckInfo
   CheckNox As String
   AcctNoxx As String
   BankName As String
   CheckDte As Date
   Amountxx As Double
End Type

Private p_oAppDrivr As clsAppDriver
Private p_oForm As frmReceipt
Private p_oMod As New clsMainModules

Private p_bCancelxx As Boolean
Private p_dTransact As Date
Private p_sORNoxxxx As String
Private p_sReceived As String
Private p_sAddressx As String
Private p_nTranTotl As Double
Private p_nAmtPaidx As Double
Private p_nCashAmtx As Double
Private p_nCheckAmt As Double
Private p_sRemarksx As String
Private p_bEmptyORx As Boolean
Private p_axChecks(4) As CheckInfo

Private bLoaded As Boolean

Dim pnCtr As Integer

Property Set AppDriver(Value As clsAppDriver)
   Set p_oAppDrivr = Value
End Property

Property Get AllowEmptyOR() As Boolean
   AllowEmptyOR = p_bEmptyORx
End Property

Property Let AllowEmptyOR(ByVal Value As Boolean)
   p_bEmptyORx = Value
End Property

Property Get Checks(ByVal Row As Integer, ByVal Index As Variant) As Variant
   If Row > 4 Then Exit Property

   Select Case Index
      Case 0, "CheckNox"
         Checks = p_axChecks(Row).CheckNox
      Case 1, "sAcctNoxx"
         Checks = p_axChecks(Row).AcctNoxx
      Case 2, "sBankName"
         Checks = p_axChecks(Row).BankName
      Case 3, "dCheckDte"
         Checks = p_axChecks(Row).CheckDte
      Case 4, "nAmountxx"
         Checks = p_axChecks(Row).Amountxx
   End Select
End Property

Property Let Checks(ByVal Row As Integer, ByVal Index As Variant, ByVal Value As Variant)
   If Row > 4 Then Exit Property

   Select Case Index
      Case 0, "CheckNox"
         p_axChecks(Row).CheckNox = Value
      Case 1, "sAcctNoxx"
         p_axChecks(Row).AcctNoxx = Value
      Case 2, "sBankName"
         p_axChecks(Row).BankName = Value
      Case 3, "dCheckDte"
         p_axChecks(Row).CheckDte = Value
      Case 4, "nAmountxx"
         p_axChecks(Row).Amountxx = Value
   End Select
End Property

Property Get TranTotal() As Double
   TranTotal = p_nTranTotl
End Property

Property Let TranTotal(ByVal Value As Double)
   p_nTranTotl = Value
End Property

Property Get AmountPaid() As Double
   AmountPaid = p_nAmtPaidx
End Property

Property Let AmountPaid(ByVal Value As Double)
   p_nAmtPaidx = Value
End Property

Property Get CashAmount() As Double
   CashAmount = p_nCashAmtx
End Property

Property Let CashAmount(ByVal Value As Double)
   p_nCashAmtx = Value
End Property

Property Get CheckAmount() As Double
   CheckAmount = p_nCheckAmt
End Property

Property Let CheckAmount(ByVal Value As Double)
   p_nCheckAmt = Value
End Property

Property Get Remarks() As String
   Remarks = p_sRemarksx
End Property

Property Let Remarks(ByVal Value As String)
   p_sRemarksx = Value
End Property

Property Let Address(ByVal Value As String)
   p_sAddressx = Value
End Property

Property Let TranDate(ByVal Value As Date)
   p_dTransact = Value
End Property

Property Let ReceiveFrom(ByVal Value As String)
   p_sReceived = Value
End Property

Property Get Cancelled() As Boolean
   Cancelled = p_bCancelxx
End Property

Property Get ORNo() As String
   ORNo = p_sORNoxxxx
End Property

Property Let ORNo(ByVal Value As String)
   p_sORNoxxxx = Value
End Property

Sub InitReceipt()
   If p_oAppDrivr Is Nothing Then Exit Sub
   
   Set p_oForm.AppDriver = p_oAppDrivr
   p_bCancelxx = False
   p_dTransact = Date
   p_sORNoxxxx = ""
   p_sReceived = ""
   p_sAddressx = ""
   p_nTranTotl = 0#
   p_nCashAmtx = 0#
   p_nCheckAmt = 0#
   p_sRemarksx = ""
   
   If p_bEmptyORx = False Then
      p_sORNoxxxx = p_oForm.GetNextOR
   Else
      p_sORNoxxxx = ""
   End If
   
   InitChecks
End Sub

Sub ShowReceipt()
   If p_oAppDrivr Is Nothing Then Exit Sub
   
   Load p_oForm
   With p_oForm
      .lblHead(0) = p_oAppDrivr.ClientName
      .lblHead(1) = Trim(p_oAppDrivr.Address) & ", " & Trim(p_oAppDrivr.TownCity) & ", " & _
            Trim(p_oAppDrivr.Province) & " " & Trim(p_oAppDrivr.ZippCode) & ", "
      .lblHead(2) = "Tel No: " & p_oAppDrivr.TelNo
      
      .AllowEmptyOR = p_bEmptyORx
      .txtField(0) = p_sORNoxxxx
      .txtField(1) = Format(p_dTransact, "MMMM DD, YYYY")
      .txtField(2) = p_sReceived
      .txtField(3) = p_sAddressx
      .txtField(4) = p_sRemarksx
      .txtField(5) = Format(p_nCashAmtx, "#,##0.00")
      .AmountPaid = Format(p_nAmtPaidx, "#,##0.00")
      .txtField(6) = Format(p_nCheckAmt, "#,##0.00")
      .lblPayment = Format(p_nTranTotl, "#,##0.00")
      .lblChange = Format(p_nAmtPaidx - p_nTranTotl, "#,##0.00")
      
      For pnCtr = 0 To 4
         .GridEditor1.TextMatrix(pnCtr + 1, 1) = p_axChecks(pnCtr).CheckNox
         .GridEditor1.TextMatrix(pnCtr + 1, 2) = p_axChecks(pnCtr).AcctNoxx
         .GridEditor1.TextMatrix(pnCtr + 1, 3) = p_axChecks(pnCtr).BankName
         .GridEditor1.TextMatrix(pnCtr + 1, 5) = p_axChecks(pnCtr).Amountxx
         .GridEditor1.TextMatrix(pnCtr + 1, 4) = p_axChecks(pnCtr).CheckDte
      Next
      
      .Show 1
      
      p_bCancelxx = .Cancelled
      
      p_sORNoxxxx = .txtField(0)
      p_dTransact = CDate(.txtField(1))
      p_sReceived = .txtField(2)
      p_sAddressx = .txtField(3)
      p_sRemarksx = .txtField(4)
      p_nCashAmtx = CDbl(.txtField(5))
      p_nCheckAmt = CDbl(.txtField(6))

      For pnCtr = 0 To 4
         p_axChecks(pnCtr).CheckNox = .GridEditor1.TextMatrix(pnCtr + 1, 1)
         p_axChecks(pnCtr).AcctNoxx = .GridEditor1.TextMatrix(pnCtr + 1, 2)
         p_axChecks(pnCtr).BankName = .GridEditor1.TextMatrix(pnCtr + 1, 3)
         p_axChecks(pnCtr).CheckDte = CDate(.GridEditor1.TextMatrix(pnCtr + 1, 4))
         p_axChecks(pnCtr).Amountxx = CDbl(.GridEditor1.TextMatrix(pnCtr + 1, 5))
      Next
   End With
   Unload p_oForm
End Sub

Private Sub InitChecks()
   For pnCtr = 0 To 4
      p_axChecks(pnCtr).CheckNox = ""
      p_axChecks(pnCtr).AcctNoxx = ""
      p_axChecks(pnCtr).BankName = ""
      p_axChecks(pnCtr).CheckDte = p_dTransact
      p_axChecks(pnCtr).Amountxx = 0#
   Next
End Sub

Private Sub Class_Initialize()
   Set p_oForm = New frmReceipt
End Sub

Private Sub Class_Terminate()
   Set p_oForm = Nothing
   Set p_oMod = Nothing
End Sub


