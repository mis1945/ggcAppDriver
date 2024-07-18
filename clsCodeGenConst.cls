VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodeGenConst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Auto-Counter
Public Property Get pxeDay2Day() As String
   pxeDay2Day = "DT"
End Property

'Period
Public Property Get pxeManualLog() As String
   pxeManualLog = "ML"
End Property

'LastName, FrstName MiddName
Public Property Get pxeForgot2Log() As String
   pxeForgot2Log = "FL"
End Property
Public Property Get pxeBusinessTrip() As String
   pxeBusinessTrip = "OB"
End Property
Public Property Get pxeBusinessTripWLog() As String
   pxeBusinessTripWLog = "OL"
End Property
Public Property Get pxeLeave() As String
   pxeLeave = "LV"
End Property
Public Property Get pxeCompensationOff() As String
   pxeCompensationOff = "CO"
End Property
Public Property Get pxeOvertime() As String
   pxeOvertime = "OT"
End Property
Public Property Get pxeShift() As String
   pxeShift = "SH"
End Property
Public Property Get pxeDayOff() As String
   pxeDayOff = "DO"
End Property
Public Property Get pxeTardiness() As String
   pxeTardiness = "TD"
End Property
Public Property Get pxeUnderTime() As String
   pxeUnderTime = "UD"
End Property

Public Property Get pxeCreditInvestigation() As String
   pxeCreditInvestigation = "CI"
End Property
Public Property Get pxeCreditApplication() As String
   pxeCreditApplication = "CA"
End Property
Public Property Get pxeWholeSaleDiscount() As String
   pxeWholeSaleDiscount = "WD"
End Property
Public Property Get pxeCashBalance() As String
   pxeCashBalance = "CB"
End Property

'Reference No
Public Property Get pxeOfficeRebate() As String
   pxeOfficeRebate = "R1"
End Property
Public Property Get pxeFieldRebate() As String
   pxeFieldRebate = "R2"
End Property
Public Property Get pxePartsDiscount() As String
   pxePartsDiscount = "SI"
End Property
Public Property Get pxeMCDiscount() As String
   pxeMCDiscount = "DR"
End Property
Public Property Get pxeSPPurcDelivery() As String
   pxeSPPurcDelivery = "PD"
End Property
Public Property Get pxeIssueORNotPR() As String
   pxeIssueORNotPR = "OR"
End Property
Public Property Get pxeIssueORNotSI() As String
   pxeIssueORNotSI = "OX"
End Property

Public Property Get pxeAdditional() As String
   pxeAdditional = "RS"
End Property
Public Property Get pxeBiyahingFiesta() As String
   pxeBiyahingFiesta = "BF"
End Property
Public Property Get pxeTeleMktg() As String
   pxeTeleMktg = "TM"
End Property

Public Property Get pxeMCIssuance() As String
   pxeMCIssuance = "MI"
End Property
Public Property Get pxeMCClusteringDelivery() As String
   pxeMCClusteringDelivery = "CD"
End Property

Public Property Get pxeFSEPActivation() As String
   pxeFSEPActivation = "FA"
End Property
Public Property Get pxeFSEXActivation() As String
   pxeFSEXActivation = "FX"
End Property
Public Property Get pxeMPDiscount() As String
   pxeMPDiscount = "MD"
End Property

Public Property Get pxePreApproved() As String
   pxePreApproved = "PA"
End Property

Public Property Get pxeJobOrderWOGCard() As String
   pxeJobOrderWOGCard = "JG"
End Property

'iMac 2017.05.23
Public Property Get pxeMCDownpayment() As String
   pxeMCDownpayment = "DP"
End Property
'iMac 2019.02.13
Public Property Get pxeMCTransfer() As String
   pxeMCTransfer = "MT"
End Property

'Jeff 2019.08.30
Public Property Get pxeARFirstPayDate() As String
   pxeARFirstPayDate = "FP"
End Property

'Jeff 2019.09.17
Public Property Get pxeHIMApplication() As String
   pxeHIMApplication = "HM"
End Property

'mac 2020.06.30
Public Property Get pxeMCSalesWoPNP() As String
   pxeMCSalesWoPNP = "VC"
End Property

'mac 2020.08.24
Public Property Get pxeSalesCommision() As String
   pxeSalesCommision = "SC"
End Property

'mac 2021.09.21
Public Property Get pxeAdvanceBackDate() As String
   pxeAdvanceBackDate = "AB"
End Property

Public Property Get pxeBranchPOReturn() As String
   pxeBranchPOReturn = "Rb"
End Property

Public Property Get pxeWarehousePOReturn() As String
   pxeWarehousePOReturn = "Rw"
End Property

Public Property Get pxeBranch2BranchTransfer() As String
   pxeBranch2BranchTransfer = "BB"
End Property

Public Property Get pxeBranch2WarehouseTransfer() As String
   pxeBranch2WarehouseTransfer = "BW"
End Property

Public Property Get pxeWarehouse2WarehouseTransfer() As String
   pxeWarehouse2WarehouseTransfer = "WW"
End Property

Public Property Get pxeJobOrderRelease() As String
   pxeJobOrderRelease = "JR"
End Property
