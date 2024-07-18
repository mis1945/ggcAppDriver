VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodeGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
' Copyright 2014 and beyond
' All Rights Reserved
'
'  Code Generator Object
'     Handles the generation of approval code for all modules requiring approval code...
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº
' €  All  rights reserved. No part of this  software  €€  This Software is Owned by        €
' €  may be reproduced or transmitted in any form or  €€                                   €
' €  by   any   means,  electronic   or  mechanical,  €€    GUANZON MERCHANDISING CORP.    €
' €  including recording, or by information  storage  €€     Guanzon Bldg. Perez Blvd.     €
' €  and  retrieval  systems, without  prior written  €€           Dagupan City            €
' €  from the author.                                €€  Tel No. 522-1085 ; 522-9275      €
' ºººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººººº

Option Explicit
Private p_oSCA_Const As New clsCodeGenConst

Private Type xObject
    System As String   'System Approval Requested
    Branch As String   'Requesting Branch
    IssuedBy As String 'Issuing Department/Person
    Date As String
    Misc As String
End Type

Private poRawxxx As xObject
Private poResult As xObject
Private pbEncode As Boolean
Private pbReferx As Boolean

Property Get Result() As String
   If pbEncode Then
      If poRawxxx.Branch <> "" Then
         'Result = MMMIBBSSDD
         
         Debug.Print poResult.Misc
         Debug.Print poResult.IssuedBy
         Debug.Print poResult.Branch
         Debug.Print poResult.System
         Debug.Print poResult.Date
         
         Result = poResult.Misc & poResult.IssuedBy & poResult.Branch & poResult.System & poResult.Date
      Else
         'Result = MMMISSDDDD
         Result = poResult.Misc & poResult.IssuedBy & poResult.System & poResult.Date
      End If
   End If
End Property

Property Let Branch(ByVal fsBranchCD As String)
   poRawxxx.Branch = fsBranchCD
End Property

Property Let DateRequested(ByVal fdRequestxx As Date)
   poRawxxx.Date = fdRequestxx
End Property

Property Let IssuedBy(ByVal fcIssuedBy As String)
   poRawxxx.IssuedBy = fcIssuedBy
End Property

Property Let System(ByVal fsSystem As String)
   poRawxxx.System = fsSystem
End Property

Property Let MiscInfo(ByVal fsValue As String)
   poRawxxx.Misc = fsValue
End Property

Property Let IsByRef(ByVal fbValue As Boolean)
   pbReferx = fbValue
End Property

Function Encode() As Boolean
   pbEncode = False
   
   'kalyptus - 2015.12.23 11:23am
   'remove verification of branch here...
'   'Verify branch
'   If poRawxxx.Branch = "" Then
'      MsgBox "Invalid branch detected!", vbCritical + vbOKOnly, "Verification"
'      Exit Function
'   End If
   
   'Verify System Approval Requested
   If poRawxxx.System = "" Then
      MsgBox "Invalid System Approval Requested detected!", vbCritical + vbOKOnly, "Verification"
      Exit Function
   End If
   
   'Verify date requested
   If poRawxxx.Date = "" Then
      MsgBox "Invalid Date Requested detected!", vbCritical + vbOKOnly, "Verification"
      Exit Function
   End If
   
   'Verify issuing Department/Person
   If poRawxxx.IssuedBy = "" Then
      MsgBox "Invalid Issuing Department/Person detected!", vbCritical + vbOKOnly, "Verification"
      Exit Function
   End If
   
   Select Case poRawxxx.System
   Case p_oSCA_Const.pxeManualLog
      'Misc should be the binary equivalent of the periods approved...
      If Not IsNumeric(poRawxxx.Misc) Then
         MsgBox "Invalid Reference Number detected!", vbCritical + vbOKOnly, "Verification"
         Exit Function
      End If
         
      poResult.Misc = Random(0, 9) & PadLeft(Hex(poRawxxx.Misc), 2, "0")
   Case p_oSCA_Const.pxeDay2Day
      'Misc should be the time the request was issued...
      poResult.Misc = Chr(Random(65, 90)) & PadLeft(Hex(Val(poRawxxx.Misc) + 70), 2, "0")
'   Case Else
'      If pbReferx Then
'         'Misc should be the reference number of the transaction approved...
'         If Not IsNumeric(poRawxxx.Misc) Then
'            MsgBox "Invalid Reference Number detected!", vbCritical + vbOKOnly, "Verification"
'            Exit Function
'         End If
'         poResult.Misc = PadLeft(Hex(TotalStr(poRawxxx.Misc)), 3, "0")
'      Else
'         ' FIRST 30 characters of fullname with the following format:
'         '   LASTNAME, FIRSTNAME(SFX) MIDDNAME
'         poResult.Misc = PadLeft(Hex(TotalStr(LCase(Left(poRawxxx.Misc, 30)))), 3, "0")
'      End If
'   End Select
   
   Case p_oSCA_Const.pxeOfficeRebate, _
        p_oSCA_Const.pxeFieldRebate, _
        p_oSCA_Const.pxeMCDiscount, _
        p_oSCA_Const.pxePartsDiscount, _
        p_oSCA_Const.pxeSPPurcDelivery, _
        p_oSCA_Const.pxeIssueORNotPR, _
        p_oSCA_Const.pxeIssueORNotSI, _
        p_oSCA_Const.pxeMCIssuance, _
        p_oSCA_Const.pxeMPDiscount, _
        p_oSCA_Const.pxeJobOrderWOGCard, _
        p_oSCA_Const.pxeMCDownpayment, _
        p_oSCA_Const.pxeMCTransfer, _
        p_oSCA_Const.pxeARFirstPayDate, _
        p_oSCA_Const.pxeMCSalesWoPNP, _
        p_oSCA_Const.pxeSalesCommision, _
        p_oSCA_Const.pxeAdvanceBackDate, _
        p_oSCA_Const.pxeBranchPOReturn, _
        p_oSCA_Const.pxeBranch2BranchTransfer, _
        p_oSCA_Const.pxeBranch2WarehouseTransfer, _
        p_oSCA_Const.pxeWarehouse2WarehouseTransfer, _
        p_oSCA_Const.pxeJobOrderRelease
      
      'kalyptus - 2016.11.02 12:01pm
      'Included condition for pxeJobOrderWOGCard
      If p_oSCA_Const.pxeJobOrderWOGCard <> poRawxxx.System And _
         p_oSCA_Const.pxeBranchPOReturn <> poRawxxx.System And _
         p_oSCA_Const.pxeBranch2BranchTransfer <> poRawxxx.System And _
         p_oSCA_Const.pxeBranch2WarehouseTransfer <> poRawxxx.System And _
         p_oSCA_Const.pxeWarehouse2WarehouseTransfer <> poRawxxx.System And _
         p_oSCA_Const.pxeJobOrderRelease <> poRawxxx.System Then
         'Misc should be the reference number of the transaction approved...
         If Not IsNumeric(poRawxxx.Misc) Then
            MsgBox "Invalid Reference Number detected!", vbCritical + vbOKOnly, "Verification"
            Exit Function
         End If
      Else
         poRawxxx.Misc = Replace(poRawxxx.Misc, "-", "")
      End If
      
      poResult.Misc = PadLeft(Hex(TotalStr(poRawxxx.Misc)), 3, "0")
   Case p_oSCA_Const.pxeForgot2Log, _
        p_oSCA_Const.pxeBusinessTrip, _
        p_oSCA_Const.pxeBusinessTripWLog, _
        p_oSCA_Const.pxeLeave, _
        p_oSCA_Const.pxeOvertime, _
        p_oSCA_Const.pxeShift, _
        p_oSCA_Const.pxeDayOff, _
        p_oSCA_Const.pxeTardiness, _
        p_oSCA_Const.pxeUnderTime, _
        p_oSCA_Const.pxeCreditInvestigation, _
        p_oSCA_Const.pxeCreditApplication, _
        p_oSCA_Const.pxeCashBalance, _
        p_oSCA_Const.pxeMCClusteringDelivery, _
        p_oSCA_Const.pxeFSEPActivation, _
        p_oSCA_Const.pxeFSEXActivation, _
        p_oSCA_Const.pxeHIMApplication
      
      'FIRST 30 characters of fullname with the following format:
      '   LASTNAME, FIRSTNAME(SFX) MIDDNAME
      poResult.Misc = PadLeft(Hex(TotalStr(LCase(Left(poRawxxx.Misc, 30)))), 3, "0")
   Case p_oSCA_Const.pxeAdditional, p_oSCA_Const.pxeBiyahingFiesta, p_oSCA_Const.pxeTeleMktg, p_oSCA_Const.pxePreApproved
      poResult.Misc = PadLeft(Hex(TotalStr(LCase(Left(poRawxxx.Misc, 30)))), 3, "0")
   Case Else
      MsgBox "Invalid System Approval Request detected!", vbCritical + vbOKOnly, "Verification"
      Exit Function
   End Select
   
   If poRawxxx.Branch <> "" Then
      poResult.Branch = PadLeft(Hex(TotalStr(Mid(poRawxxx.Branch, 2))), 2, "0")
      poResult.Date = PadLeft(Hex(Month(poRawxxx.Date) + Day(poRawxxx.Date) + Int(Format(poRawxxx.Date, "YY"))), 2, "0")
   Else
      Debug.Print PadLeft(Day(poRawxxx.Date), 2, "0")
      Debug.Print PadLeft(Month(poRawxxx.Date), 2, "0")
      Debug.Print Right(Year(poRawxxx.Date), 2)
      
      poResult.Date = Hex(PadLeft(Day(poRawxxx.Date), 2, "0") & PadLeft(Month(poRawxxx.Date), 2, "0") & Right(Year(poRawxxx.Date), 2))
      'poResult.Date = Hex(Format(poRawxxx.Date, "DDMMY"))
   End If
   
   poResult.System = PadLeft(Hex(TotalStr(poRawxxx.System)), 2, "0")
   poResult.IssuedBy = poRawxxx.IssuedBy
   
   pbEncode = True
   Encode = True
End Function

Function Equal(fsCode1 As String, fsCode2 As String) As Integer
   Equal = -100
   
   'Length is not equal to 10
   If Len(fsCode1) <> 10 Then Exit Function
   If Len(fsCode2) <> 10 Then Exit Function
   
   'Convert to uppercase the code to be checked
   fsCode1 = UCase(fsCode1)
   fsCode2 = UCase(fsCode2)
   
   'Requesting branch is different from the given code
   If Not Mid(fsCode1, 5, 2) = Mid(fsCode2, 5, 2) Then Exit Function
   
   'System approval request is different from the given code
   If Not Mid(fsCode1, 7, 2) = Mid(fsCode2, 7, 2) Then Exit Function
   
   'Date requested is different from the given code
   If Not Mid(fsCode1, 9, 2) = Mid(fsCode2, 9, 2) Then Exit Function
      
   'Issuing Department/Person is different from the given code
   If Not Mid(fsCode1, 4, 1) = Mid(fsCode2, 4, 1) Then Exit Function
      
   Select Case Mid(fsCode1, 7, 2) 'System Approval Request
   Case PadLeft(Hex(TotalStr(p_oSCA_Const.pxeDay2Day)), 2, "0")
      'Misc Info is different from the given code
      'New Issued - Old Issued => If <=0 Then INVALID SINCE we need a new code
      Equal = Val(Mid(fsCode2, 2, 2)) - Val(Mid(fsCode1, 2, 2))
      Exit Function
   Case PadLeft(Hex(TotalStr(p_oSCA_Const.pxeManualLog)), 2, "0")
      'Misc Info is different from the given code
      If Not Mid(fsCode1, 2, 2) = Mid(fsCode2, 2, 2) Then Exit Function
   Case Else
      'Misc Info is different from the given code
      If Not Mid(fsCode1, 1, 3) = Mid(fsCode2, 1, 3) Then Exit Function
   End Select
   
   Equal = 0
End Function

'fsCode2 is the Approval Code given by the authorized personel/department
Function Equalx(fsCode1 As String, fsCode2 As String) As Integer
   Dim ldDate1 As Date
   Dim ldDate2 As Date
   Dim lsDatex As String
   
   Equalx = -100
   
   'Length is not equal to 10
   If Len(fsCode1) <> Len(fsCode1) Then Exit Function
   If Len(fsCode2) < 10 Then Exit Function
   
   'Convert to uppercase the code to be checked
   fsCode1 = UCase(fsCode1)
   fsCode2 = UCase(fsCode2)
   
   'Issuing Department/Person is different from the given code
   If Not Mid(fsCode1, 4, 1) = Mid(fsCode2, 4, 1) Then Exit Function
   
   'System approval request is different from the given code
   If Not Mid(fsCode1, 5, 2) = Mid(fsCode2, 5, 2) Then Exit Function
   
   'Misc Information/Name
   If Not Mid(fsCode1, 1, 3) = Mid(fsCode2, 1, 3) Then Exit Function
   
   'Check date
   'Date has this format: DDMMYY
   lsDatex = PadLeft(CLng("&H" & Mid(fsCode2, 7)), 6, "0")
   
   ldDate1 = Mid(lsDatex, 3, 2) & "/" & Mid(lsDatex, 1, 2) & "/" & Mid(lsDatex, 5, 2)
               
   lsDatex = PadLeft(CLng("&H" & Mid(fsCode1, 7)), 6, "0")
   Debug.Print Mid(lsDatex, 3, 2)
   Debug.Print Mid(lsDatex, 1, 2)
   Debug.Print Mid(lsDatex, 5, 2)
   
   ldDate2 = Mid(lsDatex, 3, 2) & "/" & Mid(lsDatex, 1, 2) & "/" & Mid(lsDatex, 5, 2)
                  
   Select Case Mid(fsCode1, 5, 2)
   Case PadLeft(Hex(TotalStr(p_oSCA_Const.pxeTeleMktg)), 2, "0"), PadLeft(Hex(TotalStr(p_oSCA_Const.pxePreApproved)), 2, "0")
      If ldDate1 >= ldDate2 And ldDate1 <= DateAdd("D", 60, ldDate2) Then
         Equalx = 0
      End If
   Case PadLeft(Hex(TotalStr(p_oSCA_Const.pxeBiyahingFiesta)), 2, "0")
      'Biyaheng Fiesta should be dated from the start of Biyaheng Fiesta
      If ldDate1 >= ldDate2 And ldDate1 <= DateAdd("D", 3, ldDate2) Then
         Equalx = 0
      End If
   Case PadLeft(Hex(TotalStr(p_oSCA_Const.pxeAdditional)), 2, "0")
      If ldDate1 = ldDate2 Then
         Equalx = 0
      End If
   End Select
End Function

Private Function TotalStr(ByVal fsStr As String) As Integer
   Dim lnCtr As Integer
   Dim lnTotal As Integer
   fsStr = Replace(fsStr, " ", "")
   fsStr = Replace(fsStr, ",", "")
      
   For lnCtr = 1 To Len(fsStr)
      lnTotal = lnTotal + Asc(Mid(fsStr, lnCtr, 1))
   Next
   
   TotalStr = lnTotal
End Function

Private Function Bin2Dec(Num As String) As Long
   Dim n As Integer
   Dim x As Integer
   Dim a As Integer
   n = Len(Num) - 1
   a = n
   Do While n > -1
      x = Mid(Num, ((a + 1) - n), 1)
      Bin2Dec = IIf((x = "1"), Bin2Dec + (2 ^ (n)), Bin2Dec)
      n = n - 1
   Loop
End Function

Private Function Dec2Bin(ByVal Num As Long) As String
   Select Case Num
      Case Is > 1: Dec2Bin = Dec2Bin(Num \ 2) & (Num Mod 2)
      Case 0, 1:   Dec2Bin = Num
      Case Else:   Dec2Bin = -Dec2Bin(-Num)
   End Select
End Function

Private Function PadLeft(ByVal fsStr As String, ByVal fnLen As Integer, ByVal fsPad As String) As String
   If fnLen <= Len(fsStr) Then
      PadLeft = fsStr
   Else
      PadLeft = String(fnLen - Len(fsStr), Left(fsPad, 1)) & fsStr
   End If
End Function

Private Function Random(ByVal from As Long, ByVal upto As Long) As Integer
   Randomize
   Random = Int((upto - from + 1) * Rnd + from)
End Function

