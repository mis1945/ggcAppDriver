VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSpeedometer"
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
' ProgressBar v1.0
'     An object for displaying a progress bar.
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
'  XerSys [ 10/02/2004 11:20 pm ]
'     Start creating this object.
'€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€
Option Explicit

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST = -&H1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2

Private p_oProgress As New frmProgress

Private pbDisplayed As Boolean
Private pbMoveProg As Boolean

Property Let ProgressStatus(Value As String)
   If pbDisplayed = False Then Exit Property
   
   p_oProgress.ProgressStatus = Value
End Property

Property Let PrimaryRemarks(Value As String)
   p_oProgress.PrimaryRemarks = Value
End Property

Property Let SecondaryRemarks(Value As String)
   p_oProgress.SecondaryRemarks = Value
End Property

Property Get Cancelled() As Boolean
   Cancelled = p_oProgress.Cancelled
End Property

Property Get PriMaxValue() As Long
   PriMaxValue = p_oProgress.PrimaryMaxValue
End Property

Property Let PriMaxValue(ByVal Value As Long)
   p_oProgress.PrimaryMaxValue = Value
   pbMoveProg = Value <> 0
End Property

Property Get SecMaxValue() As Long
   SecMaxValue = p_oProgress.SecondaryMaxValue
End Property

Property Let SecMaxValue(ByVal Value As Long)
   p_oProgress.SecondaryMaxValue = Value
   pbMoveProg = Value <> 0
End Property

Function MoveProgress(Optional SecRemarks As Variant, Optional PriRemarks As Variant) As Boolean
   MoveProgress = True
   If pbDisplayed = False Then Exit Function
   If pbMoveProg = False Then Exit Function
   
   If Not IsMissing(PriRemarks) Then p_oProgress.PrimaryRemarks = PriRemarks
   If Not IsMissing(SecRemarks) Then p_oProgress.SecondaryRemarks = SecRemarks
   
   p_oProgress.MoveProgress
   MoveProgress = Not p_oProgress.Cancelled
End Function

Sub InitProgress(Optional ProgressStatus As String, Optional PriMaxValue As Variant, Optional SecMaxValue As Variant)

   Load p_oProgress
   If ProgressStatus <> Empty Then p_oProgress.ProgressStatus = ProgressStatus
   p_oProgress.Show

   SetWindowPos p_oProgress.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, (SWP_NOMOVE Or SWP_NOSIZE)
   
   pbMoveProg = False
   If Not IsMissing(SecMaxValue) Then
      p_oProgress.SecondaryMaxValue = SecMaxValue
      If Not IsMissing(PriMaxValue) Then
         p_oProgress.PrimaryMaxValue = PriMaxValue
      Else
         p_oProgress.PrimaryMaxValue = 1
      End If
      pbMoveProg = True
   Else
      If Not IsMissing(PriMaxValue) Then
         p_oProgress.PrimaryMaxValue = 1
         p_oProgress.SecondaryMaxValue = PriMaxValue
         pbMoveProg = True
      End If
   End If
   
   pbDisplayed = True
End Sub

Sub CloseProgress()
   Unload p_oProgress
   Set p_oProgress = Nothing
End Sub
