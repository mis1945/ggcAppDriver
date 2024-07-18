VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const DT_CENTER = &H1

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private WithEvents p_oImage As PictureBox
Attribute p_oImage.VB_VarHelpID = -1
Private WithEvents p_oForm As Form
Attribute p_oForm.VB_VarHelpID = -1
Private p_oAppDrivr As clsAppDriver
Private p_nFormType As xeFormType
Private p_bDisClose As Boolean

Property Set Form(oForm As Object)
   Set p_oForm = oForm
End Property

Property Set AppDriver(oAppDriver As clsAppDriver)
   Set p_oAppDrivr = oAppDriver
End Property

Property Let DisableClose(ByVal Value As Boolean)
   p_bDisClose = Value
End Property

Sub ApplySkin(Optional FormType As Variant)
   Dim loControl As Control
   Dim lnOldMode As Integer
   Dim lasTag() As String
   Dim lsSkin As String
   Dim lnCtr As Integer

   If p_oForm Is Nothing Then Exit Sub
   p_nFormType = xeFormMaintenance
   If Not IsMissing(FormType) Then p_nFormType = FormType

   On Error Resume Next
   
   With p_oForm
      lnOldMode = .ScaleMode
      Set p_oImage = .Controls.Add("VB.PictureBox", "PictureBox")
      .Font = "MS San Serif"
      .FontSize = 8
      .FontBold = True
      .Tag = "wb0;et0"
      .AutoRedraw = True
      .ScaleMode = 1
      
      p_oImage.Picture = LoadPicture(p_oAppDrivr.AppPath & "\Images\x_button.ico")
      p_oImage.Visible = True
      p_oImage.TabStop = False
      p_oImage.CausesValidation = False
      p_oImage.ScaleMode = vbTwips
      p_oImage.BorderStyle = 0
      p_oImage.Appearance = 0
      p_oImage.Height = 225
      p_oImage.Height = 225
      p_oImage.Top = 70
      p_oImage.Left = .ScaleWidth - 265
      
      For Each loControl In .Controls
         lnCtr = lnCtr + 1
         If loControl.Tag = Empty Then
            loControl.ForeColor = p_oAppDrivr.getColor("ET0")
            loControl.BackColor = p_oAppDrivr.getColor("EB0")
         Else
            lasTag = Split(loControl.Tag, ";")
            loControl.ForeColor = p_oAppDrivr.getColor(lasTag(0))
            
            Select Case UBound(lasTag)
            Case 0
               loControl.BackColor = p_oAppDrivr.getColor("EB0")
            Case 1
               loControl.BackColor = p_oAppDrivr.getColor(lasTag(1))
            Case 2
               loControl.BackColor = p_oAppDrivr.getColor(lasTag(1))
               loControl.ForeColorFixed = p_oAppDrivr.getColor(lasTag(2))
               loControl.BackColorFixed = p_oAppDrivr.getColor("EB0")
            Case 3
               loControl.BackColor = p_oAppDrivr.getColor(lasTag(1))
               loControl.ForeColorFixed = p_oAppDrivr.getColor(lasTag(2))
               loControl.BackColorFixed = p_oAppDrivr.getColor(lasTag(3))
            End Select
         End If

         If TypeName(loControl) = "xrButton" Then
            loControl.BackColor = p_oAppDrivr.getColor("HB1")
            loControl.BackColorDown = p_oAppDrivr.getColor("HB4")
            loControl.BorderColorFocus = p_oAppDrivr.getColor("BC0")
            loControl.BorderColorHover = p_oAppDrivr.getColor("BC1")
            loControl.ForeColor = p_oAppDrivr.getColor("ET0")
         End If
      Next
      If .Tag <> "" Then
         lasTag = Split(.Tag, ";")
         .BackColor = p_oAppDrivr.getColor(lasTag(0))
      End If
      
      .ScaleMode = lnOldMode
   End With
End Sub

Private Sub p_oForm_Activate()
   Dim lnOldMode As Integer
   Dim hRgn As Long
   Dim lbhRPen As Long
   Dim R As RECT
   Dim nCount As Integer
   Dim nCtr As Integer

   With p_oForm
      lnOldMode = .ScaleMode
      .Cls
      .ScaleMode = 1
      .Picture = LoadPicture(p_oAppDrivr.AppPath & "\Images\" & p_oAppDrivr.getPicture(xeFormMaintenance))
      TextOut .hdc, 25, 6, .Caption, Len(.Caption)

      ' draw the region now
      BeginPath .hdc
      ' this will create the title bar
      RoundRect .hdc, 0, 0, .ScaleWidth, 26, 25, 25
      
      ' then set scale mode to pixel for, coz it's hard to calculate the
      '  rounded side, accrdng to my two colleagues
      .ScaleMode = 3
      ' the body of the form varies depending on the type of form
      Select Case p_nFormType
      Case xeFormMaintenance
         RoundRect .hdc, 0, 29, .ScaleWidth + 5, .ScaleHeight - 60, 0, 0
         RoundRect .hdc, 0, .ScaleHeight - 57, .ScaleWidth + 5, .ScaleHeight, 0, 0
         EndPath .hdc
         
         'convert the path to a region
         hRgn = PathToRegion(.hdc)
          
         'set the Window-region
         SetWindowRgn .hWnd, hRgn, True
         
         'Crete a new pen
         lbhRPen = CreatePen(0, 1, &HE0E0E0)
   
         'Select our pen into the form's device context and delete the old pen
         DeleteObject SelectObject(.hdc, lbhRPen)
   
         'Draw a rectangle
         RoundRect .hdc, 0, 29, .ScaleWidth, .ScaleHeight - 61, 0, 0
         RoundRect .hdc, 0, .ScaleHeight - 57, .ScaleWidth, .ScaleHeight - 1, 0, 0
               
   Case xeFormQuickSearch
         RoundRect .hdc, 0, 29, 117, 143, 10, 10
         RoundRect .hdc, 121, 29, .ScaleWidth, 143, 10, 10
         RoundRect .hdc, 0, 147, .ScaleWidth, .ScaleHeight, 10, 10
         EndPath .hdc
         
         'convert the path to a region
         hRgn = PathToRegion(.hdc)
          
         'set the Window-region
         SetWindowRgn .hWnd, hRgn, True

         'Crete a new pen
         lbhRPen = CreatePen(0, 2, &HE0E0E0)
         
         'Select our pen into the form's device context and delete the old pen
         DeleteObject SelectObject(.hdc, lbhRPen)
   
         'Draw a rounded rectangle
         RoundRect .hdc, 0, 29, 117, 143, 10, 10
         RoundRect .hdc, 121, 29, .ScaleWidth, 143, 10, 10
         RoundRect .hdc, 0, 147, .ScaleWidth, .ScaleHeight, 10, 10

      Case xeFormTransaction
         RoundRect .hdc, 0, 29, 96, .ScaleHeight - 100, 0, 0
         RoundRect .hdc, 99, .ScaleHeight, .ScaleWidth, 29, 0, 0

         'close the path bracket
         EndPath .hdc
          
         'convert the path to a region
         hRgn = PathToRegion(.hdc)
          
         'set the Window-region
         SetWindowRgn .hWnd, hRgn, True

         'Crete a new pen
         lbhRPen = CreatePen(0, 1, &HE0E0E0)
   
         'Select our pen into the form's device context and delete the old pen
         DeleteObject SelectObject(.hdc, lbhRPen)
   
         'Draw a rectangle
         RoundRect .hdc, 0, 29, 95, .ScaleHeight - 101, 0, 0
         RoundRect .hdc, 99, .ScaleHeight - 1, .ScaleWidth - 1, 29, 0, 0

      Case xeFormLedger
         RoundRect .hdc, -1, 29, .ScaleWidth + 20, .ScaleHeight, 0, 0
    
         nCount = 50
            For nCtr = 1 To 10
               RoundRect .hdc, .ScaleWidth - 20, nCount, .ScaleWidth - 10, nCount + 10, 10, 10
               nCount = nCount + 50
            Next
         'close the path bracket
         EndPath .hdc
          
         'convert the path to a region
         hRgn = PathToRegion(.hdc)
          
         'set the Window-region
         SetWindowRgn .hWnd, hRgn, True

         'Crete a new pen
         lbhRPen = CreatePen(0, 1, &HE0E0E0)
   
         'Select our pen into the form's device context and delete the old pen
         DeleteObject SelectObject(.hdc, lbhRPen)
   
         'Draw a rectangle
         RoundRect .hdc, 0, 29, .ScaleWidth, .ScaleHeight - 1, 0, 0

      Case xeFormTransDetail
         RoundRect .hdc, .ScaleWidth - 96, 29, .ScaleWidth, 170, 0, 0
         RoundRect .hdc, 0, .ScaleHeight, .ScaleWidth - 98, 29, 0, 0

         'close the path bracket
         EndPath .hdc
          
         'convert the path to a region
         hRgn = PathToRegion(.hdc)
          
         'set the Window-region
         SetWindowRgn .hWnd, hRgn, True

         'Crete a new pen
         lbhRPen = CreatePen(0, 1, &HE0E0E0)
   
         'Select our pen into the form's device context and delete the old pen
         DeleteObject SelectObject(.hdc, lbhRPen)
   
         'Draw a rectangle
         RoundRect .hdc, .ScaleWidth - 96, 29, .ScaleWidth - 1, 169, 0, 0
         RoundRect .hdc, 0, .ScaleHeight - 1, .ScaleWidth - 99, 29, 0, 0

     Case xeFormTransMaintenance
         RoundRect .hdc, .ScaleWidth - 96, 29, .ScaleWidth, .ScaleHeight - 182, 0, 0
         RoundRect .hdc, 0, .ScaleHeight, .ScaleWidth - 98, 29, 0, 0

         'close the path bracket
         EndPath .hdc
          
         'convert the path to a region
         hRgn = PathToRegion(.hdc)
          
         'set the Window-region
         SetWindowRgn .hWnd, hRgn, True

         'Crete a new pen
         lbhRPen = CreatePen(0, 1, &HE0E0E0)
   
         'Select our pen into the form's device context and delete the old pen
         DeleteObject SelectObject(.hdc, lbhRPen)
   
         'Draw a rectangle
         RoundRect .hdc, .ScaleWidth - 96, 29, .ScaleWidth - 1, .ScaleHeight - 183, 0, 0
         RoundRect .hdc, 0, .ScaleHeight - 1, .ScaleWidth - 99, 29, 0, 0

      Case xeFormTransEqualLeft
         RoundRect .hdc, 0, 29, 96, .ScaleHeight, 0, 0
         RoundRect .hdc, 99, .ScaleHeight, .ScaleWidth, 29, 0, 0

         'close the path bracket
         EndPath .hdc
          
         'convert the path to a region
         hRgn = PathToRegion(.hdc)
          
         'set the Window-region
         SetWindowRgn .hWnd, hRgn, True

         'Crete a new pen
         lbhRPen = CreatePen(0, 1, &HE0E0E0)
   
         'Select our pen into the form's device context and delete the old pen
         DeleteObject SelectObject(.hdc, lbhRPen)
   
         'Draw a rectangle
         RoundRect .hdc, 0, 29, 95, .ScaleHeight - 1, 0, 0
         RoundRect .hdc, 99, .ScaleHeight - 1, .ScaleWidth - 1, 29, 0, 0
      
      Case xeFormTransEqualRight
         RoundRect .hdc, .ScaleWidth - 96, 29, .ScaleWidth, .ScaleHeight, 0, 0
         RoundRect .hdc, 0, .ScaleHeight, .ScaleWidth - 98, 29, 0, 0

         'close the path bracket
         EndPath .hdc
          
         'convert the path to a region
         hRgn = PathToRegion(.hdc)
          
         'set the Window-region
         SetWindowRgn .hWnd, hRgn, True

         'Crete a new pen
         lbhRPen = CreatePen(0, 1, &HE0E0E0)
   
         'Select our pen into the form's device context and delete the old pen
         DeleteObject SelectObject(.hdc, lbhRPen)
   
         'Draw a rectangle
         RoundRect .hdc, .ScaleWidth - 96, 29, .ScaleWidth - 1, .ScaleHeight - 1, 0, 0
         RoundRect .hdc, 0, .ScaleHeight - 1, .ScaleWidth - 99, 29, 0, 0
      End Select
      .ScaleMode = lnOldMode
   End With
End Sub

Private Sub p_oForm_Deactivate()
   With p_oForm
      .Picture = LoadPicture(p_oAppDrivr.AppPath & "\Images\tBarGry.jpg")
      TextOut .hdc, 25, 6, .Caption, Len(.Caption)
   End With
End Sub

Private Sub p_oForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      If y < 360 Then
         Call ReleaseCapture
         Call SendMessage(p_oForm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
      End If
   End If
End Sub

Private Sub p_oForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Screen.MousePointer <> 0 Then Screen.MousePointer = 0
End Sub

Private Sub p_oForm_Unload(Cancel As Integer)
   Set p_oImage = Nothing
End Sub

Private Sub p_oImage_Click()
   If Not p_bDisClose Then Unload p_oForm
End Sub

Private Sub p_oImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   p_oImage.Picture = LoadPicture(p_oAppDrivr.AppPath & "\Images\x_button1.ico")
End Sub

Private Sub p_oImage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   p_oImage.Picture = LoadPicture(p_oAppDrivr.AppPath & "\Images\x_button.ico")
End Sub
