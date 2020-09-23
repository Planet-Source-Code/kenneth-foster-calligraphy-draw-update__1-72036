Attribute VB_Name = "modGradient"
Option Explicit

Type RECT
   Left As Long
   top As Long
   Right As Long
   Bottom As Long
End Type

Declare Function CreateSolidBrush Lib "gdi32" _
  (ByVal crColor As Long) As Long

Declare Function DeleteObject Lib "gdi32" _
  (ByVal hObject As Long) As Long

Declare Function FillRect Lib "user32" _
  (ByVal hDC As Long, lpRect As RECT, _
  ByVal hBrush As Long) As Long

Public Sub DrawGradient(lDestHDC As Long, _
  lDestWidth As Long, lDestHeight As Long, _
  lStartColor As Long, lEndColor As Long, _
  iStyle As Integer)

   Dim udtRect As RECT

   Dim iBlueStart As Integer
   Dim iBlueEnd As Integer
   Dim iRedStart As Integer
   Dim iRedEnd As Integer
   Dim iGreenStart As Integer
   Dim iGreenEnd As Integer

   Dim hBrush As Long

   On Error Resume Next

   'Calculate the beginning colors
   iBlueStart = Int(lStartColor / &H10000)
   iGreenStart = Int(lStartColor - (iBlueStart * &H10000)) \ _
         CLng(&H100)
   iRedStart = lStartColor - (iBlueStart * &H10000) - _
         CLng(iGreenStart * CLng(&H100))

   'Calculate the End colors
   iBlueEnd = Int(lEndColor / &H10000)
   iGreenEnd = Int(lEndColor - (iBlueEnd * &H10000)) \ CLng(&H100)
   iRedEnd = lEndColor - (iBlueEnd * &H10000) - _
         CLng(iGreenEnd * CLng(&H100))

   Const intBANDWIDTH = 1

   Dim sngBlueCur As Single
   Dim sngBlueStep As Single
   Dim sngGreenCur As Single
   Dim sngGreenStep As Single
   Dim sngRedCur As Single
   Dim sngRedStep As Single

   Dim iHeight As Integer
   Dim iWidth As Integer
   Dim intY As Integer
   Dim iDrawEnd As Integer

   Dim lReturn As Long

   iHeight = lDestHeight
   iWidth = lDestWidth

   sngBlueCur = iBlueStart
   sngGreenCur = iGreenStart
   sngRedCur = iRedStart

   'Calculate the size of the color bars
   If iStyle = 0 Then
      sngBlueStep = intBANDWIDTH * _
         (iBlueEnd - iBlueStart) / (iWidth - 60) * 15
      sngGreenStep = intBANDWIDTH * _
         (iGreenEnd - iGreenStart) / (iWidth - 60) * 15
      sngRedStep = intBANDWIDTH * _
         (iRedEnd - iRedStart) / (iWidth - 60) * 15
      With udtRect
         .Left = 0
         .top = 0
         .Right = intBANDWIDTH + 2
         .Bottom = iHeight / 15 - 2
      End With
      iDrawEnd = iWidth
   ElseIf iStyle = 1 Then
      sngBlueStep = intBANDWIDTH * _
         (iBlueEnd - iBlueStart) / (iHeight - 60) * 15
      sngGreenStep = intBANDWIDTH * _
         (iGreenEnd - iGreenStart) / (iHeight - 60) * 15
      sngRedStep = intBANDWIDTH * _
         (iRedEnd - iRedStart) / (iHeight - 60) * 15
      With udtRect
         .Left = 0
         .top = 0
         .Right = iWidth / 15 - 2
         .Bottom = intBANDWIDTH + 2
      End With
      iDrawEnd = iHeight
   End If

   'Draw the Gradient
   For intY = 0 To (iDrawEnd / 15) - 5 Step intBANDWIDTH
      hBrush = CreateSolidBrush(RGB(sngRedCur, sngGreenCur, sngBlueCur))
      lReturn = FillRect(lDestHDC, udtRect, hBrush)

      lReturn = DeleteObject(hBrush)
      sngBlueCur = sngBlueCur + sngBlueStep
      sngGreenCur = sngGreenCur + sngGreenStep
      sngRedCur = sngRedCur + sngRedStep
      If iStyle = 0 Then
         udtRect.Left = udtRect.Left + intBANDWIDTH
         udtRect.Right = udtRect.Right + intBANDWIDTH
      ElseIf iStyle = 1 Then
         udtRect.top = udtRect.top + intBANDWIDTH
         udtRect.Bottom = udtRect.Bottom + intBANDWIDTH
      End If
   Next
End Sub




