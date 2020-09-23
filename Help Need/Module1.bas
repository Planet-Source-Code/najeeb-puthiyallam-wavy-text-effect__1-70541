Attribute VB_Name = "Module1"
Option Explicit

Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Public Function createSkinnedForm(SkinnedForm As Frame, skinSrc As PictureBox, Optional transparentColor As Long) As Long
Const RGN_OR = 2
Dim glSkinImage As Long
Dim glHeight    As Long
Dim glwidth     As Long
Dim lReturn     As Long
Dim lRgnTmp     As Long
Dim lSkinRgn    As Long
Dim lStart      As Long
Dim lRow        As Long
Dim lCol        As Long
skinSrc.AutoSize = True
lSkinRgn = CreateRectRgn(0, 0, 0, 0)
With skinSrc
    .AutoRedraw = True
    glHeight = .Height / Screen.TwipsPerPixelY
    glwidth = .Width / Screen.TwipsPerPixelX
    If transparentColor < 1 Then transparentColor = GetPixel(.hDC, 0, 0)
    For lRow = 0 To glHeight - 1
        lCol = 0
        Do While lCol < glwidth
            Do While lCol < glwidth And GetPixel(.hDC, lCol, lRow) = transparentColor
                lCol = lCol + 1
            Loop
            If lCol < glwidth Then
                lStart = lCol
                Do While lCol < glwidth And GetPixel(.hDC, lCol, lRow) <> transparentColor
                    lCol = lCol + 1
                Loop
                If lCol > glwidth Then lCol = glwidth
                lRgnTmp = CreateRectRgn(lStart, lRow, lCol, lRow + 1)
                lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
                Call DeleteObject(lRgnTmp)
            End If
        Loop
    Next
End With
Call SetWindowRgn(SkinnedForm.hWnd, lSkinRgn, True)
skinSrc.Picture = LoadPicture("")
End Function




