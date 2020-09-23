Attribute VB_Name = "mdlSkin"
Option Explicit

'[APIs]
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'[This function will set your form smoothly curved ]
'=============================================================
Public Sub SmoothForm(Frm As Form, Optional ByVal Curvature As Double = 25)
Dim hRgn As Long
Dim X1 As Long, Y1 As Long
    X1 = Frm.Width / Screen.TwipsPerPixelX
    Y1 = Frm.Height / Screen.TwipsPerPixelY
    hRgn = CreateRoundRectRgn(1, 1, X1, Y1, Curvature, Curvature)
    SetWindowRgn Frm.hwnd, hRgn, True
    DeleteObject hRgn
End Sub
'=============================================================

'[Arrange the images]
'=============================================================
Sub SkinForm1(Frm As Form)
'On Error Resume Next

With Frm
    .UL.Move 0, 0
    .BL.Move 0, .Height - .BL.Height - 0
    .LS.Move 0, .UL.Height, .LS.Width, .Height - .UL.Height - .BL.Height
    .UM.Move .UL.Width, 0, .Width - .UL.Width - .UR.Width, .UM.Height
    .UR.Move .Width - .UR.Width - 0, 0
    .BR.Move .Width - .BR.Width - 0, .Height - .BR.Height - 0
    .RS.Move .Width - .RS.Width - 0, .UR.Height, .RS.Width, .Height - .UR.Height - .BR.Height
    .BM.Move .BL.Width, .Height - .BM.Height - 0, .Width - .BL.Width - .BR.Width
'    .LblRestore.Left = .Width - 430,
End With

'With Frm
'    .UL.Top = 0
'    .UL.Left = 0
'    .BL.Top = .Height - .BL.Height - 0
'    .BL.Left = 0
'    .LS.Top = .UL.Height
'    .LS.Height = .Height - .UL.Height - .BL.Height
'    .LS.Left = 0
'    .UM.Width = .Width - .UL.Width - .UR.Width
'    .UM.Left = .UL.Width
'    .UM.Top = 0
'    .UR.Top = 0
'    .UR.Left = .Width - .UR.Width - 0
'    .BR.Left = .Width - .BR.Width - 0
'    .BR.Top = .Height - .BR.Height - 0
'    .RS.Left = .Width - .RS.Width - 0
'    .RS.Height = .Height - .UR.Height - .BR.Height
'    .RS.Top = .UR.Height
'    .BM.Top = .Height - .BM.Height - 0
'    .BM.Left = .BL.Width
'    .BM.Width = .Width - .BL.Width - .BR.Width
''    .LblMinimized.Left = .Width - 540
'    .LblRestore.Left = .Width - 430
''    .LblMaximized.Left = .Width - 200
'End With
End Sub
'=============================================================

'[This function can Set your form as BorderStyle=0 ]
'=============================================================
Public Sub SetZeroBorder(Frm As Form)
Dim hRgn As Long
Dim fScaleMode As Long
Dim ScrX As Long, ScrY As Long
Dim fLeft As Long, fTop As Long
Dim fBottom As Long, fRight As Long
    ScrX = Screen.TwipsPerPixelX
    ScrY = Screen.TwipsPerPixelY
    With Frm
        fScaleMode = .ScaleMode
        .ScaleMode = 1
        fLeft = (.Width - .ScaleWidth) / 2 / ScrX
        fTop = (.Height - .ScaleHeight) / ScrY - fLeft
        fRight = .Width / ScrX - fLeft
        fBottom = .Height / ScrY - fLeft
        hRgn = CreateRoundRectRgn(fLeft, fTop, fRight, fBottom, 11, 11)
        SetWindowRgn .hwnd, hRgn, True
        .ScaleMode = fScaleMode
        DeleteObject hRgn
    End With
End Sub
