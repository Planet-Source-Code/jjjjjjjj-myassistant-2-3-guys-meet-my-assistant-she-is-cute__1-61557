Attribute VB_Name = "mdlDraw"

Option Explicit

Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long


Public Sub DrawBorder(Frm As Form)
Dim hRgn As Long
Dim hBrush As Long

    ' Create form shape
    hRgn = CreateRoundRectRgn(0, 0, Frm.ScaleWidth, Frm.ScaleHeight, 20, 20)
    SetWindowRgn Frm.hwnd, hRgn, True
    
    ' Create brush/Fill it
    hRgn = CreateRoundRectRgn(0, 0, Frm.ScaleWidth - 1, Frm.ScaleHeight - 1, 20, 20)
    hBrush = CreateSolidBrush(RGB(255, 175, 1))
    FillRgn Frm.hdc, hRgn, hBrush
    DeleteObject hBrush
    
    hRgn = CreateRoundRectRgn(1, 1, Frm.ScaleWidth - 2, Frm.ScaleHeight - 2, 20, 20)
    hBrush = CreateSolidBrush(RGB(255, 128, 64))
    FillRgn Frm.hdc, hRgn, hBrush
    DeleteObject hBrush: DeleteObject hRgn

    hRgn = CreateRoundRectRgn(5, 15, Frm.ScaleWidth - 6, Frm.ScaleHeight - 6, 15, 15)
    hBrush = CreateSolidBrush(RGB(255, 175, 1))
    FillRgn Frm.hdc, hRgn, hBrush
    DeleteObject hBrush: DeleteObject hRgn
    
    hRgn = CreateRoundRectRgn(6, 16, Frm.ScaleWidth - 7, Frm.ScaleHeight - 7, 15, 15)
    hBrush = CreateSolidBrush(RGB(255, 185, 0))
    FillRgn Frm.hdc, hRgn, hBrush
    DeleteObject hBrush: DeleteObject hRgn
    
    Frm.Refresh
    
End Sub
