VERSION 5.00
Begin VB.UserControl ColorButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3450
   ClipControls    =   0   'False
   EditAtDesignTime=   -1  'True
   FillColor       =   &H00A3E8FC&
   ForeColor       =   &H00FFFFFF&
   MaskColor       =   &H00000000&
   MouseIcon       =   "ColorButton.ctx":0000
   MousePointer    =   99  'Custom
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   102
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   230
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2880
      Top             =   240
   End
End
Attribute VB_Name = "ColorButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'[Types]
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type POINT_API
    X As Long
    Y As Long
End Type

'[Propertiese]
Private ctlBackColor    As Long
Private ctlBorderColor  As Long
Private ctlPicture      As Picture
Private ctlCaption      As String
Private ctlFont         As StdFont
Private ctlForeColor    As Long

'[Variables]
Dim ScrX As Long, ScrY As Long

'[Events]
Public Event Click()
Public Event DbClick()
Public Event MouseIn()
Public Event MouseOut()
Public Event MouseMove()
Public Event MouseDown()
Public Event MouseUp()
Public Event KeyPress()
Public Event KeyUp()
Public Event KeyDown()

'[Function APIs]
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINT_API) As Long

'[Drawing APIs]
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long

'[Properties]
'------------

Public Property Get Caption() As String
    Caption = ctlCaption
End Property

Public Property Let Caption(ByVal vNewCaption As String)
    ctlCaption = vNewCaption
    Call Refresh
End Property


Public Property Get BackColor() As OLE_COLOR
    BackColor = ctlBackColor
End Property

Public Property Let BackColor(ByVal vNewColor As OLE_COLOR)
    ctlBackColor = vNewColor
    Call Refresh
End Property


Public Property Get BorderColor() As OLE_COLOR
    BorderColor = ctlBorderColor
End Property

Public Property Let BorderColor(ByVal vNewColor As OLE_COLOR)
    ctlBorderColor = vNewColor
    Call Refresh
End Property


Public Property Get Picture() As Picture
    Set Picture = ctlPicture
End Property

Public Property Set Picture(ByVal NewPicture As Picture)
    Set ctlPicture = NewPicture
    Call Refresh
End Property


Public Property Get Font() As StdFont
    Set Font = ctlFont
End Property

Public Property Set Font(ByVal vNewFont As StdFont)
    Set ctlFont = vNewFont
    Call Refresh
End Property


Public Property Get ForeColor() As OLE_COLOR
    ForeColor = ctlForeColor
End Property

Public Property Let ForeColor(ByVal vNewColor As OLE_COLOR)
    ctlForeColor = vNewColor
    Call Refresh
End Property


Private Sub tmrUpdate_Timer()
    Call Refresh
    If MouseIsOver = False Then tmrUpdate.Enabled = False: RaiseEvent MouseOut
End Sub

'------------------------------------------------------------------------------
'[Event Handle]
'------------------------------------------------------------------------------

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DbClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tmrUpdate.Enabled = False Then tmrUpdate.Enabled = True: RaiseEvent MouseIn
    RaiseEvent MouseMove
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp
End Sub

Private Sub UserControl_Initialize()
    ScrX = Screen.TwipsPerPixelX
    ScrY = Screen.TwipsPerPixelY
    Me.Caption = "ColorButton"
    Me.BackColor = vbRed
    Me.ForeColor = vbWhite
    Me.BorderColor = vbBlack
End Sub

Private Sub UserControl_Resize()
    Call Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", ctlCaption, "ColorButton"
    PropBag.WriteProperty "Picture", ctlPicture, Nothing
    PropBag.WriteProperty "BackColor", ctlBackColor, vbRed
    PropBag.WriteProperty "BorderColor", ctlBorderColor, 0
    PropBag.WriteProperty "Font", ctlFont, UserControl.Font
    PropBag.WriteProperty "ForeColor", ctlForeColor, vbWhite
    Call Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set ctlPicture = PropBag.ReadProperty("Picture", Nothing)
    ctlCaption = PropBag.ReadProperty("Caption", "ColorButton")
    ctlBackColor = PropBag.ReadProperty("BackColor", vbRed)
    ctlBorderColor = PropBag.ReadProperty("BorderColor", 0)
    ctlForeColor = PropBag.ReadProperty("ForeColor", vbWhite)
    Set ctlFont = PropBag.ReadProperty("Font", UserControl.Font)
    Call Refresh
End Sub

Public Sub Refresh()
    Cls
    UserControl.ForeColor = ctlForeColor
    UserControl.FontBold = MouseIsOver
    UserControl.BackColor = ctlBorderColor
    DrawSmoothRect 0, 0, Width / ScrY, Height / ScrX
    DrawPicture
'    DrawCaption
End Sub

Private Function IsBitmap(vMap As StdPicture) As Boolean
On Error GoTo Handle
    If vMap.Height > 0 And vMap.Width > 0 And vMap.Handle > 0 Then IsBitmap = True Else IsBitmap = False
Exit Function
Handle:
    IsBitmap = False
End Function

Private Function MouseIsOver() As Boolean
Dim Rct As RECT
Dim MousePos As POINT_API
    Call GetCursorPos(MousePos)
    ScreenToClient UserControl.hwnd, MousePos
    GetClientRect UserControl.hwnd, Rct
    If PtInRect(Rct, MousePos.X, MousePos.Y) <> 0 Then
        MouseIsOver = True
    Else
        MouseIsOver = False
    End If
End Function



'[This sub will draw smooth edged rectangles]
'--------------------------------------------
Private Sub DrawSmoothRect(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Dim Curvature As Long
Dim hBrush As Long, hRgn As Long
    If (X2 - X1) > (Y2 - Y1) Then
        Curvature = (X2 - X1) / 5
    Else
        Curvature = (Y2 - Y1) / 5
    End If
    hBrush = CreateSolidBrush(ctlBackColor)
    hRgn = CreateRoundRectRgn(X1, Y1, X2, Y2, Curvature, Curvature)
    SetWindowRgn hwnd, hRgn, True
    hRgn = CreateRoundRectRgn(X1 + 1, Y1 + 1, X2 - 1, Y2 - 1, Curvature, Curvature)
    FillRgn hdc, hRgn, hBrush
    DeleteObject hRgn
    DeleteObject hBrush
End Sub

'[Draw Picture/Caption]
'----------------------
Private Sub DrawPicture()
Dim Rct As RECT
Dim fWidth As Long, fHeight As Long
Dim fLeft As Long, fTop As Long
    With Rct
    If IsBitmap(ctlPicture) = True Then
        If ctlCaption = "" Then
            fHeight = ScaleY(ctlPicture.Height)
        Else
            fHeight = ScaleY(ctlPicture.Height) + TextHeight(ctlCaption)
        End If
        fTop = (Height / ScrY - fHeight) / 2
        fWidth = ScaleX(ctlPicture.Width)
        fLeft = (Width / ScrX - fWidth) / 2
        fHeight = ScaleY(ctlPicture.Height)
        UserControl.PaintPicture ctlPicture, fLeft, fTop, fWidth, fHeight
        .Top = fTop + fHeight
    Else
        .Top = Height / ScrY / 2 - TextHeight(ctlCaption) / 2
    End If
        .Left = 0
        .Right = Width / ScrX
        .Bottom = Height / ScrY
    End With
    DrawText hdc, ctlCaption, Len(ctlCaption), Rct, 1
End Sub

