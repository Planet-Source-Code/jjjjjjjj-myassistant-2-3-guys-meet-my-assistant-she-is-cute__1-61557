VERSION 5.00
Begin VB.UserControl ListBoxEX 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   162
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   155
   ToolboxBitmap   =   "ListBoxEx.ctx":0000
   Begin VB.VScrollBar VScroll 
      Height          =   2415
      LargeChange     =   10
      Left            =   2040
      Max             =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picDraw 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2430
      Left            =   0
      MouseIcon       =   "ListBoxEx.ctx":0312
      ScaleHeight     =   162
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "ListBoxEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^¶¶^^^^^^¶¶^^^^^^^^^^^^^¶¶¶¶¶¶¶^^^^^^^^^^^^^^^^¶¶¶¶¶¶¶^^^^^^^^^^^^^¶¶^^^^¶¶^^^¶¶^^^^^$
'$^^^^¶¶^^^^^^^^^^^^^^^^^¶¶^^¶¶^^^^¶¶^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^¶¶^^^^¶¶^^¶¶¶^^^^^$
'$^^^^¶¶^^^^^^^^^^^^^^^^^¶¶^^¶¶^^^^¶¶^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^¶¶^^¶¶^^^^¶¶^^^^^$
'$^^^^¶¶^^^^^^¶¶^^^¶¶¶¶^^¶¶¶^¶¶^^^^¶¶^^¶¶¶¶^^¶¶^^¶¶^¶¶^^^^^^^¶¶^^¶¶^^^^^^¶¶^^¶¶^^^^¶¶^^^^^$
'$^^^^¶¶^^^^^^¶¶^^¶¶^^¶¶^¶¶^^¶¶¶¶¶¶¶^^¶¶^^¶¶^^¶¶¶¶^^¶¶¶¶¶¶^^^^¶¶¶¶^^^^^^^¶¶^^¶¶^^^^¶¶^^^^^$
'$^^^^¶¶^^^^^^¶¶^^^¶¶¶^^^¶¶^^¶¶^^^^¶¶^¶¶^^¶¶^^^¶¶^^^¶¶^^^^^^^^^¶¶^^^^^^^^^¶¶¶¶^^^^^¶¶^^^^^$
'$^^^^¶¶^^^^^^¶¶^^^^^¶¶^^¶¶^^¶¶^^^^¶¶^¶¶^^¶¶^^^¶¶^^^¶¶^^^^^^^^^¶¶^^^^^^^^^¶¶¶¶^^^^^¶¶^^^^^$
'$^^^^¶¶^^^^^^¶¶^^¶¶^^¶¶^¶¶^^¶¶^^^^¶¶^¶¶^^¶¶^^¶¶¶¶^^¶¶^^^^^^^^¶¶¶¶^^^^^^^^^¶¶^^^^^^¶¶^^^^^$
'$^^^^¶¶¶¶¶¶¶^¶¶^^^¶¶¶¶^^^¶¶^¶¶¶¶¶¶¶^^^¶¶¶¶^^¶¶^^¶¶^¶¶¶¶¶¶¶^^¶¶^^¶¶^^^^^^^^¶¶^^^^^^¶¶^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'------------------------------------------------------------------------------------------
' SourceCode : ListBoxEX V1
' Auther     : Jim Jose
' Email      : jimjosev33@yahoo.com
' Date       : 3-6-2005
' Purpose    : An upgraded version of VBListBox with Icons and many more
' Comment    : This is the first version of this control.
'            : This version aimed for a clear and simple code.
'            : Use your imaginations to visualize more features.
'            : Please send me your better ideas and additional features you need.
' CopyRight  : JimJose © Gtech Creations - 2005
'------------------------------------------------------------------------------------------

Option Explicit

'[APIs]
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long

'[Types]
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

'[Enums]
Public Enum ListStyleEnum
    [Graphical] = 1
    [Standard] = 0
End Enum

Public Enum AppearanceEnum
    [Flat] = 0
    [3D] = 1
End Enum

Public Enum GradientTypeConstants
    [None_Gradient] = 0
    [Horizontal] = 1
    [Vertical] = 2
End Enum

Public Enum BorderEnum
    [None] = 0
    [Fixed Single] = 1
End Enum

Public Enum SortOrderEnum
    [Sort_None] = 0
    [Sort_Ascending] = -1
    [Sort_Desending] = 1
End Enum

'[Local Variables]
Private m_SelItem    As Long
Private m_iHeight    As Long
Private m_iCount     As Long
Private m_iTop       As Long
Private m_hMode      As Long
Private m_KeyControl As Boolean

'[Data Storage]
Private m_Items      As New Collection

'[Property Variables]
Private m_Picture       As New StdPicture
Private m_ListIcon      As New StdPicture
Private m_ForeColor     As OLE_COLOR
Private m_Font          As New StdFont
Private m_SelColor      As OLE_COLOR
Private m_FullRowSel    As Boolean
Private m_SortOrder     As SortOrderEnum
Private m_SelForeColor  As OLE_COLOR
Private m_StrechIcon    As Boolean
Private m_IconFocus     As Boolean
Private m_TextAllineMent As AlignmentConstants

Private m_Gradient      As GradientTypeConstants
Private m_StartColor    As OLE_COLOR
Private m_EndColor      As OLE_COLOR
Private m_Style         As ListStyleEnum
Private m_RightLeft     As Boolean

'[Default Property Values]
Private Const m_def_ForeColor = &H80000012
Private Const m_def_SelColor = &HFF8C1A
Private Const m_def_SelForeColor = &HFFFFFF
Private Const m_def_StrechIcon = False
Private Const m_def_Appearance = 1
Private Const m_def_BorderStyle = 1
Private Const m_def_FullRowSel = False
Private Const m_def_SortOrder = 0
Private Const m_def_IconFocus = True
Private Const m_def_TextAllignMent = vbLeftJustify

Private Const m_def_EndColor = &HFFFFFF
Private Const m_def_StartColor = &HFFD6AC
Private Const m_def_Gradient = Horizontal
Private Const m_def_Style = Standard
Private Const m_def_RightLeft = False
Private Const COLOR_MENU As Long = 4
Private Const COLOR_ACTIVECAPTION As Long = 2

'[Events]
Public Event MouseClick()
Public Event SelChange()
Public Event DbClick()
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'-------------------------------------------------------------------------
' Procedure  : AddItem
' Auther     : Jim Jose
' Input      : New item
' OutPut     : None
' Purpose    : To add an item to listBox
'-------------------------------------------------------------------------

Public Sub AddItem(ByVal vText As String, _
                        Optional vIndex As Long = -1, _
                        Optional vKey As String = "Null")
On Error GoTo Handle

    If vIndex = -1 Then
        ' Index not specified , add to last
        If vKey = "Null" Then m_Items.Add vText Else m_Items.Add vText, vKey
    Else
        ' add to specified index
        If vKey = "Null" Then m_Items.Add vText, , vIndex + 1 Else m_Items.Add vText, vKey, vIndex + 1
        
    End If
    
    ' Sort items iff needed
    SortList
    Me.Refresh

Exit Sub
Handle:
    MsgBox Err.Description, vbCritical, "Error, Item could not be added!"
End Sub

'-------------------------------------------------------------------------
' Procedure  : Remove
' Auther     : Jim Jose
' Input      : Index
' OutPut     : None
' Purpose    : To remove an item from List
'-------------------------------------------------------------------------

Public Sub Remove(Optional ByVal vIndex As Long = -1)
    
    If vIndex = -1 Then
        ' Index not specifid, remove selected item
        m_Items.Remove m_SelItem + 1
    Else
        ' Remove specified item
        m_Items.Remove vIndex + 1
    End If
    
    ' Sort It
    SortList
    Me.Refresh
    
End Sub

'-------------------------------------------------------------------------
' Procedure  : Refresh
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Arrage control and calculate local variables
'-------------------------------------------------------------------------

Public Sub Refresh()
On Error Resume Next

    ' Determine item height & item cound per Screen
    Set picDraw.Font = m_Font
    m_iHeight = picDraw.TextHeight("A")
    m_iCount = Int(ScaleHeight / m_iHeight)

    ' Arrange\Set controls
    If m_Items.Count > m_iCount Then
        VScroll.Visible = True
        VScroll.Move ScaleWidth - VScroll.Width, 0, VScroll.Width, ScaleHeight
        VScroll.Max = m_Items.Count - m_iCount
        picDraw.Move 0, 0, ScaleWidth - VScroll.Width, ScaleHeight
    Else
        VScroll.Value = 0
        VScroll.Visible = False
        picDraw.Move 0, 0, ScaleWidth, ScaleHeight
    End If
    
    ' Redraw the list
    ReDrawList
    
End Sub

'-------------------------------------------------------------------------
' Procedure  : ReDrawList
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To draw the entire region
'-------------------------------------------------------------------------

Private Sub ReDrawList()
On Error GoTo Handle
Dim x       As Long
Dim Rct     As RECT
Dim vText   As String
Dim vCount  As Long
Dim vTop    As Long
Dim vIcon   As Boolean
Dim iLeft   As Long
Dim iTop    As Long
Dim vSelCol As Long

    ' Some initial preperation
    CheckSelected

    ' Draw Gradient
    DrawGradient
    
    vCount = m_iTop + m_iCount
    picDraw.ForeColor = m_ForeColor
    Set picDraw.Font = m_Font
    If IsThere(m_Picture) Then picDraw.Picture = m_Picture
    If vCount > m_Items.Count Then vCount = m_Items.Count
    vIcon = IsThere(m_ListIcon)
    
    ' Define space for Listicon\Rect
    If vIcon Then
        Rct.Left = m_iHeight + 3
        If m_StrechIcon Then
            iLeft = 1
            iTop = 1
        Else
            iLeft = m_iHeight / 2 - ScaleX(m_ListIcon.Width) / 2
            iTop = m_iHeight / 2 - ScaleY(m_ListIcon.Height) / 2
        End If
    Else
        Rct.Left = 3
    End If
    Rct.Right = picDraw.Width
    
    
    ' Draw each item
    For x = m_iTop To vCount - 1
        
        ' Downward shift
        Rct.Top = vTop
        Rct.Bottom = Rct.Top + m_iHeight
        
        ' Get the item text
        vText = " " & m_Items(x + 1) & " "
        DrawText picDraw.hdc, vText, -1, Rct, m_TextAllineMent
        
        ' Draw Icons
        If m_StrechIcon Then
            If vIcon Then picDraw.PaintPicture m_ListIcon, iLeft, Rct.Top + iTop, m_iHeight - 1, m_iHeight - 1
        Else
            If vIcon Then picDraw.PaintPicture m_ListIcon, iLeft, Rct.Top + iTop
        End If
        
        ' Downward shift
        vTop = Rct.Bottom
        
    Next x
    
    ' Prepare to draw selection
    x = Rct.Left
    If m_FullRowSel Then Rct.Left = 0
    picDraw.DrawStyle = vbSolid
    picDraw.FillStyle = vbSolid
    
    If m_Items.Count = 0 Then GoTo Handle

    ' Draw the sel back & Focus
    If m_Style = Graphical Then
        vSelCol = m_SelColor
    Else
        vSelCol = GetSysColor(COLOR_ACTIVECAPTION)
    End If
    picDraw.FillColor = vSelCol
    Rct.Top = (m_SelItem - m_iTop) * m_iHeight
    Rct.Bottom = Rct.Top + m_iHeight
    Rectangle picDraw.hdc, Rct.Left, Rct.Top, Rct.Right, Rct.Bottom
    DrawFocusRect picDraw.hdc, Rct
    
    ' Draw iCon on selection
    If m_StrechIcon Then
        If vIcon Then picDraw.PaintPicture m_ListIcon, iLeft, Rct.Top + iTop, m_iHeight - 1, m_iHeight - 1
    Else
        If vIcon Then picDraw.PaintPicture m_ListIcon, iLeft, Rct.Top + iTop
    End If

    ' Draw selected text
    vText = " " & m_Items(m_SelItem + 1) & " "
    picDraw.ForeColor = m_SelForeColor
    Rct.Left = x
    DrawText picDraw.hdc, vText, -1, Rct, m_TextAllineMent
    
    ' Draw Icon Focus
    If vIcon And m_IconFocus Then
        picDraw.ForeColor = m_ForeColor
        Rct.Left = 1
        Rct.Right = Rct.Left + m_iHeight
        Rct.Bottom = Rct.Top + m_iHeight
        DrawFocusRect picDraw.hdc, Rct
    End If
    
Handle:
    ' Refresh Box
    picDraw.Refresh
End Sub

'-------------------------------------------------------------------------
' Procedure  : picDraw_Click
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : RaiseEvent MouseClick
'-------------------------------------------------------------------------

Private Sub picDraw_Click()
    RaiseEvent MouseClick
End Sub

'-------------------------------------------------------------------------
' Procedure  : picDraw_DblClick
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : RaiseEvent DbClick
'-------------------------------------------------------------------------

Private Sub picDraw_DblClick()
    RaiseEvent DbClick
End Sub

'-------------------------------------------------------------------------
' Procedure  : picDraw_KeyDown
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Move Selection by keyboard
'-------------------------------------------------------------------------

Private Sub picDraw_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Select each Key
    Select Case KeyCode
        Case vbKeyUp
            m_SelItem = m_SelItem - 1
        Case vbKeyDown
            m_SelItem = m_SelItem + 1
        Case vbKeyEnd
            m_SelItem = ListCount
        Case vbKeyHome
            m_SelItem = 0
        Case vbKeyPageDown
            m_SelItem = m_SelItem + m_iCount
        Case vbKeyPageUp
            m_SelItem = m_SelItem - m_iCount
    End Select
    
    ' Refrech Control
    Me.Refresh
    RaiseEvent SelChange
    
End Sub

'-------------------------------------------------------------------------
' Procedure  : picDraw_MouseDown
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To calculate selection by mouse
'-------------------------------------------------------------------------

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    ' Calculate row from mouse 'Y'
    m_SelItem = m_iTop + Int(y / m_iHeight)
    CheckSelected
    ReDrawList
    RaiseEvent SelChange
    RaiseEvent MouseDown(Button, Shift, x, y)
    
End Sub

'-------------------------------------------------------------------------
' Procedure  : picDraw_MouseMove
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To RaiseEvent MouseMove
'-------------------------------------------------------------------------

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Calculate row from mouse 'Y'
    If ListCount > m_iCount Then
        picDraw.MousePointer = vbCustom
    Else
        If y < ListCount * m_iHeight Then
            picDraw.MousePointer = vbCustom
        Else
            picDraw.MousePointer = vbNormal
        End If
    End If
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

'-------------------------------------------------------------------------
' Procedure  : picDraw_MouseUp
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To RaiseEvent MouseUp
'-------------------------------------------------------------------------

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'-------------------------------------------------------------------------
' Procedure  : UserControl_Initialize
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Initialise control
'-------------------------------------------------------------------------

Private Sub UserControl_Initialize()
    
    ' Used to prevent crashes on XP
    m_hMode = LoadLibrary("shell32.dll")
    m_KeyControl = True
    
End Sub

'-------------------------------------------------------------------------
' Procedure  : UserControl_InitProperties
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Initialise default property values
'-------------------------------------------------------------------------

Private Sub UserControl_InitProperties()

    m_ForeColor = m_def_ForeColor
    m_SelColor = m_def_SelColor
    m_SelForeColor = m_def_SelForeColor
    Set m_Picture = Nothing
    Set m_ListIcon = Nothing
    Set m_Font = Ambient.Font
    m_StrechIcon = m_def_StrechIcon
    m_iHeight = TextHeight("A")
    m_FullRowSel = m_def_FullRowSel
    m_SortOrder = m_def_SortOrder
    m_IconFocus = m_def_IconFocus
    m_TextAllineMent = m_def_TextAllignMent
    
    m_StartColor = m_def_StartColor
    m_EndColor = m_def_EndColor
    m_Gradient = m_def_Gradient
    m_RightLeft = m_def_RightLeft
    m_Style = m_def_Style
    Me.Refresh
    
End Sub

'-------------------------------------------------------------------------
' Procedure  : UserControl_Resize
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Resize Controls
'-------------------------------------------------------------------------

Private Sub UserControl_Resize()
    Me.Refresh
End Sub

'-------------------------------------------------------------------------
' Procedure  : UserControl_Terminate
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Free the memory
'-------------------------------------------------------------------------

Private Sub UserControl_Terminate()
    FreeLibrary m_hMode
    Me.Clear
    Set m_Items = Nothing
End Sub

'-------------------------------------------------------------------------
' Procedure  : UserControl_WriteProperties
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Write design time propery changes
'-------------------------------------------------------------------------

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("ListIcon", m_ListIcon, Nothing)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("SelColor", m_SelColor, m_def_SelColor)
    Call PropBag.WriteProperty("SelForeColor", m_SelForeColor, m_def_SelForeColor)
    Call PropBag.WriteProperty("StrechIcon", m_StrechIcon, m_def_StrechIcon)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("FullRowSelect", m_FullRowSel, m_def_FullRowSel)
    Call PropBag.WriteProperty("SortOrder", m_SortOrder, m_def_SortOrder)
    Call PropBag.WriteProperty("IconFocus", m_IconFocus, m_def_IconFocus)
    Call PropBag.WriteProperty("TextAlignment", m_TextAllineMent, m_def_TextAllignMent)
    Call PropBag.WriteProperty("StartColor", m_StartColor, m_def_StartColor)
    Call PropBag.WriteProperty("EndColor", m_EndColor, m_def_EndColor)
    Call PropBag.WriteProperty("Gradient", m_Gradient, m_def_Gradient)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)

End Sub

'-------------------------------------------------------------------------
' Procedure  : UserControl_ReadProperties
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Read design time propery changes
'-------------------------------------------------------------------------

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    Set m_ListIcon = PropBag.ReadProperty("ListIcon", Nothing)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_SelColor = PropBag.ReadProperty("SelColor", m_def_SelColor)
    m_SelForeColor = PropBag.ReadProperty("SelForeColor", m_def_SelForeColor)
    m_StrechIcon = PropBag.ReadProperty("StrechIcon", m_def_StrechIcon)
    Me.Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    Me.BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_FullRowSel = PropBag.ReadProperty("FullRowSelect", m_def_FullRowSel)
    m_SortOrder = PropBag.ReadProperty("SortOrder", m_def_SortOrder)
    m_IconFocus = PropBag.ReadProperty("IconFocus", m_def_IconFocus)
    m_TextAllineMent = PropBag.ReadProperty("TextAlignment", m_def_TextAllignMent)
    m_StartColor = PropBag.ReadProperty("StartColor", m_def_StartColor)
    m_EndColor = PropBag.ReadProperty("EndColor", m_def_EndColor)
    m_Gradient = PropBag.ReadProperty("Gradient", m_def_Gradient)
    m_Style = PropBag.ReadProperty("Style", m_def_Style)
    ReDrawList
    
End Sub

'-------------------------------------------------------------------------
' Procedure  : VScroll_Change
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Scroll List
'-------------------------------------------------------------------------

Private Sub VScroll_Change()
    m_iTop = VScroll.Value
    ReDrawList
End Sub

'-------------------------------------------------------------------------
' Procedure  : Clear
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Clear List
'-------------------------------------------------------------------------

Public Sub Clear()
Dim x As Long

    ' Remove each Item
    For x = 1 To m_Items.Count
        m_Items.Remove (1)
    Next x
    
    ' Redraw
    picDraw.Cls
    Me.Refresh
    
End Sub

'-------------------------------------------------------------------------
' Procedure  : ListCount
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : to get ListCount
'-------------------------------------------------------------------------

Public Function ListCount() As Long
On Error GoTo Handle
    ListCount = m_Items.Count
Exit Function
Handle:
    ListCount = 0
End Function

'-------------------------------------------------------------------------
' Procedure  : ListIcon
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property ListIcon
'-------------------------------------------------------------------------

Public Property Get ListIcon() As Picture
    Set ListIcon = m_ListIcon
End Property

Public Property Set ListIcon(ByVal vNewPicture As Picture)
    Set m_ListIcon = vNewPicture
    PropertyChanged "ListIcon"
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : Picture
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property Picture
'-------------------------------------------------------------------------

Public Property Get Picture() As Picture
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal vNewPicture As Picture)
    Set m_Picture = vNewPicture
    PropertyChanged "Picture"
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : Font
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property Font
'-------------------------------------------------------------------------

Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal vNewFont As Font)
    Set m_Font = vNewFont
    PropertyChanged "Font"
    Me.Refresh
End Property

'-------------------------------------------------------------------------
' Procedure  : ForeColor
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property ForeColor
'-------------------------------------------------------------------------

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal vNewCol As OLE_COLOR)
    m_ForeColor = vNewCol
    PropertyChanged "ForeColor"
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : SelColor
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property SelColor
'-------------------------------------------------------------------------

Public Property Get SelColor() As OLE_COLOR
    SelColor = m_SelColor
End Property

Public Property Let SelColor(ByVal vNewCol As OLE_COLOR)
    m_SelColor = vNewCol
    PropertyChanged "SelColor"
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : SelForeColor
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property SelForeColor
'-------------------------------------------------------------------------

Public Property Get SelForeColor() As OLE_COLOR
    SelForeColor = m_SelForeColor
End Property

Public Property Let SelForeColor(ByVal vNewCol As OLE_COLOR)
    m_SelForeColor = vNewCol
    PropertyChanged "SelForeColor"
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : StrechIcon
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property StrechIcon
'-------------------------------------------------------------------------

Public Property Get StrechIcon() As Boolean
    StrechIcon = m_StrechIcon
End Property

Public Property Let StrechIcon(ByVal vNewValue As Boolean)
    m_StrechIcon = vNewValue
    PropertyChanged "StrechIcon"
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : Appearance
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property Appearance
'-------------------------------------------------------------------------

Public Property Get Appearance() As AppearanceEnum
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal vNewAppearance As AppearanceEnum)
    UserControl.Appearance = vNewAppearance
    PropertyChanged "Appearance"
End Property

'-------------------------------------------------------------------------
' Procedure  : BorderStyle
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property BorderStyle
'-------------------------------------------------------------------------

Public Property Get BorderStyle() As BorderEnum
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal vNewBorder As BorderEnum)
    UserControl.BorderStyle = vNewBorder
    PropertyChanged "BorderStyle"
End Property

'-------------------------------------------------------------------------
' Procedure  : ListIndex
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property ListIndex
'-------------------------------------------------------------------------

Public Property Get ListIndex() As Long
    If m_Items.Count = 0 Then
        ListIndex = -1
    Else
        ListIndex = m_SelItem
    End If
End Property

Public Property Let ListIndex(ByVal vNewValue As Long)
    m_SelItem = vNewValue
    PropertyChanged "ListIndex"
    CheckSelected
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : Text
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To get selected text
'-------------------------------------------------------------------------

Public Property Get Text() As String
    If ListCount = 0 Then Exit Property
    Text = m_Items(m_SelItem + 1)
End Property

'-------------------------------------------------------------------------
' Procedure  : ListItems
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Get/Let Property ListItems
'-------------------------------------------------------------------------

Public Property Get ListItems(ByVal vIndex As Long) As String
    ListItems = m_Items(vIndex + 1)
End Property

Public Property Let ListItems(ByVal vIndex As Long, ByVal vNewValue As String)
    m_Items(vIndex + 1) = vNewValue
End Property

'-------------------------------------------------------------------------
' Procedure  : FullRowSelect
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Get/Let Property FullRowSelect
'-------------------------------------------------------------------------

Public Property Get FullRowSelect() As Boolean
    FullRowSelect = m_FullRowSel
End Property

Public Property Let FullRowSelect(ByVal vNewValue As Boolean)
    m_FullRowSel = vNewValue
    PropertyChanged "FullRowSelect"
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : SortOrder
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Get/Let Property SortOrder
'-------------------------------------------------------------------------

Public Property Get SortOrder() As SortOrderEnum
    SortOrder = m_SortOrder
End Property

Public Property Let SortOrder(ByVal vNewValue As SortOrderEnum)
    m_SortOrder = vNewValue
    PropertyChanged "SortOrder"
    SortList
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : SortItems
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Get/Let Property SortItems
'-------------------------------------------------------------------------

Public Property Get IconFocus() As Boolean
    IconFocus = m_IconFocus
End Property

Public Property Let IconFocus(ByVal vNewValue As Boolean)
    m_IconFocus = vNewValue
    PropertyChanged "IconFocus"
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : StartColor
' Auther     : Jim Jose
' Input      : New Value
' OutPut     : Current Value
' Purpose    : To Get/Let Property StartColor
'-------------------------------------------------------------------------

Public Property Get StartColor() As OLE_COLOR
    StartColor = m_StartColor
End Property

Public Property Let StartColor(ByVal vNewValue As OLE_COLOR)
    m_StartColor = vNewValue
    PropertyChanged "StartColor"
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : EndColor
' Auther     : Jim Jose
' Input      : New Value
' OutPut     : Current Value
' Purpose    : To Get/Let Property EndColor
'-------------------------------------------------------------------------

Public Property Get EndColor() As OLE_COLOR
    EndColor = m_EndColor
End Property

Public Property Let EndColor(ByVal vNewValue As OLE_COLOR)
    m_EndColor = vNewValue
    PropertyChanged "EndColor"
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : Gradient
' Auther     : Jim Jose
' Input      : New Value
' OutPut     : Current Value
' Purpose    : To Get/Let Property Gradient Type
'-------------------------------------------------------------------------

Public Property Get Gradient() As GradientTypeConstants
    Gradient = m_Gradient
End Property

Public Property Let Gradient(ByVal vNewValue As GradientTypeConstants)
    m_Gradient = vNewValue
    PropertyChanged "Gradient"
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : TextAlignment
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Get/Let Property TextAlignment
'-------------------------------------------------------------------------

Public Property Get TextAlignment() As AlignmentConstants
    TextAlignment = m_TextAllineMent
End Property

Public Property Let TextAlignment(ByVal vNewValue As AlignmentConstants)
    m_TextAllineMent = vNewValue
    PropertyChanged "TextAlignment"
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : Style
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Property Get/Let  Style
'-------------------------------------------------------------------------

Public Property Get Style() As ListStyleEnum
    Style = m_Style
End Property

Public Property Let Style(ByVal vNewValue As ListStyleEnum)
    m_Style = vNewValue
    PropertyChanged "Style"
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : RightLeft
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Property Get/Let  RightLeft
'-------------------------------------------------------------------------

Public Property Get RightLeft() As Boolean
    RightLeft = m_RightLeft
End Property

Public Property Let RightLeft(ByVal vNewValue As Boolean)
    m_RightLeft = vNewValue
    PropertyChanged "RightLeft"
    ReDrawList
End Property

'-------------------------------------------------------------------------
' Procedure  : IsThere
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To check if the Picture is loaded
'-------------------------------------------------------------------------

Private Function IsThere(vPicture As StdPicture) As Boolean
On Error GoTo Handle
    If Not vPicture.Handle = 0 And Not vPicture.Height = 0 And Not vPicture.Width = 0 Then
        IsThere = True
    Else
        IsThere = False
    End If
Exit Function
Handle:
    IsThere = False
End Function

'-------------------------------------------------------------------------
' Procedure  : VScroll_GotFocus/VScroll_LostFocus
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To determine control using keyboard
'-------------------------------------------------------------------------

Private Sub VScroll_GotFocus()
    m_KeyControl = False
End Sub

Private Sub VScroll_LostFocus()
    m_KeyControl = True
End Sub

'-------------------------------------------------------------------------
' Procedure  : CheckSelected
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To check if the selected item is in permissible range
'              and reset the scroll bars
'-------------------------------------------------------------------------

Private Sub CheckSelected()

    If m_SelItem > m_Items.Count - 1 Then m_SelItem = m_Items.Count - 1
    If m_SelItem < 0 Then m_SelItem = 0
    If m_KeyControl = False Then Exit Sub
    If m_SelItem < m_iTop Then VScroll.Value = m_SelItem
    If m_SelItem > m_iTop + m_iCount - 1 Then VScroll.Value = m_SelItem - m_iCount + 1
    
End Sub

'-------------------------------------------------------------------------
' Procedure  : SortList
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To sort the Data-Collection Ascending/Descending
'-------------------------------------------------------------------------

Private Sub SortList()
Dim x As Long
Dim vPos As Long
Dim vRtn  As Long
Dim vCount As Long
Dim vStart As Long
Dim vNewCount As Long
Dim vNew As New Collection

    ' Check Sort?
    If m_SortOrder = Sort_None Then Exit Sub
    
    ' Get current Count
    vStart = 1
    vCount = m_Items.Count

    ' Loop through Current collection
    For x = vStart To vCount
        
        ' Get new collection count
        vNewCount = vNew.Count
        
        ' Loop through new collection
        For vPos = 1 To vNewCount
        
            ' Compair each item in new collection
            vRtn = StrComp(m_Items(x), vNew(vPos), vbTextCompare)
            ' Escape with purpose
            If vRtn = m_SortOrder Then Exit For
        
        Next vPos
        
        If x = vStart Or vPos = vNewCount + 1 Then
            ' New item at last
            vNew.Add m_Items(x)
        Else
            ' New item somewhere b/w
            vNew.Add m_Items(x), , vPos
        End If
        
    Next x
    
    ' Return Sorted Collection
    Set m_Items = vNew
    
End Sub

'-------------------------------------------------------------------------
' Procedure  : DrawGradient
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Draw DrawGradient ( Horizontal/Vertical)
'-------------------------------------------------------------------------

Private Sub DrawGradient()
On Error GoTo Handle
Dim x       As Long
Dim Col1    As Long
Dim Col2    As Long
Dim tmpCol  As Long
Dim fWidth  As Long
Dim fHeight As Long
Dim R1      As Double, G1       As Double, B1       As Double
Dim R2      As Double, G2       As Double, B2       As Double
Dim Rincr   As Double, Gincr    As Double, Bincr    As Double
    
    ' Clear it
    picDraw.Cls
    
    ' Check Style & Transalate oleColor to long
    If m_Style = Graphical Then ' Graphical style (uses custom colors)
        Col1 = TranslateColor(m_StartColor)
        Col2 = TranslateColor(m_EndColor)
    Else    ' Standard style (use system colors)
        Col1 = TranslateColor(GetSysColor(COLOR_MENU))
        Col2 = vbWhite
    End If
    
    ' Check, Need to apply?
    If m_Gradient = None_Gradient Then picDraw.BackColor = m_StartColor: Exit Sub
        
    ' Check Right to Left
    If Not m_RightLeft Then ' Invert
        tmpCol = Col1
        Col1 = Col2
        Col2 = tmpCol
    End If
    
    ' Get RGB values on start & end colors
    GetRGB Col1, R1, G1, B1
    GetRGB Col2, R2, G2, B2
    fWidth = picDraw.ScaleWidth
    fHeight = picDraw.ScaleHeight
    
    ' Select each case & draw it
    Select Case Gradient
        Case Horizontal
        
            Rincr = (R2 - R1) / fWidth
            Gincr = (G2 - G1) / fWidth
            Bincr = (B2 - B1) / fWidth
            
            For x = 0 To fWidth
                R1 = R1 + Rincr
                G1 = G1 + Gincr
                B1 = B1 + Bincr
                tmpCol = RGB(R1, G1, B1)
                picDraw.ForeColor = tmpCol
                Rectangle picDraw.hdc, x, 0, x + 1, fHeight
            Next x
            
        Case Vertical
        
            Rincr = (R2 - R1) / fHeight
            Gincr = (G2 - G1) / fHeight
            Bincr = (B2 - B1) / fHeight
            
            For x = 0 To fHeight
                R1 = R1 + Rincr
                G1 = G1 + Gincr
                B1 = B1 + Bincr
                tmpCol = RGB(R1, G1, B1)
                picDraw.ForeColor = tmpCol
                Rectangle picDraw.hdc, 0, x, fWidth, x + 1
            Next x
            
    End Select
    
Handle:
End Sub

'-------------------------------------------------------------------------
' Procedure  : TranslateColor
' Auther     : Charls P.v
' Input      : OleColor
' OutPut     : Long Color
' Purpose    : Convert OleColor to Long Color
'-------------------------------------------------------------------------

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If
End Function

'-------------------------------------------------------------------------
' Procedure  : GetRGB
' Auther     : Jim Jose
' Input      : Long color
' OutPut     : RGB colors
' Purpose    : Long Color to RGB
'-------------------------------------------------------------------------

Private Function GetRGB(ByVal LngCol As Long, R As Double, G As Double, B As Double)
    R = LngCol Mod 256
    G = (LngCol And vbGreen) / 256 'Green
    B = (LngCol And vbBlue) / 65536 'Blue
End Function


Public Property Get Hwnd() As Variant
    Hwnd = UserControl.Hwnd
End Property

