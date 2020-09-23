Attribute VB_Name = "mdlTooltip"

'-------------------------------------------------------------------------------------------------------------------------
' Procedure : ShowTooltip
' Auther    : Jim Jose
' BasicCode : Fred.cpp
' Upgraded  : Dana Seaman, for unicode support
' Purpose   : Simple and efficient tooltip generation with baloon style
'-------------------------------------------------------------------------------------------------------------------------

Option Explicit

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public Enum ToolTipStyleEnum
    [Tip_Standard] = 0
    [Tip_Balloon] = 1
End Enum

Public Enum ToolTipTypeEnum
    [Typ_None] = 0
    [Typ_Info] = 1
    [Typ_Warning] = 1
    [Typ_Error] = 1
End Enum

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   bottom As Long
End Type

Private Type TOOLINFO
    lSize           As Long
    lFlags          As Long
    lHwnd           As Long
    lId             As Long
    lpRect          As RECT
    hInstance       As Long
    lpStr           As Long
    lParam          As Long
End Type

''Tooltip Window Constants
Private Const WM_USER                   As Long = &H400
Private Const TTS_NOPREFIX              As Long = &H2
Private Const TTF_TRANSPARENT           As Long = &H100
Private Const TTF_CENTERTIP             As Long = &H2
Private Const TTM_ADDTOOLA              As Long = (WM_USER + 4)
Private Const TTM_ADDTOOLW              As Long = (WM_USER + 50)
Private Const TTM_DELTOOLA              As Long = (WM_USER + 5)
Private Const TTM_DELTOOLW              As Long = (WM_USER + 51)
Private Const TTM_ACTIVATE              As Long = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA        As Long = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH        As Long = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR         As Long = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR       As Long = (WM_USER + 20)
Private Const TTM_SETTITLE              As Long = (WM_USER + 32)
Private Const TTM_SETTITLEW             As Long = (WM_USER + 33)
Private Const TTS_BALLOON               As Long = &H40
Private Const TTS_ALWAYSTIP             As Long = &H1
Private Const TTF_SUBCLASS              As Long = &H10
Private Const TOOLTIPS_CLASSA           As String = "tooltips_class32"
Private Const CW_USEDEFAULT             As Long = &H80000000
Private Const TTM_SETMARGIN             As Long = (WM_USER + 26)

Private Const SWP_FRAMECHANGED          As Long = &H20
Private Const SWP_DRAWFRAME             As Long = SWP_FRAMECHANGED
Private Const SWP_HIDEWINDOW            As Long = &H80
Private Const SWP_NOACTIVATE            As Long = &H10
Private Const SWP_NOCOPYBITS            As Long = &H100
Private Const SWP_NOMOVE                As Long = &H2
Private Const SWP_NOOWNERZORDER         As Long = &H200
Private Const SWP_NOREDRAW              As Long = &H8
Private Const SWP_NOREPOSITION          As Long = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE                As Long = &H1
Private Const SWP_NOZORDER              As Long = &H4
Private Const HWND_TOPMOST              As Long = -&H1

Private m_ToolTipHwnd   As Long
Private m_ToolTipInfo   As TOOLINFO
Private m_PrevMousePos  As POINTAPI

Public Sub ShowTooltip(ByVal hwnd As Long, _
                        ByVal mToolTipHead As String, _
                        ByVal mToolTipText As String, _
                        ByVal mToolTipStyle As ToolTipStyleEnum, _
                        ByVal mToolTipType As ToolTipTypeEnum)
Dim lpRect As RECT
Dim lWinStyle As Long
Dim lMousePos As POINTAPI

    
    ' Get the position
    GetCursorPos lMousePos
    If lMousePos.X = m_PrevMousePos.X And lMousePos.Y = m_PrevMousePos.Y Then Exit Sub
    
    'Remove previous ToolTip
    RemoveToolTip
    If mToolTipText = vbNullString And mToolTipHead = vbNullString Then Exit Sub

    ''create baloon style if desired
    If mToolTipStyle = Tip_Standard Then
        lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    Else
        lWinStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX Or TTS_BALLOON
    End If
        
    m_ToolTipHwnd = CreateWindowEx(0&, _
                TOOLTIPS_CLASSA, _
                vbNullString, _
                lWinStyle, _
                CW_USEDEFAULT, _
                CW_USEDEFAULT, _
                CW_USEDEFAULT, _
                CW_USEDEFAULT, _
                hwnd, _
                0&, _
                App.hInstance, _
                0&)
                   
    'make our tooltip window a topmost window
    SetWindowPos m_ToolTipHwnd, _
        HWND_TOPMOST, _
        0&, _
        0&, _
        0&, _
        0&, _
        SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE


    ''get the rect of the parent control
    GetClientRect hwnd, lpRect

    ''now set our tooltip info structure
    With m_ToolTipInfo
        .lSize = Len(m_ToolTipInfo)
        .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
        .lHwnd = hwnd
        .lId = 0
        .hInstance = App.hInstance
        .lpStr = StrPtr(mToolTipText)
        .lpRect = lpRect
    End With

    ''add the tooltip structure
    SendMessage m_ToolTipHwnd, TTM_ADDTOOLW, 0&, m_ToolTipInfo

    ''if we want a title or we want an icon
    SendMessage m_ToolTipHwnd, TTM_SETTIPTEXTCOLOR, vbRed, 0&
    SendMessage m_ToolTipHwnd, TTM_SETTIPBKCOLOR, vbBlue, 0&
    SendMessage m_ToolTipHwnd, TTM_SETTITLEW, 1&, ByVal StrPtr(mToolTipHead)

    Debug.Print "Show tip " & m_ToolTipHwnd

m_PrevMousePos.X = lMousePos.X
m_PrevMousePos.Y = lMousePos.Y

Exit Sub
ErrHandler:
   Debug.Print "Error " & Err.Description
End Sub


'[Important. If not included, tooltips don't change when you try to set the toltip text]
Private Sub RemoveToolTip()
   Dim lR As Long
   Debug.Print "Removed " & m_ToolTipHwnd
   If m_ToolTipHwnd <> 0 Then
      lR = SendMessage(m_ToolTipInfo.lHwnd, TTM_DELTOOLW, 0, m_ToolTipInfo)
      DestroyWindow m_ToolTipHwnd
      m_ToolTipHwnd = 0
   End If
End Sub
