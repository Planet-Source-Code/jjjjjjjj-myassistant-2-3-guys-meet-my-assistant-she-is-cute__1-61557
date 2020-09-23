Attribute VB_Name = "mdlTray"
Option Explicit

'[Tray Constants]
Const NIF_MESSAGE    As Long = &H1     'Message
Const NIF_ICON       As Long = &H2     'Icon
Const NIF_TIP        As Long = &H4     'TooTipText
Const NIM_ADD        As Long = &H0     'Add to tray
Const NIM_MODIFY     As Long = &H1     'Modify
Const NIM_DELETE     As Long = &H2     'Delete From Tray

'[Type NotifyIconData For Tray Icon]
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'[Return Events]
Public Enum TrayRetunEventEnum
    MouseMove = &H200       'On Mousemove
    LeftUp = &H202          'Left Button Mouse Up
    LeftDown = &H201        'Left Button MouseDown
    LeftDbClick = &H203     'Left Button Double Click
    RightUp = &H205         'Right Button Up
    RightDown = &H204       'Right Button Down
    RightDbClick = &H206    'Right Button Double Click
    MiddleUp = &H208        'Middle Button Up
    MiddleDown = &H207      'Middle Button Down
    MiddleDbClick = &H209   'Middle Button Double Click
End Enum

'[Modify Items]
Public Enum ModifyItemEnum
    ToolTip = 1             'Modify ToolTip
    Icon = 2                'Modify Icon
End Enum

'[API]
Private TrayIcon As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'[Add to Tray]
Public Sub TrayAdd(hwnd As Long, Icon As Picture, _
                    ToolTip As String, ReturnCallEvent As TrayRetunEventEnum)
    With TrayIcon
        .cbSize = Len(TrayIcon)
        .hwnd = hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = ReturnCallEvent
        .hIcon = Icon
        .szTip = ToolTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, TrayIcon
End Sub

'[Remove From tray]
Public Sub TrayDelete()
    Shell_NotifyIcon NIM_DELETE, TrayIcon
End Sub

'[Modify the tray]
Public Sub TrayModify(Item As ModifyItemEnum, vNewValue As Variant)
    Select Case Item
        Case ToolTip
            TrayIcon.szTip = vNewValue & vbNullChar
        Case Icon
            TrayIcon.hIcon = vNewValue
    End Select
    Shell_NotifyIcon NIM_MODIFY, TrayIcon
End Sub


