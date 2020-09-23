Attribute VB_Name = "mdlGlobal"
Option Explicit

Public Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
            
Public Enum GlbModeEnum
    [Search Assistant] = 0
    [Web Assistant] = 1
    [Folder Assistant] = 2
End Enum

Public GlbMode As GlbModeEnum
Public GlbMinimized As Boolean

Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_NOTOPMOST = -2

'[Save information to text file]
Public Sub SaveList()
Dim x As Long
Dim TxtData() As String
ReDim TxtData(frmMain.lstAddress.ListCount) As String

    For x = 0 To frmMain.lstAddress.ListCount - 1
        TxtData(x) = frmMain.lstAddress.ListItems(x)
    Next x
    
    frmMain.lstAddress.Clear
    frmMain.lstTitles.Clear
    
    If GlbMode = [Folder Assistant] Then
        Kill App.Path & "\Bin\FolderList.txt"
        SaveFile Join(TxtData, "<Split>"), App.Path & "\Bin\FolderList.txt"
    ElseIf GlbMode = [Web Assistant] Then
        Kill App.Path & "\Bin\Websites.txt"
        SaveFile Join(TxtData, "<Split>"), App.Path & "\Bin\Websites.txt"
    End If
    
End Sub


'[Reading information from text file]
Public Sub LoadList()
Dim TxtSource As String
Dim SplitData() As String
Dim x As Integer, y As Integer

    ' Load the data
    If GlbMode = [Web Assistant] Then
        TxtSource = LoadFile(App.Path & "\Bin\WebSites.txt")
    ElseIf GlbMode = [Folder Assistant] Then
        TxtSource = LoadFile(App.Path & "\Bin\FolderList.txt")
    End If
    
    frmMain.lstAddress.Clear
    frmMain.lstTitles.Clear
    
    If TxtSource = "" Then Exit Sub
    SplitData = Split(TxtSource, "<Split>")
    
    For x = LBound(SplitData) To UBound(SplitData) - 1
        frmMain.lstAddress.AddItem SplitData(x), -1
        frmMain.lstTitles.AddItem Split(SplitData(x), " <Address>")(0), -1
    Next x
    
    Erase SplitData
    
End Sub

'[Load the user settings]
Public Sub LoadSettingsEX()
Dim TxtData As String
Dim TxtSplit() As String

    On Error GoTo Handle
    TxtData = LoadFile(App.Path & "\Bin\Settings.txt")
    TxtSplit = Split(TxtData, "<Split>")
    
    frmMain.chkAnimate.Value = Val(TxtSplit(0))
    GlbMode = Val(TxtSplit(1))
    frmMain.chkOpen.Value = Val(TxtSplit(2))
    frmMain.chkStartUp.Value = Val(TxtSplit(3))
    frmMain.txtCount = Val(TxtSplit(4))
    frmMain.chkExact = Val(TxtSplit(5))
Handle:
End Sub

'[Save the user settings]
Public Sub SaveSettingsEX()
Dim TxtData(5) As String

    Kill App.Path & "\Bin\Settings.txt"
    
    TxtData(0) = frmMain.chkAnimate.Value
    TxtData(1) = GlbMode
    TxtData(2) = frmMain.chkOpen.Value
    TxtData(3) = frmMain.chkStartUp.Value
    TxtData(4) = Val(frmMain.txtCount)
    TxtData(5) = frmMain.chkExact
    SaveFile Join(TxtData, "<Split>"), App.Path & "\Bin\Settings.txt"
    
End Sub

Public Sub BringToTop(mObject As Object)
    mObject.ZOrder (0)
    AnimateForm mObject.Hwnd, aload, eZoomOut, 1
End Sub
