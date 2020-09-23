Attribute VB_Name = "mdlExecute"
Option Explicit

'[ APIs ]
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

'[ This function can Launch the input file ]
'-------------------------------------------
Public Function LaunchBrowser(Frm As Form, _
                            ByVal WebURL As String, _
                            ByVal NewWindow As Boolean, _
                            ByVal IsFolder As Boolean)
Dim RetVal As Long
Dim TempPath As String
Dim Executer As String
    
    ' Get temp directory
    Executer = Space(255)
    
    If IsFolder Then

        ' Find the executable
        RetVal = FindExecutable(App.Path, "", Executer)
        Executer = Split(Executer, "")(0)
        
    Else
    
        TempPath = GetTempDir & "Temp.html"
        
        ' Create a temp html file
        Open TempPath For Binary As #1
        Close #1
        
        ' Find the executable
        RetVal = FindExecutable(TempPath, "", Executer)
        Executer = Split(Executer, "")(0)
        
        ' Kill the temp html
        Kill TempPath
        
    End If
    
    ' Execute our Address on the executable
    If RetVal <= 32 Or IsEmpty(Executer) Then
        MsgBox "Could Not find associated Browser", vbExclamation, "Browser Not Found"
        Exit Function
    Else
        If NewWindow Then
            RetVal = ShellExecute(Frm.hwnd, "open", Executer, WebURL, vbNullString, 1)
        Else
            Call ShellExecute(Frm.hwnd, "open", WebURL, vbNullString, 1, 1)
        End If
    End If
    
End Function

' [ Getting the Temporary Directory ]
'------------------------------------
Public Function GetTempDir() As String
Dim RetVal As Long
Dim StrTemp As String
    StrTemp = Space$(1024)
    RetVal = GetTempPath(Len(StrTemp), StrTemp)
    GetTempDir = Left$(StrTemp, RetVal)
End Function

