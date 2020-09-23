VERSION 5.00
Begin VB.Form frmNew 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0005BAFE&
   BorderStyle     =   0  'None
   Caption         =   "New Address"
   ClientHeight    =   1530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   102
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAddress 
      BackColor       =   &H00A3E8FC&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox txtTitle 
      BackColor       =   &H00A3E8FC&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin MyAssistant.ColorButton clbAdd 
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   840
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Caption         =   "Add"
      BackColor       =   33023
      BorderColor     =   14737632
      ForeColor       =   0
   End
   Begin MyAssistant.ColorButton clbClose 
      Height          =   225
      Left            =   5520
      TabIndex        =   6
      Top             =   15
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   397
      Caption         =   "X"
      BackColor       =   33023
      BorderColor     =   33023
      ForeColor       =   0
   End
   Begin VB.Label lbTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Add New WebSite"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   360
      Width           =   1800
   End
   Begin VB.Label lbAddress 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub clbAdd_Click()
    If clbAdd.Caption = "Add" Then
        frmMain.lstTitles.AddItem txtTitle, -1
        frmMain.lstAddress.AddItem txtTitle & " <Address>" & txtAddress, -1
    Else
        frmMain.lstTitles.Remove
        frmMain.lstAddress.Remove frmMain.lstTitles.ListIndex
        frmMain.lstTitles.AddItem txtTitle, -1
        frmMain.lstAddress.AddItem txtTitle & " <Address>" & txtAddress, -1
    End If
    Unload Me
    
End Sub

Private Sub clbClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Visible = True
    AnimateForm Me.Hwnd, aload, eStrechHorizontally, 1
    DrawBorder Me
    SetWindowPos Hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS

End Sub

Private Sub Form_Unload(Cancel As Integer)
    AnimateForm Me.Hwnd, aUnload, eStrechHorizontally, 1
End Sub

'[Address AutoFill]
Private Sub txtAddress_LostFocus()
Dim Pos As Integer
    If Not GlbMode = [Web Assistant] Then Exit Sub
    If txtAddress = vbNullString Then Exit Sub
    txtAddress = RTrim$(LTrim$(txtAddress))
    Pos = InStr(1, txtAddress, "www.", vbTextCompare)
    If Pos = 0 Then txtAddress = "www." & txtAddress
    Pos = InStr(1, txtAddress, "http://", vbTextCompare)
    If Pos = 0 Then txtAddress = "http://" & txtAddress
End Sub

'[Add new address]
Public Sub NewAddress()
    If GlbMode = [Folder Assistant] Then
        lbTitle = "Add New Folder"
        lbAddress = "Path"
    ElseIf GlbMode = [Web Assistant] Then
        lbTitle = "Add New WebSite"
        lbAddress = "Address"
    End If
    clbAdd.ToolTipText = "Add"
    clbAdd.Caption = "Add"
    
End Sub

'[Edit current entry]
Public Sub EditAddress()
    If GlbMode = [Folder Assistant] Then
        lbTitle = "Edit Path"
        lbAddress = "Path"
    ElseIf GlbMode = [Web Assistant] Then
        lbTitle = "Edit Address"
        lbAddress = "Address"
    End If

    clbAdd.ToolTipText = "Save"
    clbAdd.Caption = "Save"
    txtAddress = Split(frmMain.lstAddress.ListItems(frmMain.lstTitles.ListIndex), " <Address>")(1)
    txtTitle = frmMain.lstTitles.Text

End Sub
