VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0038BCFF&
   BorderStyle     =   0  'None
   Caption         =   "My Assistant"
   ClientHeight    =   7965
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   10740
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   531
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   716
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picModes 
      BackColor       =   &H00A3E8FC&
      Height          =   3405
      Left            =   2880
      ScaleHeight     =   3345
      ScaleWidth      =   2355
      TabIndex        =   4
      Top             =   720
      Width           =   2415
      Begin MyAssistant.ColorButton clbSearch 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Search Web"
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Picture         =   "frmMain.frx":08CA
         BackColor       =   33023
         BorderColor     =   16777215
         ForeColor       =   0
      End
      Begin MyAssistant.ColorButton clbWAssistant 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Web List"
         Top             =   840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Picture         =   "frmMain.frx":0C64
         BackColor       =   33023
         BorderColor     =   16777088
         ForeColor       =   0
      End
      Begin MyAssistant.ColorButton clbFolderAssis 
         Height          =   375
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Folder List"
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Picture         =   "frmMain.frx":11FE
         BackColor       =   33023
         BorderColor     =   16777088
         ForeColor       =   0
      End
      Begin MyAssistant.ColorButton clbOptions 
         Height          =   375
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "Options"
         Top             =   2280
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Picture         =   "frmMain.frx":1798
         BackColor       =   33023
         BorderColor     =   14737632
         ForeColor       =   0
      End
      Begin MyAssistant.ColorButton clbAbout 
         Height          =   375
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "About"
         Top             =   2760
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ""
         Picture         =   "frmMain.frx":1B32
         BackColor       =   33023
         BorderColor     =   14737632
         ForeColor       =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0097D1FD&
         X1              =   0
         X2              =   2400
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  About 'My Assistant'"
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
         MouseIcon       =   "frmMain.frx":20CC
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   2880
         Width           =   1830
      End
      Begin VB.Label lbOptions 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  Options"
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
         MouseIcon       =   "frmMain.frx":23D6
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   2400
         Width           =   765
      End
      Begin VB.Label lbSearch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  Search Assistant"
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
         MouseIcon       =   "frmMain.frx":26E0
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   480
         Width           =   1545
      End
      Begin VB.Label lbWeb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  Web Assistant"
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
         MouseIcon       =   "frmMain.frx":29EA
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbFolder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  Folder Assistant"
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
         MouseIcon       =   "frmMain.frx":2CF4
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   1440
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H001A9BFB&
         Caption         =   "Plug Modes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2385
      End
   End
   Begin VB.PictureBox picAbout 
      BackColor       =   &H00A3E8FC&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3405
      Left            =   8160
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   30
      Top             =   4320
      Width           =   2415
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright.  Jim Jose 2005"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         MouseIcon       =   "frmMain.frx":2FFE
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   3000
         Width           =   1860
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "          This version also have Google search mode which can be used for Advanced and quick search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   45
         MouseIcon       =   "frmMain.frx":3308
         MousePointer    =   99  'Custom
         TabIndex        =   38
         Top             =   2160
         Width           =   2370
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":3612
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2355
         Left            =   45
         MouseIcon       =   "frmMain.frx":3717
         MousePointer    =   99  'Custom
         TabIndex        =   37
         Top             =   360
         Width           =   2250
      End
      Begin VB.Label lbAbout 
         Alignment       =   2  'Center
         BackColor       =   &H001A9BFB&
         Caption         =   "About "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   2385
      End
   End
   Begin VB.PictureBox picOptions 
      BackColor       =   &H00A3E8FC&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3405
      Left            =   8160
      ScaleHeight     =   3345
      ScaleWidth      =   2355
      TabIndex        =   27
      Top             =   720
      Width           =   2415
      Begin VB.CheckBox chkStartUp 
         BackColor       =   &H00A3E8FC&
         Caption         =   "Load On StartUp"
         Height          =   240
         Left            =   240
         TabIndex        =   40
         Top             =   840
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkAnimate 
         BackColor       =   &H00A3E8FC&
         Caption         =   "Animate Form"
         Height          =   240
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H001A9BFB&
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   2385
      End
   End
   Begin VB.CheckBox chkOpen 
      BackColor       =   &H0038BCFF&
      Caption         =   "&Open in a new window"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   4155
      Width           =   2175
   End
   Begin MyAssistant.ListBoxEX lstTitles 
      Height          =   3405
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "Address List"
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   6006
      ListIcon        =   "frmMain.frx":3A21
      Picture         =   "frmMain.frx":3A3D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SortOrder       =   -1
      StartColor      =   10742012
      EndColor        =   375550
      Gradient        =   0
      Style           =   1
   End
   Begin MyAssistant.ColorButton clbAdd 
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      ToolTipText     =   "Add"
      Top             =   4440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      Caption         =   "Add"
      Picture         =   "frmMain.frx":3A59
      BackColor       =   33023
      BorderColor     =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin MyAssistant.ColorButton clbDelete 
      Height          =   615
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Delete"
      Top             =   4440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      Caption         =   "Delete"
      Picture         =   "frmMain.frx":3DF3
      BackColor       =   33023
      BorderColor     =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin MyAssistant.ColorButton cmdEdit 
      Height          =   615
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Edit"
      Top             =   4440
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      Caption         =   "Edit"
      Picture         =   "frmMain.frx":438D
      BackColor       =   33023
      BorderColor     =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin MyAssistant.ColorButton clbClose 
      Height          =   225
      Left            =   2520
      TabIndex        =   14
      ToolTipText     =   "Close"
      Top             =   15
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   397
      Caption         =   "X"
      BackColor       =   33023
      BorderColor     =   33023
      ForeColor       =   0
   End
   Begin MyAssistant.ColorButton clbMinimize 
      Height          =   225
      Left            =   2280
      TabIndex        =   15
      ToolTipText     =   "Minimize to Tray"
      Top             =   15
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   397
      Caption         =   "-"
      BackColor       =   33023
      BorderColor     =   33023
      ForeColor       =   0
   End
   Begin MyAssistant.ListBoxEX lstAddress 
      Height          =   3405
      Left            =   240
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   6006
      ListIcon        =   "frmMain.frx":4727
      Picture         =   "frmMain.frx":4743
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SortOrder       =   -1
      StartColor      =   10742012
      EndColor        =   375550
      Gradient        =   0
      Style           =   1
   End
   Begin VB.PictureBox picSearch 
      BackColor       =   &H00A3E8FC&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3405
      Left            =   5520
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   17
      Top             =   720
      Width           =   2415
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1800
      End
      Begin VB.TextBox txtCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00A3E8FC&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   41
         Text            =   "100"
         Top             =   1620
         Width           =   360
      End
      Begin VB.OptionButton optImage 
         BackColor       =   &H00A3E8FC&
         Caption         =   "Image"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optGroup 
         BackColor       =   &H00A3E8FC&
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton optNews 
         BackColor       =   &H00A3E8FC&
         Caption         =   "News"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   23
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton optWeb 
         BackColor       =   &H00A3E8FC&
         Caption         =   "Web"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin MyAssistant.ColorButton clbFind 
         Height          =   285
         Left            =   1920
         TabIndex        =   20
         ToolTipText     =   "Search On Web"
         Top             =   1200
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   503
         Caption         =   "Go"
         BackColor       =   33023
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
      End
      Begin VB.CheckBox chkExact 
         BackColor       =   &H00A3E8FC&
         Caption         =   "Exact phrase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label lbAdvanced 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advanced Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   120
         MouseIcon       =   "frmMain.frx":475F
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   1485
         Width           =   1260
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Count "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   42
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search For"
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
         Left            =   120
         MouseIcon       =   "frmMain.frx":4A69
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H001A9BFB&
         Caption         =   "Google Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   2385
      End
   End
   Begin MyAssistant.ColorButton clbPlug 
      Height          =   375
      Left            =   2280
      TabIndex        =   36
      ToolTipText     =   "Plug Modes"
      Top             =   330
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ""
      Picture         =   "frmMain.frx":4D73
      BackColor       =   33023
      BorderColor     =   14737632
      ForeColor       =   0
   End
   Begin VB.Image imgTitle 
      Height          =   255
      Left            =   240
      Top             =   360
      Width           =   255
   End
   Begin VB.Label lbTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "My Assistant"
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
      Left            =   600
      TabIndex        =   7
      Top             =   360
      Width           =   1230
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub clbAbout_Click()
    BringToTop picAbout
End Sub

Private Sub clbFind_Click()
Dim GoogleString As String
    
    If optWeb Then
        GoogleString = GetGoogleSearchString(txtFind, Val(txtCount), Srch_Web, chkExact)
    
    ElseIf optImage Then
        GoogleString = GetGoogleSearchString(txtFind, Val(txtCount), Srch_Image, chkExact)
    
    ElseIf optNews Then
        GoogleString = GetGoogleSearchString(txtFind, Val(txtCount), Srch_News, chkExact)
    
    ElseIf optGroup Then
        GoogleString = GetGoogleSearchString(txtFind, Val(txtCount), Srch_Group, chkExact)
    
    End If
    
    LaunchBrowser Me, GoogleString, chkOpen, False
    
End Sub

Private Sub clbFolderAssis_Click()
    If GlbMode = [Web Assistant] Then SaveList
    GlbMode = [Folder Assistant]
    BringToTop lstTitles
    LoadList
    SetMode
End Sub

Private Sub clbMinimize_Click()
    GlbMinimized = True
    AnimateForm Me.Hwnd, aUnload, eAppearFromTop, 1
End Sub

Private Sub clbPlug_Click()
    BringToTop picModes
    imgTitle.Picture = clbAbout.Picture
    lbTitle = "My Assistant"
End Sub

Private Sub clbSearch_Click()
    SaveList
    GlbMode = [Search Assistant]
    BringToTop picSearch
    SetMode
End Sub

Private Sub Form_Load()

    picOptions.Move lstTitles.Left, lstTitles.Top, lstTitles.Width, lstTitles.Height
    picAbout.Move lstTitles.Left, lstTitles.Top, lstTitles.Width, lstTitles.Height
    picModes.Move lstTitles.Left, lstTitles.Top, lstTitles.Width, lstTitles.Height
    picSearch.Move lstTitles.Left, lstTitles.Top, lstTitles.Width, lstTitles.Height
    picSearch.Move lstTitles.Left, lstTitles.Top, lstTitles.Width, lstTitles.Height

    Me.Width = 2925
    Me.Height = 5310
    Me.Move Screen.Width - Me.Width, Screen.Height - Me.Height - 500
    Me.Visible = True
    AnimateForm Me.Hwnd, aload, eStrechVertically, 1

    LoadSettingsEX
    SetMode
    LoadList
    DrawBorder Me
    TrayAdd Hwnd, Me.Icon, "My Assistant", MouseMove

End Sub

Private Sub clbAdd_Click()
    If GlbMode = [Search Assistant] Then Exit Sub
    frmNew.NewAddress
End Sub

Private Sub clbClose_Click()
    AnimateForm Me.Hwnd, aUnload, eAppearFromLeft, 1
    Unload Me
End Sub

Private Sub clbDelete_Click()
    If GlbMode = [Search Assistant] Then Exit Sub
    If lstTitles.Text = "" Then Exit Sub
    lstTitles.Remove
    lstAddress.Remove lstTitles.ListIndex
End Sub

Private Sub clbOptions_Click()
    SetMode
    BringToTop picOptions
End Sub

Private Sub clbWAssistant_Click()
    If GlbMode = [Folder Assistant] Then SaveList
    GlbMode = [Web Assistant]
    BringToTop lstTitles
    LoadList
    SetMode
End Sub

Private Sub cmdEdit_Click()
    If GlbMode = [Search Assistant] Then Exit Sub
    If lstTitles.Text = "" Then Exit Sub
    frmNew.EditAddress
End Sub

Private Sub lbClose_Click()
    Unload Me
End Sub

Private Sub lbMinimise_Click()
    Me.WindowState = 1
End Sub

'[Mouse Move]- Caching the events on tray icon
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim cEvent As Single
cEvent = x

Select Case cEvent
    Case LeftUp
        SetWindowPos Hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
        SetWindowPos Hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS
        If GlbMinimized Then
            GlbMinimized = False
            AnimateForm Me.Hwnd, aload, eAppearFromBottom, 1
            DrawBorder Me
        End If
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not GlbMode = [Search Assistant] Then SaveList
    SaveSettingsEX
    CheckStartUp
    TrayDelete
    End
End Sub

Private Sub LblRestore_Click()
    TrayAdd Me.Hwnd, Me.Icon, "My Assistant", MouseMove
    AnimateForm Me.Hwnd, aUnload, eAppearFromTop
    Me.Visible = False
End Sub


Public Sub SetMode()
Select Case GlbMode
    Case [Search Assistant]
        lbSearch.FontBold = True
        lbWeb.FontBold = False
        lbFolder.FontBold = False
        lbTitle = "Search Assistant"
        imgTitle.Picture = clbSearch.Picture
        Set lstTitles.ListIcon = clbSearch.Picture
        picSearch.ZOrder (0)
        
    Case [Web Assistant]
        lbWeb.FontBold = True
        lbFolder.FontBold = False
        lbSearch.FontBold = False
        lbTitle = "Web Assistant"
        Set lstTitles.ListIcon = clbWAssistant.Picture
        imgTitle.Picture = clbWAssistant.Picture
        Set lstTitles.ListIcon = clbWAssistant.Picture
        lstTitles.ZOrder (0)
        
    Case [Folder Assistant]
        lbFolder.FontBold = True
        lbWeb.FontBold = False
        lbSearch.FontBold = False
        lbTitle = "Folder Assistant"
        imgTitle.Picture = clbFolderAssis.Picture
        Set lstTitles.ListIcon = clbFolderAssis.Picture
        lstTitles.ZOrder (0)
    
End Select

End Sub

Private Sub Label5_Click()
    clbAbout_Click
End Sub

Private Sub lbAdvanced_Click()
Dim GoogleString As String
    GoogleString = GetGoolgeAdvancedSearchPage(txtFind)
    LaunchBrowser Me, GoogleString, chkOpen, False
End Sub

Private Sub lbFolder_Click()
    clbFolderAssis_Click
End Sub

Private Sub lbOptions_Click()
    clbOptions_Click
End Sub

Private Sub lbSearch_Click()
    clbSearch_Click
End Sub

Private Sub lbWeb_Click()
    clbWAssistant_Click
End Sub

Private Sub lstTitles_DbClick()
    If GlbMode = [Folder Assistant] Then
        LaunchBrowser Me, Split(lstAddress.ListItems(lstTitles.ListIndex), " <Address>")(1), chkOpen, True
    Else
        LaunchBrowser Me, Split(lstAddress.ListItems(lstTitles.ListIndex), " <Address>")(1), chkOpen, False
    End If
End Sub

Private Sub lstTitles_MouseClick()
    lstTitles.ToolTipText = Split(lstAddress.ListItems(lstTitles.ListIndex), " <Address>")(1)
End Sub

Public Sub CheckStartUp()
If chkStartUp.Value = 1 Then
     AddToStartUp App.EXEName, App.Path & "\" & App.EXEName & ".exe", True
Else
     AddToStartUp App.EXEName, App.Path & "\" & App.EXEName & ".exe", False
End If
End Sub
