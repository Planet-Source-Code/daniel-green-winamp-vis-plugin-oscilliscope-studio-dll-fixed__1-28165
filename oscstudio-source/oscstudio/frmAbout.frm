VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A b o u t  O s c i l l i s c o p e  * S t u d i o *"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.Label lblGreenAudio 
         Alignment       =   2  'Center
         Caption         =   "Oscilliscope *Studio* >>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   2175
      End
      Begin VB.Line Line1 
         X1              =   2400
         X2              =   2400
         Y1              =   120
         Y2              =   360
      End
      Begin VB.Label lblCredits 
         Alignment       =   2  'Center
         Caption         =   "Credits"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
      Begin VB.Line Line2 
         X1              =   4080
         X2              =   4080
         Y1              =   120
         Y2              =   360
      End
      Begin VB.Label lblHistory 
         Alignment       =   2  'Center
         Caption         =   "Version History"
         Height          =   255
         Left            =   4200
         TabIndex        =   1
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   360
      Width           =   6735
      Begin VB.Frame Frame2 
         Caption         =   "Credits"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   6615
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Various other features, including VisRelay - the open source community."
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   6375
         End
         Begin VB.Label lblLiquid 
            Caption         =   "System Code, ""effects"", ideas? - Dan Green (dan.green@morphedmedia.com)"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   6375
         End
      End
      Begin VB.Label lblIntro 
         Caption         =   "Copyright (C) 2001"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblVer 
         BackStyle       =   0  'Transparent
         Caption         =   "Oscilliscope *Studio*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Oscilliscope *Studio*:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "A compilation of open-source and closed source code."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label Label3 
         Caption         =   " MorphedMedia"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3360
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblVersion 
         Caption         =   " Version:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3480
         TabIndex        =   9
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   6735
      Begin VB.TextBox txtHistory 
         BackColor       =   &H80000000&
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   120
         Width           =   6495
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const pi = 3.14159
Dim isdone As Boolean
Private Sub cmdClose_Click()
isdone = True
Unload frmAbout
End Sub

Private Sub Form_Load()
lblVersion = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
frmAbout.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
frmAbout.MousePointer = vbArrow
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

frmAbout.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub



Private Sub Frame3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbBlue
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbBlue
frmAbout.MousePointer = vbArrow
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbBlue
frmAbout.MousePointer = vbArrow
End Sub

Private Sub Label3_Click()
ShellExecute Me.hwnd, "Open", "http://www.morphedmedia.com", vbNullString, vbNullString, 0
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbRed
End Sub

Private Sub lblIntro_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbBlue
frmAbout.MousePointer = vbArrow
End Sub

Private Sub lblLiquid_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbBlue
frmAbout.MousePointer = vbArrow
End Sub

Private Sub lblVer_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbBlue
frmAbout.MousePointer = vbArrow
End Sub

Private Sub lblVersion_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label3.ForeColor = vbBlue
frmAbout.MousePointer = vbArrow
End Sub

Private Sub lblCredits_Click()
Dim i
For i = 0 To 1
Frame3(i).Visible = False
Next i
lblHistory.FontBold = False
Frame3(1).Visible = True
End Sub

Private Sub lblCredits_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblCredits.ForeColor = vbBlue
End Sub

Private Sub lblCredits_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblCredits.FontBold = True
lblCredits.ForeColor = vbBlack
End Sub

Private Sub lblGreenAudio_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblGreenAudio.ForeColor = vbBlue
End Sub

Private Sub lblGreenAudio_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblGreenAudio.FontBold = True
lblGreenAudio.ForeColor = vbBlack
End Sub

Private Sub lblHistory_Click()
Dim i
For i = 0 To 1
Frame3(i).Visible = False
Next i
lblCredits.FontBold = False
Frame3(0).Visible = True
End Sub

Private Sub lblHistory_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblHistory.ForeColor = vbBlue
End Sub

Private Sub lblHistory_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
lblHistory.FontBold = True
lblHistory.ForeColor = vbBlack
End Sub

